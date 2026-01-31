try:
    from env_loader import load_env

    load_env()
except Exception:
    pass

import os, sys, re, traceback, warnings, logging
from datetime import datetime
from time import perf_counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from checklist_loader import load_thresholds_from_excel
from input_resolver import resolve_to_ticker
from metrics import compute_metrics_v2
from reversal import trend_reversal_scores_from_data
from report_writer import create_report_workbook, _normalize_reversal_pack
from ui_progress import ProgressWindow, success_popup
from ui_stock_picker import ask_stocks

# COMPREHENSIVE WARNING SUPPRESSION
warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*Timestamp.utcnow.*")
warnings.filterwarnings("ignore", message=".*Pandas4Warning.*")
warnings.filterwarnings("ignore", message=".*possibly delisted.*")
warnings.filterwarnings("ignore", category=UserWarning)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)


def _resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.getcwd()))
    return os.path.join(base, relative_path)


def _find_checklist_file() -> str:
    p = os.path.join(os.getcwd(), "Checklist", "Fundamental_Checklist_v3_value_matrix_fixed.xlsx")
    return p if os.path.exists(p) else _resource_path("Fundamental_Checklist_v3_value_matrix_fixed.xlsx")


def _get_individual_category_scores(metrics: dict, thresholds: dict) -> dict:
    from checklist_loader import get_threshold_set
    from scoring import score_with_threshold_txt, compute_category_score_and_coverage, adjusted_from_raw_and_coverage
    from report_writer import WHITELIST, _metric_weight
    category_maps = {"Valuation": "Valuation", "Profitability": "Quality", "Balance Sheet": "Safety",
                     "Growth": "Growth", "Risk": "Risk"}
    bucket = metrics.get("Sector Bucket", "Default (All)");
    results = {}
    for cat_sheet, display_name in category_maps.items():
        ratings, weights = {}, {}
        for metric in thresholds.get(cat_sheet, {}):
            if metric not in WHITELIST: continue
            val = metrics.get(metric);
            th = get_threshold_set(thresholds, cat_sheet, metric, bucket)
            rating = "NA"
            if th and val is not None:
                rating, _ = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
            ratings[metric], weights[metric] = rating, _metric_weight(metric)
        raw, cov = compute_category_score_and_coverage(ratings, weights)
        results[display_name] = adjusted_from_raw_and_coverage(raw, cov)
    return results


def _analyze_one(sym: str, use_fmp: bool, fmp_mode: str):
    try:
        m = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode);
        b = m.get("__yf_bundle__", {})
        rev = trend_reversal_scores_from_data(q_income=b.get("q_income"), q_cf=b.get("q_cf"),
                                              annual_bs=b.get("annual_bs"), annual_income=b.get("annual_income"),
                                              annual_cf=b.get("annual_cf"), h_1y=b.get("h1y"), h_2y=b.get("h2y"),
                                              metrics=m)
        return (sym, m, rev, None)
    except Exception:
        return (sym, None, None, traceback.format_exc())


# ... (imports and helper functions remain the same)

# ... (imports and helper functions remain the same)

def main():
    # 1. Setup default directory
    out_dir = os.path.join(os.getcwd(), "reports")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    checklist_path = _find_checklist_file()
    thresholds = load_thresholds_from_excel(checklist_path)

    picker_result = ask_stocks()
    if not picker_result: return
    # picker_result now returns (text, indices, rule_mode, use_fmp, include_watch)
    raw_text, indices, rule_mode, use_fmp, include_watch = picker_result

    target_threshold = {"Strict": 60.0, "Moderate": 50.0, "Loose": 40.0}.get(rule_mode, 60.0)

    pwin = ProgressWindow(100, title=f"Scraping ({rule_mode} Mode)...")
    pwin.step(main_text="Preparing data...", sub_text="Resolving symbols...")

    all_raw = [s.strip() for s in re.split(r'[,\n\s]+', raw_text) if s.strip()]
    universe_dir = _resource_path("Ticker universe")
    for idx_name in indices:
        csv_path = os.path.join(universe_dir, f"{idx_name}.csv")
        if os.path.exists(csv_path):
            with open(csv_path, 'r') as f:
                all_raw.extend([t.strip() for t in f.read().split(',') if t.strip()])

    tickers = sorted(list(set([resolve_to_ticker(tok) for tok in all_raw if resolve_to_ticker(tok)])))
    if not tickers:
        pwin.close()
        return

    pwin.max_steps = len(tickers) + 2
    pwin.current_step = 0

    metrics_map, reversal_map = {}, {}
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_analyze_one, t, use_fmp, "full"): t for t in tickers}
        for fut in as_completed(futures):
            t_sym, m_data, r_data, err = fut.result()
            if not err:
                metrics_map[t_sym], reversal_map[t_sym] = m_data, r_data
            pwin.step(sub_text=f"Processed: {t_sym}")

    # --- UPDATED TRUNCATION LOGIC ---
    pwin.step(main_text="Filtering results...", sub_text="Selecting Strong Buy/Watch candidates...")
    final_filtered_list = []

    for t in tickers:
        if t in metrics_map and t in reversal_map:
            cat_scores = _get_individual_category_scores(metrics_map[t], thresholds)
            rev_pack = _normalize_reversal_pack(reversal_map[t])
            r_score = rev_pack.get("total_score_pct") or 0.0

            # Check fundamental pass (All 5 categories >= threshold)
            valid_scores = [s for s in cat_scores.values() if s is not None]
            all_cats_pass = len(valid_scores) == 5 and all(s >= target_threshold for s in valid_scores)

            # 1. STRONG BUY: Fundamentals pass AND Reversal passes
            is_strong_buy = all_cats_pass and r_score >= target_threshold

            # 2. WATCH: Fundamentals pass BUT Reversal is low
            is_watch = all_cats_pass and r_score < target_threshold

            if is_strong_buy:
                final_filtered_list.append(t)
            elif include_watch and is_watch:
                final_filtered_list.append(t)

    if not final_filtered_list:
        pwin.close()
        print(f"No candidates met the {rule_mode} criteria.")
        return

    # Generate filename and report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(out_dir, f"Filtered_Report_{rule_mode}_{timestamp}.xlsx")

    create_report_workbook(final_filtered_list, thresholds, metrics_map, reversal_map, out_file, target_threshold)

    pwin.close()
    success_popup(out_file)


if __name__ == "__main__":
    main()