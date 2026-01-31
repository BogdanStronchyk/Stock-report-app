try:
    from env_loader import load_env

    load_env()
except Exception:
    pass

import os
import sys
import re
import traceback
from datetime import datetime
from time import perf_counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from openpyxl import Workbook

from config import CHECKLIST_FILE
from checklist_loader import load_thresholds_from_excel, get_threshold_set
from input_resolver import resolve_to_ticker
from metrics import compute_metrics_v2
from reversal import trend_reversal_scores_from_data
from report_writer import create_report_workbook, _normalize_reversal_pack, _metric_weight
from scoring import score_with_threshold_txt, compute_category_score_and_coverage, adjusted_from_raw_and_coverage
from ui_progress import ProgressWindow, success_popup
from ui_dialogs import ask_output_directory
from ui_stock_picker import ask_stocks

# CORE 12 WHITELIST used for filtering logic to match report_writer.py
WHITELIST = {
    "P/E (TTM, positive EPS)", "EV/EBIT", "FCF Yield (TTM FCF / Market Cap)",
    "Gross Margin %", "Operating Margin %", "ROIC % (standardized)",
    "Net Debt / EBITDA", "Interest Coverage (EBIT / Interest)",
    "Revenue per Share CAGR (5Y)", "FCF per Share CAGR (5Y)",
    "Market Cap", "Max Drawdown (3–5Y)"
}


def _fmt_seconds(sec: float) -> str:
    if sec < 60: return f"{sec:.1f}s"
    return f"{int(sec // 60)}m {sec % 60:.0f}s"


def _resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.getcwd()))
    return os.path.join(base, relative_path)


def _find_checklist_file() -> str:
    candidates = [
        os.path.join(os.getcwd(), "Checklist", CHECKLIST_FILE),
        _resource_path(os.path.join("Checklist", CHECKLIST_FILE)),
        _resource_path(CHECKLIST_FILE),
    ]
    for p in candidates:
        if os.path.exists(p): return p
    return candidates[0]


def _get_individual_category_scores(metrics: dict, thresholds: dict) -> dict:
    """Calculates separate scores for Valuation, Quality, Safety, Growth, and Risk."""
    category_maps = {
        "Valuation": "Valuation",
        "Profitability": "Quality",
        "Balance Sheet": "Safety",
        "Growth": "Growth",
        "Risk": "Risk"
    }
    bucket = metrics.get("Sector Bucket", "Default (All)")
    results = {}

    for cat_sheet, display_name in category_maps.items():
        ratings, weights = {}, {}
        th_set = thresholds.get(cat_sheet, {})
        for metric in th_set:
            if metric not in WHITELIST: continue
            val = metrics.get(metric)
            th = get_threshold_set(thresholds, cat_sheet, metric, bucket)
            rating = "NA"
            if th and val is not None:
                rating, _ = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
            ratings[metric] = rating
            weights[metric] = _metric_weight(metric)

        raw, cov = compute_category_score_and_coverage(ratings, weights)
        results[display_name] = adjusted_from_raw_and_coverage(raw, cov)
    return results


def _analyze_one(sym: str, use_fmp: bool, fmp_mode: str):
    try:
        metrics = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode)
        bundle = metrics.get("__yf_bundle__", {})
        rev = trend_reversal_scores_from_data(
            q_income=bundle.get("q_income"), q_cf=bundle.get("q_cf"),
            annual_bs=bundle.get("annual_bs"), annual_income=bundle.get("annual_income"),
            annual_cf=bundle.get("annual_cf"), h_1y=bundle.get("h1y"),
            h_2y=bundle.get("h2y"), metrics=metrics
        )
        return (sym, metrics, rev, None)
    except Exception:
        return (sym, None, None, traceback.format_exc())


def main():
    print("Stock Report App - Starting...")
    checklist_path = _find_checklist_file()
    if not os.path.exists(checklist_path):
        print(f"ERROR: Checklist file not found at {checklist_path}")
        return
    thresholds = load_thresholds_from_excel(checklist_path)

    # 1. Handle selection of stocks and multiple indices from UI
    picker_result = ask_stocks()
    if not picker_result: return
    raw_text, selected_indices, user_wants_fmp = picker_result

    # 2. Collect and deduplicate tickers from all selected sources
    all_raw_tokens = [s.strip() for s in re.split(r'[,\n\s]+', raw_text) if s.strip()]
    universe_dir = _resource_path("Ticker universe")
    for idx_name in selected_indices:
        csv_path = os.path.join(universe_dir, f"{idx_name}.csv")
        if os.path.exists(csv_path):
            with open(csv_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                all_raw_tokens.extend([t.strip() for t in content.split(',') if t.strip()])

    print(f"Resolving {len(all_raw_tokens)} potential symbols...")
    tickers = sorted(list(set([resolve_to_ticker(tok) for tok in all_raw_tokens if resolve_to_ticker(tok)])))
    if not tickers: return

    out_dir = ask_output_directory()
    if not out_dir: return

    metrics_map, reversal_map = {}, {}
    pwin = ProgressWindow(len(tickers) + 1, title="Analyzing Stocks...")

    # 3. Parallel analysis
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_analyze_one, t, user_wants_fmp, "full"): t for t in tickers}
        for fut in as_completed(futures):
            t_sym, m_data, r_data, err = fut.result()
            if not err:
                metrics_map[t_sym], reversal_map[t_sym] = m_data, r_data
            pwin.step(sub_text=f"Processed {t_sym}")

    # 4. Truncate list to STRONG BUY only (Each fundamental category >= 60% AND reversal >= 60%)
    strong_buy_list = []
    for t in tickers:
        if t in metrics_map and t in reversal_map:
            cat_scores = _get_individual_category_scores(metrics_map[t], thresholds)
            rev_pack = _normalize_reversal_pack(reversal_map[t])
            r_score = rev_pack.get("total_score_pct") or 0.0

            # Strict logic: No individual category score can be less than 60%
            if all(s is not None and s >= 60.0 for s in cat_scores.values()) and r_score >= 60.0:
                strong_buy_list.append(t)

    if not strong_buy_list:
        pwin.close()
        print("No Strong Buy candidates found meeting all category thresholds.")
        return

    # 5. Generate final report for the truncated list only
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(out_dir, f"Strong_Buys_Only_{timestamp}.xlsx")
    pwin.step(main_text="Writing Report...", sub_text="Formatting Strong Buys...")
    create_report_workbook(strong_buy_list, thresholds, metrics_map, reversal_map, out_file)

    pwin.close()
    success_popup(out_file)


if __name__ == "__main__":
    main()