try:
    from env_loader import load_env
    load_env()
except Exception: pass
import os, sys, re, traceback, warnings
from datetime import datetime
from time import perf_counter
from concurrent.futures import ThreadPoolExecutor, as_completed
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

WHITELIST = {
    "P/E (TTM, positive EPS)", "EV/EBIT", "FCF Yield (TTM FCF / Market Cap)",
    "Gross Margin %", "Operating Margin %", "ROIC % (standardized)",
    "Net Debt / EBITDA", "Interest Coverage (EBIT / Interest)",
    "Revenue per Share CAGR (5Y)", "FCF per Share CAGR (5Y)",
    "Market Cap", "Max Drawdown (3–5Y)"
}

def _resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.getcwd()))
    return os.path.join(base, relative_path)

def _find_checklist_file() -> str:
    p = os.path.join(os.getcwd(), "Checklist", CHECKLIST_FILE)
    return p if os.path.exists(p) else _resource_path(CHECKLIST_FILE)

def _get_individual_category_scores(metrics: dict, thresholds: dict) -> dict:
    category_maps = {"Valuation": "Valuation", "Profitability": "Quality", "Balance Sheet": "Safety", "Growth": "Growth", "Risk": "Risk"}
    bucket = metrics.get("Sector Bucket", "Default (All)"); results = {}
    for cat_sheet, display_name in category_maps.items():
        ratings, weights = {}, {}
        for metric in thresholds.get(cat_sheet, {}):
            if metric not in WHITELIST: continue
            val = metrics.get(metric); th = get_threshold_set(thresholds, cat_sheet, metric, bucket)
            rating = "NA"
            if th and val is not None:
                rating, _ = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
            ratings[metric], weights[metric] = rating, _metric_weight(metric)
        raw, cov = compute_category_score_and_coverage(ratings, weights)
        results[display_name] = adjusted_from_raw_and_coverage(raw, cov)
    return results

def _analyze_one(sym: str, use_fmp: bool, fmp_mode: str):
    try:
        m = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode); b = m.get("__yf_bundle__", {})
        rev = trend_reversal_scores_from_data(q_income=b.get("q_income"), q_cf=b.get("q_cf"), annual_bs=b.get("annual_bs"), annual_income=b.get("annual_income"), annual_cf=b.get("annual_cf"), h_1y=b.get("h1y"), h_2y=b.get("h2y"), metrics=m)
        return (sym, m, rev, None)
    except Exception: return (sym, None, None, traceback.format_exc())

def main():
    checklist_path = _find_checklist_file()
    thresholds = load_thresholds_from_excel(checklist_path)
    picker_result = ask_stocks()
    if not picker_result: return
    raw_text, indices, rule_mode, use_fmp = picker_result
    target_threshold = {"Strict": 60.0, "Moderate": 50.0, "Loose": 40.0}.get(rule_mode, 60.0)
    all_raw = [s.strip() for s in re.split(r'[,\n\s]+', raw_text) if s.strip()]
    universe_dir = _resource_path("Ticker universe")
    for idx_name in indices:
        csv_path = os.path.join(universe_dir, f"{idx_name}.csv")
        if os.path.exists(csv_path):
            with open(csv_path, 'r') as f: all_raw.extend([t.strip() for t in f.read().split(',') if t.strip()])
    tickers = sorted(list(set([resolve_to_ticker(tok) for tok in all_raw if resolve_to_ticker(tok)])))
    if not tickers: return
    out_dir = ask_output_directory()
    if not out_dir: return
    metrics_map, reversal_map = {}, {}
    pwin = ProgressWindow(len(tickers) + 1, title=f"Analyzing ({rule_mode})...")
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_analyze_one, t, use_fmp, "full"): t for t in tickers}
        for fut in as_completed(futures):
            t_sym, m_data, r_data, err = fut.result()
            if not err: metrics_map[t_sym], reversal_map[t_sym] = m_data, r_data
            pwin.step(sub_text=f"Processed {t_sym}")
    strong_buy_list = []
    for t in tickers:
        if t in metrics_map and t in reversal_map:
            cat_scores = _get_individual_category_scores(metrics_map[t], thresholds)
            r_score = _normalize_reversal_pack(reversal_map[t]).get("total_score_pct") or 0.0
            if all(s is not None and s >= target_threshold for s in cat_scores.values()) and r_score >= target_threshold:
                strong_buy_list.append(t)
    if not strong_buy_list:
        pwin.close(); print(f"No Strong Buys at {int(target_threshold)}%"); return
    out_file = os.path.join(out_dir, f"Strong_Buys_{rule_mode}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    create_report_workbook(strong_buy_list, thresholds, metrics_map, reversal_map, out_file, target_threshold)
    pwin.close(); success_popup(out_file)

if __name__ == "__main__": main()