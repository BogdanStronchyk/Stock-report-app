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
from checklist_loader import load_thresholds_from_excel
from input_resolver import resolve_to_ticker
from metrics import compute_metrics_v2
from reversal import trend_reversal_scores_from_data
from report_writer import create_report_workbook
from ui_dialogs import ask_output_directory
from ui_stock_picker import ask_stocks
from ui_progress import ProgressWindow, success_popup


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


def _write_errors_workbook(out_path: str, tickers: list, errors: dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Errors"
    ws.append(["Ticker", "Error"])
    for t in tickers:
        ws.append([t, errors.get(t, "")[:30000]])
    wb.save(out_path)


def _analyze_one(sym: str, use_fmp: bool, fmp_mode: str):
    try:
        metrics = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode)

        # Pass the FULL data bundle to reversal script so it doesn't say "Missing Data"
        bundle = metrics.get("__yf_bundle__", {})

        rev = trend_reversal_scores_from_data(
            q_income=bundle.get("q_income"),
            q_cf=bundle.get("q_cf"),
            annual_bs=bundle.get("annual_bs"),
            annual_income=bundle.get("annual_income"),
            annual_cf=bundle.get("annual_cf"),
            h_1y=bundle.get("h1y"),
            h_2y=bundle.get("h2y"),
            metrics=metrics
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

    print(f"Loading checklist from: {checklist_path}")
    thresholds = load_thresholds_from_excel(checklist_path)

    # --- FIX: Correctly unpack the Tuple from the UI ---
    picker_result = ask_stocks()
    if not picker_result: return

    raw_text, user_wants_fmp = picker_result

    raw_tokens = [s.strip() for s in re.split(r'[,\n\s]+', raw_text) if s.strip()]
    if not raw_tokens:
        print("No tickers entered.")
        return

    print(f"Resolving {len(raw_tokens)} symbols...")
    tickers = []
    for tok in raw_tokens:
        sym = resolve_to_ticker(tok)
        if sym: tickers.append(sym)

    tickers = sorted(list(set(tickers)))
    if not tickers: return

    out_dir = ask_output_directory()
    if not out_dir: return

    env_key = os.environ.get("FMP_API_KEY")
    use_fmp = bool(env_key) and (user_wants_fmp or bool(os.environ.get("FORCE_FMP_FALLBACK")))
    fmp_mode = os.environ.get("FMP_MODE", "full")

    print(f"Analyzing {len(tickers)} tickers...")

    metrics_map = {}
    reversal_map = {}
    errors_map = {}

    pwin = ProgressWindow(len(tickers) + 1, title="Analyzing Stocks...")

    start_t = perf_counter()
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_analyze_one, t, use_fmp, fmp_mode): t for t in tickers}

        for fut in as_completed(futures):
            t_sym, m_data, r_data, err = fut.result()
            if err:
                errors_map[t_sym] = err
                print(f"X {t_sym}")
            else:
                metrics_map[t_sym] = m_data
                reversal_map[t_sym] = r_data
                print(f"OK {t_sym}")

            pwin.step(sub_text=f"Analyzed {t_sym}")

    duration = perf_counter() - start_t
    print(f"Analysis done in {_fmt_seconds(duration)}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized = "_".join(tickers[:5])
    if len(tickers) > 5: sanitized += "_etc"

    out_file = os.path.join(out_dir, f"{sanitized}_{timestamp}.xlsx")

    pwin.step(main_text="Writing Report...", sub_text="Formatting Excel...")

    create_report_workbook(
        tickers=[t for t in tickers if t in metrics_map],
        thresholds=thresholds,
        metrics_by_ticker=metrics_map,
        reversal_by_ticker=reversal_map,
        out_path=out_file
    )

    pwin.close()

    if errors_map:
        err_file = os.path.join(out_dir, f"{sanitized}_{timestamp}_errors.xlsx")
        _write_errors_workbook(err_file, tickers, errors_map)
        print(f"Errors written to {err_file}")

    success_popup(out_file)


if __name__ == "__main__":
    main()