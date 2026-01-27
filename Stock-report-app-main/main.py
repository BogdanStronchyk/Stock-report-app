
import os
import sys
import warnings
from datetime import datetime

import yfinance as yf

from config import CHECKLIST_FILE
from checklist_loader import load_thresholds_from_excel
from input_resolver import resolve_to_ticker
from metrics import compute_metrics_v2
from reversal import trend_reversal_scores
from report_writer import create_report_workbook
from ui_dialogs import ask_output_directory
from ui_stock_picker import ask_stocks
from ui_progress import ProgressWindow, success_popup


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
        if os.path.exists(p):
            return p
    return candidates[0]


def main():
    warnings.filterwarnings("ignore")

    checklist_path = _find_checklist_file()
    thresholds = load_thresholds_from_excel(checklist_path)

    raw = ask_stocks()
    if not raw:
        print("Canceled.")
        return

    parts = [p.strip() for p in raw.split(",") if p.strip()]
    resolved, unresolved = [], []

    for p in parts:
        sym = resolve_to_ticker(p)
        if sym:
            resolved.append(sym)
        else:
            unresolved.append(p)

    tickers, seen = [], set()
    for t in resolved:
        if t not in seen:
            seen.add(t)
            tickers.append(t)

    if not tickers:
        print("Could not resolve any tickers.")
        if unresolved:
            print("Unresolved:", ", ".join(unresolved))
        return

    if unresolved:
        print("Skipped unresolved:", ", ".join(unresolved))

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = "_".join(tickers)
    filename = f"{base}_{ts}.xlsx"

    out_dir = ask_output_directory(default_subfolder="reports")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)

    steps_total = len(tickers) * 2 + 2
    prog = ProgressWindow(total_steps=steps_total, title="Stock Report App — Generating")

    metrics_by_ticker = {}
    reversal_by_ticker = {}

    try:
        prog.set_status("Fetching data...", f"Tickers: {', '.join(tickers)}")

        for i, t in enumerate(tickers, start=1):
            prog.step(main_text=f"Fetching fundamentals ({i}/{len(tickers)})", sub_text=t)
            metrics_by_ticker[t] = compute_metrics_v2(t)

            prog.step(main_text=f"Reversal scoring ({i}/{len(tickers)})", sub_text=t)
            reversal_by_ticker[t] = trend_reversal_scores(yf.Ticker(t), metrics_by_ticker[t])

        prog.step(main_text="Writing Excel report...", sub_text="Applying checklist + scores + colors")
        create_report_workbook(
            tickers=tickers,
            thresholds=thresholds,
            metrics_by_ticker=metrics_by_ticker,
            reversal_by_ticker=reversal_by_ticker,
            out_path=out_path
        )
        prog.step(main_text="Done!", sub_text=out_path)

    finally:
        prog.close()

    print("✅ DONE:", out_path)
    success_popup(out_path)


if __name__ == "__main__":
    main()
