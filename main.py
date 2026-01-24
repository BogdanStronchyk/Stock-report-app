import os
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


def main():
    warnings.filterwarnings("ignore")

    checklist_path = os.path.join(os.getcwd()+r"\Checklist", CHECKLIST_FILE)
    thresholds = load_thresholds_from_excel(checklist_path)

    print("\n=== Stock Report Generator v2 (Modular) ===")
    print("Enter ticker symbols or company names, comma separated.")
    raw = input("Input: ").strip()
    if not raw:
        print("No input provided.")
        return

    parts = [p.strip() for p in raw.split(",") if p.strip()]
    resolved, unresolved = [], []

    for p in parts:
        sym = resolve_to_ticker(p)
        if sym:
            resolved.append(sym)
        else:
            unresolved.append(p)

    tickers = []
    seen = set()
    for t in resolved:
        if t not in seen:
            seen.add(t)
            tickers.append(t)

    if not tickers:
        print("Could not resolve any tickers.")
        return

    if unresolved:
        print("\nSkipped unresolved:")
        print(", ".join(unresolved))

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = "_".join(tickers)
    filename = f"{base}_{ts}.xlsx"

    # ✅ UI folder picker
    out_dir = ask_output_directory(default_subfolder="reports")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)

    print("\nFetching metrics + building report...")

    metrics_by_ticker = {}
    reversal_by_ticker = {}

    for t in tickers:
        metrics_by_ticker[t] = compute_metrics_v2(t)
        reversal_by_ticker[t] = trend_reversal_scores(yf.Ticker(t))

    create_report_workbook(
        tickers=tickers,
        thresholds=thresholds,
        metrics_by_ticker=metrics_by_ticker,
        reversal_by_ticker=reversal_by_ticker,
        out_path=out_path
    )

    print("\n✅ DONE:", out_path)


if __name__ == "__main__":
    main()
