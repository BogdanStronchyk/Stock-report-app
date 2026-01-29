# Auto-load environment variables from .env (project root) if present.
# IMPORTANT: must run before importing modules that read os.environ (config/metrics/fmp).
try:
    from env_loader import load_env
    load_env()
except Exception:
    pass

import os
import sys
import warnings
import traceback
from datetime import datetime
from time import perf_counter
from concurrent.futures import ThreadPoolExecutor, as_completed

import yfinance as yf
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
    sec = max(0.0, float(sec))
    if sec < 60:
        return f"{sec:.1f}s"
    m = int(sec // 60)
    s = sec - m * 60
    if m < 60:
        return f"{m}m {s:0.0f}s"
    h = int(m // 60)
    m2 = m - h * 60
    return f"{h}h {m2}m"


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


def _metrics_looks_empty(m: dict) -> bool:
    """Heuristic: if we only have identity/notes keys, analysis effectively failed."""
    if not isinstance(m, dict) or not m:
        return True
    nontrivial = [k for k in m.keys() if k not in ("Ticker", "__notes__", "__yf_bundle__")]
    return len(nontrivial) == 0


def _write_errors_workbook(out_path: str, tickers: list[str], errors: dict[str, str]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Errors"
    ws.append(["Ticker", "Error (traceback)"])
    for t in tickers:
        ws.append([t, (errors.get(t) or "")[:30000]])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 140
    wb.save(out_path)


def main():
    warnings.filterwarnings("ignore")

    checklist_path = _find_checklist_file()
    thresholds = load_thresholds_from_excel(checklist_path)

    picked = ask_stocks()
    if not picked:
        print("Canceled.")
        return

    raw, use_fmp = picked

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

    steps_total = len(tickers) + 2
    prog = ProgressWindow(total_steps=steps_total, title="Stock Report App — Generating")

    metrics_by_ticker: dict[str, dict] = {}
    reversal_by_ticker: dict[str, dict] = {}
    errors: dict[str, str] = {}
    failed: set[str] = set()

    try:
        prog.set_status("Fetching data...", f"Tickers: {', '.join(tickers)}")

        fetch_durations = []

        # Stage-based FMP behavior:
        #  - broad/scan      : no FMP
        #  - narrow/shortlist: minimal FMP bundle
        #  - portfolio       : full FMP bundle
        stage = (os.environ.get("ANALYSIS_STAGE", "") or "").strip().lower()
        if stage in ("broad", "scan", "wide"):
            stage_fmp_mode = "off"
        elif stage in ("narrow", "shortlist", "screen"):
            stage_fmp_mode = "minimal"
        else:
            stage_fmp_mode = "full"

        # User toggle can still disable FMP entirely
        fmp_mode = stage_fmp_mode if use_fmp else "off"

        max_workers = int(float(os.environ.get("YF_MAX_WORKERS", "6") or "6"))
        max_workers = max(1, min(32, max_workers))

        def _analyze_one(sym: str):
            t0 = perf_counter()
            metrics = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode)
            bundle = metrics.pop("__yf_bundle__", {}) or {}

            rev = trend_reversal_scores_from_data(
                ticker=sym,
                info=bundle.get("info"),
                q_income=bundle.get("q_income"),
                q_cf=bundle.get("q_cf"),
                annual_bs=bundle.get("annual_bs"),
                h_1y=bundle.get("h1y"),
                h_2y=bundle.get("h2y"),
                metrics=metrics,
            )
            dt = perf_counter() - t0
            return sym, metrics, rev, dt

        prog.set_status("Fetching + scoring (concurrent)...", f"Tickers: {', '.join(tickers)}")

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futs = {ex.submit(_analyze_one, t): t for t in tickers}
            done_n = 0
            for fut in as_completed(futs):
                sym = futs[fut]
                done_n += 1
                try:
                    t_sym, m_sym, r_sym, dt = fut.result()
                except Exception:
                    dt = 0.0
                    tb = traceback.format_exc()
                    errors[sym] = tb
                    failed.add(sym)
                    # Keep a placeholder so the workbook can still include the ticker tab/summary row.
                    t_sym, m_sym, r_sym = sym, {"Ticker": sym, "__notes__": {"Error": "See Errors sheet / error log"}}, {}

                # Mark as failed if output is effectively empty
                if _metrics_looks_empty(m_sym):
                    failed.add(t_sym)
                    if t_sym not in errors:
                        errors[t_sym] = "Metrics returned empty/near-empty dict."

                metrics_by_ticker[t_sym] = m_sym
                reversal_by_ticker[t_sym] = r_sym

                fetch_durations.append(dt)
                avg = sum(fetch_durations) / max(1, len(fetch_durations))
                remaining = max(0, len(tickers) - done_n)
                eta = remaining * avg
                done_txt = (
                    f"Done: {done_n}/{len(tickers)} | ({t_sym}): {_fmt_seconds(dt)} "
                    f"| Avg: {_fmt_seconds(avg)} | ETA: {_fmt_seconds(eta)} "
                    f"| Workers: {max_workers} | FMP: {fmp_mode}"
                )
                if t_sym in failed:
                    done_txt += " | ⚠ failed"
                prog.step(main_text=f"Processed ({done_n}/{len(tickers)})", sub_text=t_sym, done_text=done_txt)

        # If EVERYTHING failed, do not emit a misleading "empty" report.
        if len(failed) == len(tickers):
            prog.step(main_text="All tickers failed", sub_text="Writing an Errors workbook instead…")
            _write_errors_workbook(out_path, tickers, errors)

            # also write a plain-text error log for quick copy/paste
            try:
                log_path = os.path.splitext(out_path)[0] + "_errors.txt"
                with open(log_path, "w", encoding="utf-8") as f:
                    for t in tickers:
                        f.write(f"=== {t} ===\n")
                        f.write(errors.get(t, "") + "\n\n")
                print("⚠ Error log:", log_path)
            except Exception:
                pass

            print("❌ All tickers failed. Saved errors workbook:", out_path)
            success_popup(out_path)
            return

        prog.step(main_text="Writing Excel report...", sub_text="Applying checklist + scores + colors")
        create_report_workbook(
            tickers=tickers,
            thresholds=thresholds,
            metrics_by_ticker=metrics_by_ticker,
            reversal_by_ticker=reversal_by_ticker,
            out_path=out_path
        )
        prog.step(main_text="Done!", sub_text=out_path)

        # Write error log if some tickers failed (so you can quickly see why without hunting in the workbook).
        if failed:
            try:
                log_path = os.path.splitext(out_path)[0] + "_errors.txt"
                with open(log_path, "w", encoding="utf-8") as f:
                    for t in sorted(failed):
                        f.write(f"=== {t} ===\n")
                        f.write(errors.get(t, "") + "\n\n")
                print("⚠ Some tickers failed. Error log:", log_path)
            except Exception:
                pass

    finally:
        prog.close()

    print("✅ DONE:", out_path)
    success_popup(out_path)


if __name__ == "__main__":
    main()
