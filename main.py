# Auto-load environment variables
try:
    from env_loader import load_env

    load_env()
except Exception:
    pass

import os
import sys
import csv
import warnings
import traceback
import logging
from datetime import datetime, timedelta
from time import perf_counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from openpyxl import Workbook

# --- CACHING SETUP (CRITICAL FOR LOW QUOTA) ---
try:
    import requests_cache

    requests_cache.install_cache(
        'fmp_cache',
        backend='sqlite',
        expire_after=timedelta(hours=24),
        allowable_codes=(200, 404, 400),  # Cache Errors too!
        allowable_methods=('GET', 'POST')
    )
    print(f"[SYSTEM] Global API Cache Installed: 'fmp_cache.sqlite'")
except ImportError:
    print("[WARN] 'requests_cache' not found. Run: pip install requests-cache")
# ----------------------------------------------

from config import CHECKLIST_FILE
from checklist_loader import load_thresholds_from_excel
from input_resolver import resolve_to_ticker
from metrics import compute_metrics_v2
from reversal import trend_reversal_scores_from_data
from report_writer import create_report_workbook
from ui_stock_picker import ask_stocks
from ui_progress import ProgressWindow, success_popup
from scoring import score_ticker
from eligibility import evaluate_eligibility
from fmp_provider import FMPClient


# ---------------------------------------------------------
# SMART NORMALIZATION
# ---------------------------------------------------------
def normalize_ticker_for_fmp(ticker: str) -> str | None:
    if not ticker or not isinstance(ticker, str):
        return None

    clean_ticker = ticker.strip().upper()

    if '.' in clean_ticker:
        parts = clean_ticker.split('.')
        suffix = parts[-1]
        # US Share Classes (BRK.B) -> Hyphen
        if len(suffix) == 1 and suffix in ['A', 'B', 'C']:
            return clean_ticker.replace('.', '-')
        # International (DSY.PA) -> Keep Dot
        return clean_ticker

    return clean_ticker


def _fmt_seconds(sec: float) -> str:
    sec = max(0.0, float(sec))
    if sec < 60: return f"{sec:.1f}s"
    m = int(sec // 60)
    s = sec - m * 60
    if m < 60: return f"{m}m {s:0.0f}s"
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
        if os.path.exists(p): return p
    return candidates[0]


def _metrics_looks_empty(m: dict) -> bool:
    if not isinstance(m, dict) or not m: return True
    nontrivial = [k for k in m.keys() if
                  k not in ("Ticker", "__notes__", "__yf_bundle__", "__scoring__", "__eligibility__")]
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
    picked = ask_stocks()
    if not picked:
        print("Canceled.")
        return

    prog = ProgressWindow(total_steps=100, title="Initializing...")
    prog.set_status("Loading configuration...")

    try:
        use_fmp = bool(picked.get("use_fmp", False))
        mode = picked.get("mode", "manual")
        target_eligibility_mode = picked.get("eligibility_mode", "strict")
        fmp_mode = picked.get("fmp_mode", "minimal") if use_fmp else "off"

        prog.set_status("Loading checklist...", "Reading Excel thresholds")
        checklist_path = _find_checklist_file()
        try:
            thresholds = load_thresholds_from_excel(checklist_path)
        except Exception as e:
            prog.close()
            print(f"Error loading checklist from {checklist_path}: {e}")
            return

        prog.set_status("Reading input lists...", f"Mode: {mode}")
        raw_tickers_input = []

        if mode == "universe":
            uni_paths = picked.get("universe_paths", [])
            for path in uni_paths:
                if not os.path.exists(path): continue
                try:
                    with open(path, "r", encoding="utf-8-sig", newline="") as f:
                        reader = csv.reader(f)
                        for row in reader:
                            if not row: continue
                            v = (row[0] or "").strip()
                            if v and v.lower() not in ("ticker", "symbol"):
                                raw_tickers_input.append(v)
                except Exception:
                    pass
        else:
            raw_str = picked.get("raw") or ""
            raw_str = raw_str.replace("\n", ",").replace("\r", ",")
            raw_tickers_input = [p.strip() for p in raw_str.split(",") if p.strip()]

        prog.set_status("Normalizing tickers...", "Formatting for FMP")
        normalized_tickers = []
        seen = set()

        for t in raw_tickers_input:
            clean = normalize_ticker_for_fmp(t)
            if clean and clean not in seen:
                normalized_tickers.append(clean)
                seen.add(clean)

        raw_tickers = normalized_tickers

        if not raw_tickers:
            prog.close()
            print("No valid tickers found after normalization.")
            return

        resolved = []
        prog.pb["maximum"] = len(raw_tickers)
        prog.total_steps = len(raw_tickers)
        prog.current = 0

        prog.set_status(f"Resolving {len(raw_tickers)} tickers...", "Validating existence...")

        with ThreadPoolExecutor(max_workers=8) as ex:
            def _resolve(t):
                return resolve_to_ticker(t)

            futs = {ex.submit(_resolve, t): t for t in raw_tickers}
            for i, fut in enumerate(as_completed(futs)):
                try:
                    sym = fut.result()
                    if sym: resolved.append(sym)
                except Exception:
                    pass
                prog.step(main_text=f"Resolving ({i + 1}/{len(raw_tickers)})", sub_text=f"Found: {len(resolved)} valid")

        tickers = sorted(list(set(resolved)))

        if not tickers:
            prog.close()
            print("No valid tickers found after resolution.")
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = "MULTI_UNIVERSE" if mode == "universe" else "MANUAL_PORTFOLIO"
        if mode == "universe" and len(picked.get("universe_paths", [])) == 1:
            base_name = os.path.splitext(os.path.basename(picked["universe_paths"][0]))[0].upper()

        filename = f"{base_name}_{target_eligibility_mode.upper()}_{ts}.xlsx"
        out_dir = os.path.join(os.getcwd(), "reports")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, filename)

        steps_total = len(tickers) + 1
        prog.pb["maximum"] = steps_total
        prog.total_steps = steps_total
        prog.current = 0
        prog.pb["value"] = 0
        prog.set_status(f"Analyzing {len(tickers)} candidates...", f"Mode: {target_eligibility_mode.upper()}")

        metrics_by_ticker: dict[str, dict] = {}
        reversal_by_ticker: dict[str, dict] = {}
        errors: dict[str, str] = {}
        failed_tickers: set[str] = set()
        filtered_tickers: int = 0

        max_workers = int(float(os.environ.get("YF_MAX_WORKERS", "6") or "6"))
        fetch_durations = []

        def _process_ticker(sym: str):
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

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futs = {ex.submit(_process_ticker, t): t for t in tickers}
            done_n = 0

            for fut in as_completed(futs):
                sym = futs[fut]
                done_n += 1
                try:
                    t_sym, m_sym, r_sym, dt = fut.result()

                    # --- SILENT DROP ALERT ---
                    if _metrics_looks_empty(m_sym):
                        failed_tickers.add(t_sym)
                        print(f"[ALERT] Dropped {t_sym}: FMP data incomplete (Circuit Breaker hit).")
                        prog.step(
                            main_text=f"Scanning ({done_n}/{len(tickers)})",
                            sub_text=f"Dropped: {t_sym}",
                            done_text=f"REASON: Data Retrieval Failed"
                        )
                        continue
                    # -------------------------

                    scores = score_ticker(m_sym, thresholds)
                    elig = evaluate_eligibility(
                        mode=target_eligibility_mode,
                        cat_adj=scores["cat_adj"],
                        cat_cov=scores["cat_cov"],
                        category_ratings=scores["ratings"],
                        sector_bucket=m_sym.get("Sector Bucket"),
                        fund_adj=scores["fund_adj"],
                        reversal_total=r_sym.get("total_score_pct"),
                        davf_label=m_sym.get("DAVF Downside Protection")
                    )

                    if elig.status == "FAIL":
                        filtered_tickers += 1
                        prog.step(
                            main_text=f"Scanning ({done_n}/{len(tickers)})",
                            sub_text=f"Filtered: {t_sym}",
                            done_text=f"Dropped {t_sym}: {elig.reasons_text(1)}"
                        )
                        continue

                    m_sym["__scoring__"] = scores
                    m_sym["__eligibility__"] = elig
                    m_sym["Decision Status"] = elig.status
                    m_sym["Decision Label"] = elig.label
                    m_sym["Eligibility Notes"] = elig.reasons_text()

                    metrics_by_ticker[t_sym] = m_sym
                    reversal_by_ticker[t_sym] = r_sym

                    fetch_durations.append(dt)
                    avg = sum(fetch_durations) / max(1, len(fetch_durations))

                    fmp_ok, fmp_err = FMPClient.get_stats()
                    fmp_status = f" | FMP: {fmp_ok} OK, {fmp_err} Err" if use_fmp else ""

                    prog.step(
                        main_text=f"Candidates found: {len(metrics_by_ticker)}",
                        sub_text=f"Added: {t_sym} ({elig.label})",
                        done_text=f"Avg: {_fmt_seconds(avg)}{fmp_status}"
                    )

                except Exception:
                    tb = traceback.format_exc()
                    errors[sym] = tb
                    failed_tickers.add(sym)

        if not metrics_by_ticker:
            prog.step(main_text="Finished", sub_text="No candidates found.")
            if errors:
                err_path = out_path.replace(".xlsx", "_ERRORS.xlsx")
                _write_errors_workbook(err_path, tickers, errors)
                success_popup(err_path)
            else:
                success_popup(f"0 candidates remained. {filtered_tickers} filtered.")
            return

        prog.step(main_text="Generating Report...", sub_text=f"Writing {len(metrics_by_ticker)} candidates")
        create_report_workbook(
            tickers=list(metrics_by_ticker.keys()),
            thresholds=thresholds,
            metrics_by_ticker=metrics_by_ticker,
            reversal_by_ticker=reversal_by_ticker,
            out_path=out_path
        )
        prog.step(main_text="Done!", sub_text=out_path)

    finally:
        prog.close()

    print(f"✅ DONE. Report saved: {out_path}")
    success_popup(out_path)


if __name__ == "__main__":
    main()