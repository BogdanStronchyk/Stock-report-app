# Auto-load environment variables from .env (project root) if present.
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
from ui_stock_picker import ask_stocks
from ui_progress import ProgressWindow, success_popup
from scoring import score_ticker
from eligibility import evaluate_eligibility


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
    if not isinstance(m, dict) or not m:
        return True
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

    # UI Picker
    picked = ask_stocks()
    if not picked:
        print("Canceled.")
        return

    # 1. Setup Configuration based on UI
    use_fmp = bool(picked.get("use_fmp", False))
    mode = picked.get("mode", "manual")

    # Extract User Choice for Ruleset (Strict/Permissible/Loose)
    target_eligibility_mode = picked.get("eligibility_mode", "strict")

    fmp_mode = "full" if use_fmp else "off"

    # Load Checklist
    checklist_path = _find_checklist_file()
    try:
        thresholds = load_thresholds_from_excel(checklist_path)
    except Exception as e:
        print(f"Error loading checklist from {checklist_path}: {e}")
        return

    # 2. Build Ticker List
    raw_tickers = []

    if mode == "universe":
        # Handle multiple universes
        uni_paths = picked.get("universe_paths", [])
        for path in uni_paths:
            if not os.path.exists(path):
                print(f"Skipping missing file: {path}")
                continue
            try:
                with open(path, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if not row: continue
                        v = (row[0] or "").strip()
                        if v and v.lower() not in ("ticker", "symbol"):
                            raw_tickers.append(v)
            except Exception as e:
                print(f"Error reading {path}: {e}")
    else:
        # Manual input
        raw_str = picked.get("raw") or ""
        raw_tickers = [p.strip() for p in raw_str.split(",") if p.strip()]

    # Resolve Tickers
    resolved = []
    for p in raw_tickers:
        sym = resolve_to_ticker(p)
        if sym: resolved.append(sym)

    # Deduplicate
    tickers = sorted(list(set(resolved)))

    if not tickers:
        print("No valid tickers found to process.")
        return

    # 3. Output Setup (Auto-Save to 'reports' folder)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if mode == "universe":
        if len(picked.get("universe_paths", [])) == 1:
            base_name = os.path.splitext(os.path.basename(picked["universe_paths"][0]))[0].upper()
        else:
            base_name = "MULTI_UNIVERSE"
    else:
        base_name = "MANUAL_PORTFOLIO"

    filename = f"{base_name}_{target_eligibility_mode.upper()}_{ts}.xlsx"

    # --- CHANGE: Hardcoded output directory ---
    out_dir = os.path.join(os.getcwd(), "reports")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)

    # 4. Processing Loop
    steps_total = len(tickers) + 1
    prog = ProgressWindow(total_steps=steps_total, title=f"Generator ({target_eligibility_mode.title()})")

    metrics_by_ticker: dict[str, dict] = {}
    reversal_by_ticker: dict[str, dict] = {}
    errors: dict[str, str] = {}
    failed_tickers: set[str] = set()  # Technical failures
    filtered_tickers: int = 0  # Eligibility failures

    try:
        prog.set_status(f"Analyzing {len(tickers)} candidates...",
                        f"Mode: {target_eligibility_mode.upper()} | FMP: {fmp_mode}")

        max_workers = int(float(os.environ.get("YF_MAX_WORKERS", "6") or "6"))
        fetch_durations = []

        def _process_ticker(sym: str):
            t0 = perf_counter()
            # Fetch Data (Yahoo + Optional FMP)
            metrics = compute_metrics_v2(sym, use_fmp_fallback=use_fmp, fmp_mode=fmp_mode)

            # Reversal Scoring
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

                    # Check for data emptiness
                    if _metrics_looks_empty(m_sym):
                        failed_tickers.add(t_sym)
                        continue

                    # --- SCORING & GATING ---
                    # 1. Score
                    scores = score_ticker(m_sym, thresholds)

                    # 2. Check Eligibility (using user-selected mode)
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

                    # 3. Filter: If it fails criteria, DROP IT.
                    if elig.status == "FAIL":
                        filtered_tickers += 1
                        prog.step(
                            main_text=f"Scanning ({done_n}/{len(tickers)})",
                            sub_text=f"Filtered: {t_sym}",
                            done_text=f"Dropped {t_sym}: {elig.reasons_text(1)}"
                        )
                        continue

                    # 4. Success: Store data
                    m_sym["__scoring__"] = scores
                    m_sym["__eligibility__"] = elig
                    m_sym["Decision Status"] = elig.status
                    m_sym["Decision Label"] = elig.label
                    m_sym["Eligibility Notes"] = elig.reasons_text()

                    metrics_by_ticker[t_sym] = m_sym
                    reversal_by_ticker[t_sym] = r_sym

                    fetch_durations.append(dt)
                    avg = sum(fetch_durations) / max(1, len(fetch_durations))

                    prog.step(
                        main_text=f"Candidates found: {len(metrics_by_ticker)}",
                        sub_text=f"Added: {t_sym} ({elig.label})",
                        done_text=f"Kept {t_sym} | Avg Time: {_fmt_seconds(avg)}"
                    )

                except Exception:
                    tb = traceback.format_exc()
                    errors[sym] = tb
                    failed_tickers.add(sym)

        # 5. Final Report Generation
        if not metrics_by_ticker:
            prog.step(main_text="Finished", sub_text="No candidates found.")
            if errors:
                _write_errors_workbook(out_path.replace(".xlsx", "_ERRORS.xlsx"), tickers, errors)
                print(f"No candidates met criteria. {filtered_tickers} filtered, {len(failed_tickers)} failed.")
            else:
                message_text = f"Analyzed {len(tickers)} tickers.\n{filtered_tickers} filtered out ({target_eligibility_mode}).\n0 candidates remained."
                success_popup(message_text)
            return

        prog.step(main_text="Generating Report...", sub_text=f"Writing {len(metrics_by_ticker)} candidates to Excel")

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