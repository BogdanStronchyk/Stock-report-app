import os
import glob
import tkinter as tk

from config import FORCE_FMP_FALLBACK
from tkinter import messagebox
from typing import Optional, Dict, Any


def _discover_universe_csvs() -> Dict[str, str]:
    """Find universe CSVs in common locations.

    Returns mapping: display_name -> filepath
    """
    candidates: Dict[str, str] = {}

    # 1) ./universes/*.csv
    for p in glob.glob(os.path.join(os.getcwd(), "universes", "*.csv")):
        name = os.path.splitext(os.path.basename(p))[0]
        candidates[name.upper()] = p

    # 2) project root *.csv (for quick drop-in)
    for p in glob.glob(os.path.join(os.getcwd(), "*.csv")):
        name = os.path.splitext(os.path.basename(p))[0]
        # only keep known-ish names to avoid picking random CSVs
        if name.lower() in {
            "sp500", "nasdaq100", "dow30", "russell1000",
            "ftse100", "dax40", "nikkei225", "stoxx600"
        }:
            candidates[name.upper()] = p

    # Stable ordering (UI dropdown)
    ordered = {}
    for key in sorted(candidates.keys()):
        ordered[key] = candidates[key]
    return ordered


def ask_stocks(default_text: str = "") -> Optional[Dict[str, Any]]:
    """UI to choose either manual tickers or an index universe.

    Returns a dict or None if canceled.
      {
        'mode': 'manual' | 'universe',
        'raw': 'AAPL,MSFT' (manual mode only),
        'universe_name': 'SP500' (universe mode only),
        'universe_path': '/path/to/sp500.csv' (universe mode only),
        'analysis_stage': 'portfolio' | 'shortlist' | 'broad',
        'use_fmp': bool
      }
    """
    result: Dict[str, Any] = {
        "mode": "manual",
        "raw": None,
        "universe_name": None,
        "universe_path": None,
        "analysis_stage": None,
        "use_fmp": True,
    }

    # Default: enable FMP only if an API key is present.
    # If FORCE_FMP_FALLBACK=True, the checkbox is locked ON when a key exists.
    has_fmp_key = bool(os.environ.get("FMP_API_KEY", "").strip())
    forced = bool(FORCE_FMP_FALLBACK) and has_fmp_key
    result["use_fmp"] = has_fmp_key

    universes = _discover_universe_csvs()
    universe_names = list(universes.keys()) or ["(no CSVs found)"]

    root = tk.Tk()
    root.title("Stock Report App — Pick Universe / Stocks")
    root.geometry("640x380")
    root.resizable(False, False)

    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    title = tk.Label(root, text="Choose a mode", font=("Segoe UI", 11, "bold"))
    title.pack(pady=(14, 6))

    mode_var = tk.StringVar(value="manual")

    mode_frame = tk.Frame(root)
    mode_frame.pack(pady=(0, 8))

    rb_manual = tk.Radiobutton(mode_frame, text="Manual tickers / company names", variable=mode_var, value="manual")
    rb_univ = tk.Radiobutton(mode_frame, text="Index universe (CSV)", variable=mode_var, value="universe")
    rb_manual.pack(side="left", padx=12)
    rb_univ.pack(side="left", padx=12)

    # Manual entry
    man_frame = tk.Frame(root)
    man_frame.pack(fill="x", padx=16)

    hint = tk.Label(
        man_frame,
        text="Comma separated • Case-insensitive • Dots allowed (e.g. ABC.DE)",
        font=("Segoe UI", 9)
    )
    hint.pack(anchor="w", pady=(0, 6))

    entry = tk.Text(man_frame, height=4, font=("Consolas", 11))
    entry.pack(fill="x")
    if default_text:
        entry.insert("1.0", default_text)

    # Universe picker
    univ_frame = tk.Frame(root)
    univ_frame.pack(fill="x", padx=16, pady=(10, 0))

    univ_label = tk.Label(univ_frame, text="Universe CSV:", font=("Segoe UI", 9))
    univ_label.grid(row=0, column=0, sticky="w")

    univ_var = tk.StringVar(value=universe_names[0])
    univ_menu = tk.OptionMenu(univ_frame, univ_var, *universe_names)
    univ_menu.config(width=24)
    univ_menu.grid(row=0, column=1, sticky="w", padx=(8, 0))

    univ_hint = tk.Label(
        univ_frame,
        text="Place CSVs in ./universes/ or project root (sp500.csv, nasdaq100.csv, ...).",
        font=("Segoe UI", 8)
    )
    univ_hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

    # Stage selector (affects speed + FMP usage)
    stage_frame = tk.Frame(root)
    stage_frame.pack(fill="x", padx=16, pady=(12, 0))

    stage_label = tk.Label(stage_frame, text="Run mode:", font=("Segoe UI", 9))
    stage_label.pack(anchor="w")

    stage_var = tk.StringVar(value=(os.environ.get("ANALYSIS_STAGE", "portfolio") or "portfolio").strip().lower())
    # normalize
    if stage_var.get() not in ("broad", "scan", "wide", "shortlist", "narrow", "screen", "portfolio", "full"):
        stage_var.set("portfolio")

    stage_row = tk.Frame(stage_frame)
    stage_row.pack(anchor="w", pady=(4, 0))

    tk.Radiobutton(stage_row, text="Broad scan (fast, no statements)", variable=stage_var, value="broad").pack(side="left", padx=(0, 10))
    tk.Radiobutton(stage_row, text="Shortlist screen (minimal FMP)", variable=stage_var, value="shortlist").pack(side="left", padx=(0, 10))
    tk.Radiobutton(stage_row, text="Portfolio deep dive (full)", variable=stage_var, value="portfolio").pack(side="left")

    # FMP toggle
    use_fmp_var = tk.BooleanVar(value=result["use_fmp"])
    label = "Use FMP fallback (requires FMP_API_KEY)"
    if forced:
        label = "Use FMP fallback (FORCED ON — FMP_API_KEY detected)"

    chk = tk.Checkbutton(root, text=label, variable=use_fmp_var)
    if forced:
        chk.configure(state="disabled")
        use_fmp_var.set(True)
    chk.pack(pady=(10, 0))

    def _apply_mode_ui():
        is_univ = (mode_var.get() == "universe")
        # Enable/disable manual entry and universe dropdown
        entry.configure(state=("disabled" if is_univ else "normal"))
        univ_menu.configure(state=("normal" if is_univ else "disabled"))

    def on_generate():
        mode = mode_var.get()
        result["mode"] = mode
        result["analysis_stage"] = (stage_var.get() or "").strip().lower()
        result["use_fmp"] = bool(use_fmp_var.get())

        if mode == "manual":
            raw = entry.get("1.0", "end").strip()
            if not raw:
                messagebox.showwarning("Missing input", "Please enter at least one ticker or company name.")
                return
            result["raw"] = raw
        else:
            name = univ_var.get()
            if name.startswith("("):
                messagebox.showwarning("Missing universes", "No universe CSVs were found. Put CSVs into ./universes/ or project root.")
                return
            path = universes.get(name)
            if not path or not os.path.exists(path):
                messagebox.showwarning("Universe missing", f"Could not find CSV file for: {name}")
                return
            result["universe_name"] = name
            result["universe_path"] = path

        root.destroy()

    def on_cancel():
        result.clear()
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=14)

    btn_gen = tk.Button(btn_frame, text="Run", width=18, command=on_generate)
    btn_gen.pack(side="left", padx=8)

    btn_cancel = tk.Button(btn_frame, text="Cancel", width=12, command=on_cancel)
    btn_cancel.pack(side="left", padx=8)

    mode_var.trace_add("write", lambda *_: _apply_mode_ui())
    _apply_mode_ui()

    root.bind("<Return>", lambda _e: on_generate())
    root.bind("<Escape>", lambda _e: on_cancel())

    root.mainloop()
    return result if result else None
