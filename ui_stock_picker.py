import os
import glob
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional, Dict, Any

from config import FORCE_FMP_FALLBACK


def _discover_universe_csvs() -> Dict[str, str]:
    """Find universe CSVs in common locations.

    Returns mapping: display_name -> filepath
    """
    candidates: Dict[str, str] = {}

    # 1) ./universes/*.csv
    if os.path.exists(os.path.join(os.getcwd(), "universes")):
        for p in glob.glob(os.path.join(os.getcwd(), "universes", "*.csv")):
            name = os.path.splitext(os.path.basename(p))[0]
            candidates[name.upper()] = p

    # 2) project root *.csv (for quick drop-in)
    for p in glob.glob(os.path.join(os.getcwd(), "*.csv")):
        name = os.path.splitext(os.path.basename(p))[0]
        # only keep known-ish names to avoid picking random CSVs
        if name.lower() in {
            "sp500", "nasdaq100", "dow30", "russell1000",
            "ftse100", "dax40", "nikkei225", "stoxx600", "universe"
        }:
            candidates[name.upper()] = p

    # Stable ordering
    return dict(sorted(candidates.items()))


def ask_stocks(default_text: str = "") -> Optional[Dict[str, Any]]:
    """UI to choose manual tickers or multiple index universes.

    Returns a dict or None if canceled.
      {
        'mode': 'manual' | 'universe',
        'raw': 'AAPL,MSFT' (manual mode only),
        'universe_paths': ['/path/to/sp500.csv', ...],
        'use_fmp': bool,
        'eligibility_mode': 'strict' | 'permissible' | 'loose'
      }
    """
    result: Dict[str, Any] = {
        "mode": "manual",
        "raw": None,
        "universe_paths": [],
        "use_fmp": False,
        "eligibility_mode": "strict"
    }

    # FMP Toggle Logic
    has_fmp_key = bool(os.environ.get("FMP_API_KEY", "").strip())
    forced = bool(FORCE_FMP_FALLBACK) and has_fmp_key
    result["use_fmp"] = has_fmp_key

    universes = _discover_universe_csvs()

    root = tk.Tk()
    root.title("Stock Report App — Portfolio Generator")

    # --- CHANGE: Increased Height for better button visibility ---
    root.geometry("680x750")
    root.resizable(True, True)

    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    # -- Header --
    title = tk.Label(root, text="Portfolio Candidates Generator", font=("Segoe UI", 14, "bold"))
    title.pack(pady=(15, 5))

    mode_var = tk.StringVar(value="manual")

    # -- Mode Selection --
    mode_frame = tk.Frame(root)
    mode_frame.pack(pady=(0, 10))

    rb_manual = tk.Radiobutton(mode_frame, text="Manual Tickers", variable=mode_var, value="manual",
                               font=("Segoe UI", 10))
    rb_univ = tk.Radiobutton(mode_frame, text="Select Universes (CSV)", variable=mode_var, value="universe",
                             font=("Segoe UI", 10))
    rb_manual.pack(side="left", padx=15)
    rb_univ.pack(side="left", padx=15)

    # -- Manual Entry Section --
    man_frame = tk.Frame(root)
    man_frame.pack(fill="x", padx=20, pady=5)

    hint = tk.Label(
        man_frame,
        text="Enter tickers separated by commas (e.g. AAPL, MSFT, GOOG):",
        font=("Segoe UI", 9)
    )
    hint.pack(anchor="w", pady=(0, 5))

    entry = tk.Text(man_frame, height=3, font=("Consolas", 11))
    entry.pack(fill="x")
    if default_text:
        entry.insert("1.0", default_text)

    # -- Universe Selection Section (Scrollable Checkboxes) --
    univ_frame = tk.LabelFrame(root, text="Available Universes", padx=10, pady=10)
    univ_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # Canvas + Scrollbar for list of checkboxes
    canvas = tk.Canvas(univ_frame)
    scrollbar = ttk.Scrollbar(univ_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Populate checkboxes
    univ_vars = {}
    if not universes:
        tk.Label(scrollable_frame, text="No CSV files found in ./universes/", fg="red").pack(anchor="w")
    else:
        for name, path in universes.items():
            var = tk.BooleanVar()
            chk = tk.Checkbutton(scrollable_frame, text=name, variable=var, font=("Segoe UI", 10))
            chk.pack(anchor="w", pady=2)
            univ_vars[path] = var

    # -- Filtering Strategy (NEW) --
    strat_frame = tk.LabelFrame(root, text="Filtering Strategy", padx=15, pady=10)
    strat_frame.pack(fill="x", padx=20, pady=(5, 10))

    rule_var = tk.StringVar(value="strict")

    r_strict = tk.Radiobutton(strat_frame, text="Strict (Portfolio Candidates)", variable=rule_var, value="strict",
                              font=("Segoe UI", 9, "bold"))
    r_perm = tk.Radiobutton(strat_frame, text="Standard (Watchlist)", variable=rule_var, value="permissible",
                            font=("Segoe UI", 9))
    r_loose = tk.Radiobutton(strat_frame, text="Loose (Broad Scan)", variable=rule_var, value="loose",
                             font=("Segoe UI", 9))

    r_strict.pack(side="left", padx=(0, 20))
    r_perm.pack(side="left", padx=(0, 20))
    r_loose.pack(side="left")

    # -- FMP Toggle --
    fmp_frame = tk.Frame(root)
    fmp_frame.pack(pady=(0, 10))

    use_fmp_var = tk.BooleanVar(value=result["use_fmp"])
    fmp_label = "Enable FMP Data Fallback"
    if forced:
        fmp_label += " (Locked: FORCE_FMP_FALLBACK=1)"
    elif not has_fmp_key:
        fmp_label += " (Requires FMP_API_KEY)"

    chk_fmp = tk.Checkbutton(fmp_frame, text=fmp_label, variable=use_fmp_var, font=("Segoe UI", 9))
    if forced:
        chk_fmp.configure(state="disabled")
        use_fmp_var.set(True)
    chk_fmp.pack()

    # -- Logic --
    def _apply_mode_ui():
        is_univ = (mode_var.get() == "universe")
        entry.configure(state=("disabled" if is_univ else "normal"))
        entry.configure(bg=("#f0f0f0" if is_univ else "#ffffff"))

        if is_univ:
            univ_frame.configure(fg="black", text="Available Universes (Select at least one)")
        else:
            univ_frame.configure(fg="gray", text="Available Universes (Disabled)")

    def on_generate():
        mode = mode_var.get()
        result["mode"] = mode
        result["use_fmp"] = bool(use_fmp_var.get())
        result["eligibility_mode"] = rule_var.get()

        if mode == "manual":
            raw = entry.get("1.0", "end").strip()
            if not raw:
                messagebox.showwarning("Input Required", "Please enter at least one ticker.")
                return
            result["raw"] = raw
        else:
            selected_paths = [path for path, var in univ_vars.items() if var.get()]
            if not selected_paths:
                messagebox.showwarning("Selection Required", "Please select at least one universe CSV.")
                return
            result["universe_paths"] = selected_paths

        root.destroy()

    def on_cancel():
        result.clear()
        root.destroy()

    # -- Buttons --
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=15)

    btn_gen = tk.Button(btn_frame, text="GENERATE REPORT", width=25, height=2, bg="#e1f5fe",
                        font=("Segoe UI", 10, "bold"), command=on_generate)
    btn_gen.pack(side="left", padx=10)

    btn_cancel = tk.Button(btn_frame, text="Cancel", width=15, height=2, command=on_cancel)
    btn_cancel.pack(side="left", padx=10)

    mode_var.trace_add("write", lambda *_: _apply_mode_ui())
    _apply_mode_ui()

    root.bind("<Return>", lambda _e: on_generate())
    root.bind("<Escape>", lambda _e: on_cancel())

    root.mainloop()
    return result if result else None