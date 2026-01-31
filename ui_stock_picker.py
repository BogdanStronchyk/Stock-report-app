import os
import tkinter as tk
from tkinter import messagebox, ttk
from config import FORCE_FMP_FALLBACK
from typing import Optional, Tuple, List

def ask_stocks(default_text: str = "") -> Optional[Tuple[str, List[str], str, bool, bool]]:
    """Open a UI window to collect stock tickers, index selections, rule set, and Watch preference."""
    result = {"value": None, "selected_indices": [], "rule_mode": "Strict", "use_fmp": True, "include_watch": False}
    has_fmp_key = bool(os.environ.get("FMP_API_KEY", "").strip())
    forced = bool(FORCE_FMP_FALLBACK) and has_fmp_key
    result["use_fmp"] = has_fmp_key

    root = tk.Tk()
    root.title("Stock Report App")
    root.geometry("600x600")

    # Manual Input
    tk.Label(root, text="Enter ticker symbols or company names", font=("Segoe UI", 11, "bold")).pack(pady=(10, 5))
    entry = tk.Text(root, height=4, font=("Consolas", 11))
    entry.pack(fill="x", padx=20)
    if default_text: entry.insert("1.0", default_text)

    # Index Selection
    tk.Label(root, text="Or select from Indices:", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))
    index_frame = tk.Frame(root)
    index_frame.pack(pady=5)
    available_indices = ["SP500", "nasdaq100", "NYSE", "LSE"]
    index_vars = {}
    for idx in available_indices:
        var = tk.BooleanVar(value=False)
        tk.Checkbutton(index_frame, text=idx, variable=var).pack(side="left", padx=10)
        index_vars[idx] = var

    # Rule Set Selection
    tk.Label(root, text="Recommendation Rule Set:", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))
    mode_var = tk.StringVar(value="Strict (60%)")
    mode_cb = ttk.Combobox(root, textvariable=mode_var, state="readonly", width=25)
    mode_cb['values'] = ("Strict (60%)", "Moderate (50%)", "Loose (40%)")
    mode_cb.pack(pady=5)

    # Options Frame (FMP + Watch)
    opt_frame = tk.LabelFrame(root, text="Analysis Options", padx=10, pady=10)
    opt_frame.pack(fill="x", padx=20, pady=10)

    use_fmp_var = tk.BooleanVar(value=result["use_fmp"])
    chk_fmp = tk.Checkbutton(opt_frame, text="Use FMP fallback", variable=use_fmp_var)
    if forced: chk_fmp.configure(state="disabled")
    chk_fmp.pack(anchor="w")

    include_watch_var = tk.BooleanVar(value=False)
    tk.Checkbutton(opt_frame, text="Include 'WATCH' candidates (High Fundamentals / Low Technicals)", variable=include_watch_var).pack(anchor="w")

    def on_generate():
        raw = entry.get("1.0", "end").strip()
        selected = [name for name, var in index_vars.items() if var.get()]
        if not raw and not selected:
            messagebox.showwarning("Input Error", "Please provide tickers.")
            return
        result["value"] = raw
        result["selected_indices"] = selected
        result["rule_mode"] = mode_var.get().split(" ")[0]
        result["use_fmp"] = bool(use_fmp_var.get())
        result["include_watch"] = bool(include_watch_var.get())
        root.destroy()

    tk.Button(root, text="Generate Report", width=20, bg="#2ecc71", fg="white", font=("Segoe UI", 10, "bold"), command=on_generate).pack(pady=20)
    root.mainloop()

    if result["value"] is not None or result["selected_indices"]:
        return (result["value"] or "", result["selected_indices"], result["rule_mode"], result["use_fmp"], result["include_watch"])
    return None