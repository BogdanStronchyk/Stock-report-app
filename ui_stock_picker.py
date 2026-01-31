import os
import tkinter as tk
from tkinter import messagebox, ttk
from config import FORCE_FMP_FALLBACK
from typing import Optional, Tuple, List

def ask_stocks(default_text: str = "") -> Optional[Tuple[str, List[str], str, bool]]:
    """Open a UI window to collect stock tickers, index selections, and rule set."""
    result = {"value": None, "selected_indices": [], "rule_mode": "Strict", "use_fmp": True}
    has_fmp_key = bool(os.environ.get("FMP_API_KEY", "").strip())
    forced = bool(FORCE_FMP_FALLBACK) and has_fmp_key
    result["use_fmp"] = has_fmp_key

    root = tk.Tk()
    root.title("Stock Report App — Pick Stocks & Indices")
    root.geometry("600x550")

    tk.Label(root, text="Enter ticker symbols or company names", font=("Segoe UI", 11, "bold")).pack(pady=(10, 5))
    frame = tk.Frame(root)
    frame.pack(fill="both", expand=False, padx=16)
    entry = tk.Text(frame, height=4, font=("Consolas", 11))
    entry.pack(fill="both", expand=True)
    if default_text: entry.insert("1.0", default_text)

    tk.Label(root, text="Or select from Indices:", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))
    index_frame = tk.Frame(root)
    index_frame.pack(pady=5)
    available_indices = ["SP500", "nasdaq100", "NYSE", "LSE"]
    index_vars = {}
    for idx in available_indices:
        var = tk.BooleanVar(value=False)
        chk = tk.Checkbutton(index_frame, text=idx, variable=var)
        chk.pack(side="left", padx=10)
        index_vars[idx] = var

    # --- Rule Set Selection ---
    tk.Label(root, text="Recommendation Rule Set:", font=("Segoe UI", 10, "bold")).pack(pady=(10, 0))
    mode_var = tk.StringVar(value="Strict (60%)")
    mode_cb = ttk.Combobox(root, textvariable=mode_var, state="readonly", width=25)
    mode_cb['values'] = ("Strict (60%)", "Moderate (50%)", "Loose (40%)")
    mode_cb.pack(pady=5)

    def on_generate():
        raw = entry.get("1.0", "end").strip()
        selected = [name for name, var in index_vars.items() if var.get()]
        if not raw and not selected:
            messagebox.showwarning("Missing input", "Please enter tickers or select an index.")
            return
        result["value"] = raw
        result["selected_indices"] = selected
        result["rule_mode"] = mode_var.get().split(" ")[0]
        result["use_fmp"] = bool(use_fmp_var.get())
        root.destroy()

    def on_cancel():
        result["value"] = None
        root.destroy()

    use_fmp_var = tk.BooleanVar(value=result["use_fmp"])
    chk_fmp = tk.Checkbutton(root, text="Use FMP fallback", variable=use_fmp_var)
    if forced: chk_fmp.configure(state="disabled")
    chk_fmp.pack(pady=(10, 0))

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="Generate report", width=18, command=on_generate).pack(side="left", padx=8)
    tk.Button(btn_frame, text="Cancel", width=12, command=on_cancel).pack(side="left", padx=8)

    root.mainloop()
    if result["value"] is not None or result["selected_indices"]:
        return (result["value"] or "", result["selected_indices"], result["rule_mode"], bool(result["use_fmp"]))
    return None