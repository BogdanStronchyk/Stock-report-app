import os
import tkinter as tk

from config import FORCE_FMP_FALLBACK
from tkinter import messagebox
from typing import Optional, Tuple

def ask_stocks(default_text: str = "") -> Optional[Tuple[str, bool]]:
    """Open a small UI window to collect stock tickers/company names.

    Returns:
      - (raw string, use_fmp_fallback) if user clicks Generate
      - None if user cancels/closes
    """
    result = {"value": None, "use_fmp": True}

    # Default: enable FMP only if an API key is present.
    # If FORCE_FMP_FALLBACK=True, the checkbox is locked ON when a key exists.
    has_fmp_key = bool(os.environ.get("FMP_API_KEY", "").strip())
    forced = bool(FORCE_FMP_FALLBACK) and has_fmp_key
    result["use_fmp"] = has_fmp_key

    root = tk.Tk()
    root.title("Stock Report App — Pick Stocks")
    root.geometry("560x260")
    root.resizable(False, False)

    # Try to keep it on top
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    title = tk.Label(root, text="Enter ticker symbols or company names", font=("Segoe UI", 11, "bold"))
    title.pack(pady=(14, 6))

    hint = tk.Label(
        root,
        text="Comma separated • Case-insensitive • Dots allowed (e.g. ABC.DE)",
        font=("Segoe UI", 9)
    )
    hint.pack(pady=(0, 8))

    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True, padx=16)

    entry = tk.Text(frame, height=4, font=("Consolas", 11))
    entry.pack(fill="both", expand=True)
    if default_text:
        entry.insert("1.0", default_text)

    def on_generate():
        raw = entry.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Missing input", "Please enter at least one ticker or company name.")
            return
        result["value"] = raw
        result["use_fmp"] = bool(use_fmp_var.get())
        root.destroy()

    def on_cancel():
        result["value"] = None
        root.destroy()

    # FMP toggle (optional fallback provider)
    use_fmp_var = tk.BooleanVar(value=result["use_fmp"])
    label = "Use FMP fallback (requires FMP_API_KEY)"
    if forced:
        label = "Use FMP fallback (FORCED ON — FMP_API_KEY detected)"

    chk = tk.Checkbutton(
        root,
        text=label,
        variable=use_fmp_var
    )
    if forced:
        chk.configure(state="disabled")
        use_fmp_var.set(True)
    chk.pack(pady=(6, 0))

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=12)

    btn_gen = tk.Button(btn_frame, text="Generate report", width=18, command=on_generate)
    btn_gen.pack(side="left", padx=8)

    btn_cancel = tk.Button(btn_frame, text="Cancel", width=12, command=on_cancel)
    btn_cancel.pack(side="left", padx=8)

    # Enter triggers generate
    root.bind("<Return>", lambda _e: on_generate())
    root.bind("<Escape>", lambda _e: on_cancel())

    root.mainloop()
    return (result["value"], bool(result.get("use_fmp", False))) if result["value"] else None
