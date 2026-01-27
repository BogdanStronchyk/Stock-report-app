import tkinter as tk
from tkinter import messagebox
from typing import Optional

def ask_stocks(default_text: str = "") -> Optional[str]:
    """Open a small UI window to collect stock tickers/company names.

    Returns:
      - raw string (comma separated) if user clicks Generate
      - None if user cancels/closes
    """
    result = {"value": None}

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
        root.destroy()

    def on_cancel():
        result["value"] = None
        root.destroy()

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
    return result["value"]
