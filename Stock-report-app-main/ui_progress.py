import os
import tkinter as tk
from tkinter import ttk, messagebox

class ProgressWindow:
    """Small always-on-top progress window for report generation."""

    def __init__(self, total_steps: int, title: str = "Generating report..."):
        self.total_steps = max(1, int(total_steps))
        self.current = 0

        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("520x160")
        self.root.resizable(False, False)

        try:
            self.root.attributes("-topmost", True)
        except Exception:
            pass

        self.label = tk.Label(self.root, text="Starting...", font=("Segoe UI", 10, "bold"))
        self.label.pack(pady=(18, 10))

        self.pb = ttk.Progressbar(self.root, orient="horizontal", length=460, mode="determinate")
        self.pb.pack(pady=6)
        self.pb["maximum"] = self.total_steps
        self.pb["value"] = 0

        self.sub = tk.Label(self.root, text="", font=("Segoe UI", 9))
        self.sub.pack(pady=(6, 0))

        self.root.update_idletasks()
        self.root.update()

    def set_status(self, main_text: str, sub_text: str = ""):
        self.label.config(text=main_text)
        self.sub.config(text=sub_text)
        self.root.update_idletasks()
        self.root.update()

    def step(self, main_text: str = None, sub_text: str = ""):
        self.current = min(self.total_steps, self.current + 1)
        self.pb["value"] = self.current
        if main_text is not None:
            self.label.config(text=main_text)
        self.sub.config(text=sub_text)
        self.root.update_idletasks()
        self.root.update()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass


def success_popup(filepath: str):
    """Shows a success popup and offers to open the folder (Windows)."""
    folder = os.path.dirname(os.path.abspath(filepath))
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    msg = f"Report saved successfully!\n\n{filepath}\n\nOpen containing folder?"
    open_it = messagebox.askyesno("✅ Report generated", msg)

    if open_it:
        try:
            os.startfile(folder)  # Windows only
        except Exception:
            messagebox.showinfo("Folder path", folder)

    root.destroy()
