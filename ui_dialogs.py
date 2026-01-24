import os
import tkinter as tk
from tkinter import filedialog

def ask_output_directory(default_subfolder: str = "reports") -> str:
    """Open a native folder picker dialog and return chosen directory.

    If user cancels, returns ./<default_subfolder>.
    Always returns an absolute path.
    """
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    initial = os.path.abspath(os.path.join(os.getcwd(), default_subfolder))
    folder = filedialog.askdirectory(
        title="Select folder to save the Excel report",
        initialdir=initial,
        mustexist=False
    )

    root.destroy()

    if not folder:
        folder = initial

    return os.path.abspath(folder)
