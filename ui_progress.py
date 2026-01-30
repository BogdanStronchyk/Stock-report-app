import os
import tkinter as tk
from tkinter import ttk, messagebox


class ProgressWindow:
    """
    Modern progress window to visualize the Portfolio Candidate generation stages.
    """

    def __init__(self, total_steps: int, title: str = "Processing..."):
        self.total_steps = max(1, int(total_steps))
        self.current = 0

        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("500x220")
        self.root.resizable(False, False)

        # Keep on top
        try:
            self.root.attributes("-topmost", True)
        except Exception:
            pass

        # -- Main Container --
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # -- Header --
        self.lbl_main = tk.Label(
            main_frame,
            text="Initializing...",
            font=("Segoe UI", 11, "bold"),
            wraplength=460,
            justify="center"
        )
        self.lbl_main.pack(pady=(0, 10))

        # -- Progress Bar --
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure("green.Horizontal.TProgressbar", background='#4caf50')

        self.pb = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            length=460,
            mode="determinate",
            style="green.Horizontal.TProgressbar"
        )
        self.pb.pack(pady=5)
        self.pb["maximum"] = self.total_steps
        self.pb["value"] = 0

        # -- Percentage Label --
        self.lbl_pct = tk.Label(main_frame, text="0%", font=("Segoe UI", 9), fg="#666")
        self.lbl_pct.pack(pady=(2, 10))

        # -- Detail Lines --
        self.lbl_sub = tk.Label(
            main_frame,
            text="Waiting for workers...",
            font=("Segoe UI", 9),
            fg="#333"
        )
        self.lbl_sub.pack(pady=(0, 2))

        self.lbl_done = tk.Label(
            main_frame,
            text="",
            font=("Segoe UI", 8),
            fg="#777"
        )
        self.lbl_done.pack(pady=(5, 0))

        self.root.update()

    def set_status(self, main_text: str, sub_text: str = ""):
        """Update text headers without moving the bar."""
        self.lbl_main.config(text=main_text)
        self.lbl_sub.config(text=sub_text)
        self.root.update()

    def step(self, main_text: str = None, sub_text: str = "", done_text: str = None):
        """Advance progress and update status labels."""
        self.current = min(self.total_steps, self.current + 1)
        self.pb["value"] = self.current

        # Calculate percentage
        pct = int((self.current / self.total_steps) * 100)
        self.lbl_pct.config(text=f"{pct}%")

        if main_text:
            self.lbl_main.config(text=main_text)

        if sub_text:
            self.lbl_sub.config(text=sub_text)

        if done_text:
            self.lbl_done.config(text=done_text)

        self.root.update()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass


def success_popup(message: str):
    """
    Shows a success popup. If message is a path, offers to open folder.
    """
    is_path = os.path.exists(message) or os.path.isabs(message)

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    if is_path:
        # It's a file path
        folder = os.path.dirname(os.path.abspath(message))
        msg_body = f"Report generated successfully!\n\nSaved to:\n{message}\n\nOpen containing folder?"
        open_it = messagebox.askyesno("Process Complete", msg_body)
        if open_it:
            try:
                os.startfile(folder)
            except Exception:
                messagebox.showinfo("Folder", folder)
    else:
        # It's just a message
        messagebox.showinfo("Process Complete", message)

    root.destroy()