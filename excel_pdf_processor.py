"""
Excel to PDF Individual Sheet Processor (Pro Edition)
=====================================================
GUI application that monitors a folder for Excel files and converts
each sheet to a separate PDF with timestamped filenames.

Requirements:
    - Windows OS with Microsoft Excel installed
    - Python 3.10+ (or standalone EXE)

Dependencies:
    - pywin32 (win32com for Excel automation)
    - watchdog (file system monitoring)
"""

import os
import sys
import logging

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core import FolderMonitor


# ---------------------------------------------------------------------------
# GUI logging handler
# ---------------------------------------------------------------------------

class TextHandlerWidget(logging.Handler):
    """Send log records to a Tkinter ScrolledText widget (thread-safe)."""

    def __init__(self, text_widget: scrolledtext.ScrolledText):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        # Schedule UI update on the main thread
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg: str):
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------

class ExcelPdfProcessorApp:
    """Main application window."""

    APP_TITLE = "Excel to PDF Processor Pro"
    DEFAULT_INTERVAL = 10  # seconds
    MIN_INTERVAL = 1
    MAX_INTERVAL = 36000

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry("780x620")
        self.root.minsize(700, 550)
        self.root.resizable(True, True)

        # State
        self.watch_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.interval_var = tk.IntVar(value=self.DEFAULT_INTERVAL)
        self.monitor: FolderMonitor | None = None

        # Logger
        self.logger = logging.getLogger("ExcelPdfProcessor")
        self.logger.setLevel(logging.DEBUG)

        self._build_ui()
        self._setup_logging()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # -- UI construction -----------------------------------------------------

    def _build_ui(self):
        # Style
        style = ttk.Style()
        style.configure("Status.TLabel", font=("", 11, "bold"))

        pad = {"padx": 8, "pady": 4}

        # --- Folder settings frame ---
        folder_frame = ttk.LabelFrame(self.root, text="Folder Settings / フォルダ設定")
        folder_frame.pack(fill="x", **pad)

        # Watch folder
        ttk.Label(folder_frame, text="Watch Folder / 監視フォルダ:").grid(
            row=0, column=0, sticky="w", **pad
        )
        ttk.Entry(folder_frame, textvariable=self.watch_dir, width=50).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(folder_frame, text="Browse...", command=self._browse_watch).grid(
            row=0, column=2, **pad
        )

        # Output folder
        ttk.Label(folder_frame, text="Output Folder / 出力先フォルダ:").grid(
            row=1, column=0, sticky="w", **pad
        )
        ttk.Entry(folder_frame, textvariable=self.output_dir, width=50).grid(
            row=1, column=1, sticky="ew", **pad
        )
        ttk.Button(folder_frame, text="Browse...", command=self._browse_output).grid(
            row=1, column=2, **pad
        )

        folder_frame.columnconfigure(1, weight=1)

        # --- Interval settings frame ---
        interval_frame = ttk.LabelFrame(
            self.root, text="Monitoring Interval / 監視スパン (seconds)"
        )
        interval_frame.pack(fill="x", **pad)

        self.interval_slider = ttk.Scale(
            interval_frame,
            from_=self.MIN_INTERVAL,
            to=self.MAX_INTERVAL,
            orient="horizontal",
            variable=self.interval_var,
            command=self._on_slider_change,
        )
        self.interval_slider.grid(row=0, column=0, sticky="ew", **pad)

        vcmd = (self.root.register(self._validate_interval), "%P")
        self.interval_entry = ttk.Spinbox(
            interval_frame,
            from_=self.MIN_INTERVAL,
            to=self.MAX_INTERVAL,
            textvariable=self.interval_var,
            width=8,
            validate="key",
            validatecommand=vcmd,
        )
        self.interval_entry.grid(row=0, column=1, **pad)
        ttk.Label(interval_frame, text="sec").grid(row=0, column=2, sticky="w")

        interval_frame.columnconfigure(0, weight=1)

        # --- Control frame ---
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.pack(fill="x", **pad)

        self.start_btn = ttk.Button(
            ctrl_frame, text="Start Monitoring / 監視開始", command=self._start_monitoring
        )
        self.start_btn.pack(side="left", **pad)

        self.stop_btn = ttk.Button(
            ctrl_frame,
            text="Stop Monitoring / 監視停止",
            command=self._stop_monitoring,
            state="disabled",
        )
        self.stop_btn.pack(side="left", **pad)

        # Status indicator
        self.status_label = ttk.Label(
            ctrl_frame, text="● STOPPED / 停止中", foreground="gray", style="Status.TLabel"
        )
        self.status_label.pack(side="right", **pad)

        # --- Log frame ---
        log_frame = ttk.LabelFrame(self.root, text="Log / ログ")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word", font=("Consolas", 9)
        )
        self.log_text.pack(fill="both", expand=True, **pad)

        # Log clear button
        ttk.Button(log_frame, text="Clear Log", command=self._clear_log).pack(
            anchor="e", **pad
        )

    # -- Logging setup -------------------------------------------------------

    def _setup_logging(self):
        handler = TextHandlerWidget(self.log_text)
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        )
        self.logger.addHandler(handler)

        # Also log to stderr for debugging
        stderr_handler = logging.StreamHandler(sys.stderr)
        stderr_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        )
        self.logger.addHandler(stderr_handler)

    # -- Callbacks -----------------------------------------------------------

    def _browse_watch(self):
        d = filedialog.askdirectory(title="Select Watch Folder / 監視フォルダを選択")
        if d:
            self.watch_dir.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select Output Folder / 出力先フォルダを選択")
        if d:
            self.output_dir.set(d)

    def _on_slider_change(self, _value):
        # Slider gives float; round to int
        self.interval_var.set(int(float(_value)))
        if self.monitor and self.monitor.is_running():
            self.monitor.update_interval(self.interval_var.get())

    @staticmethod
    def _validate_interval(value: str) -> bool:
        if value == "":
            return True
        try:
            int(value)
            return True
        except ValueError:
            return False

    def _start_monitoring(self):
        watch = self.watch_dir.get().strip()
        output = self.output_dir.get().strip()

        if not watch:
            messagebox.showwarning(
                self.APP_TITLE,
                "Please select a watch folder.\n監視フォルダを選択してください。",
            )
            return
        if not output:
            messagebox.showwarning(
                self.APP_TITLE,
                "Please select an output folder.\n出力先フォルダを選択してください。",
            )
            return
        if not os.path.isdir(watch):
            messagebox.showerror(
                self.APP_TITLE, f"Watch folder does not exist:\n{watch}"
            )
            return

        os.makedirs(output, exist_ok=True)

        interval = self.interval_var.get()
        if interval < self.MIN_INTERVAL:
            interval = self.MIN_INTERVAL
            self.interval_var.set(interval)

        self.monitor = FolderMonitor(
            watch_dir=watch,
            output_dir=output,
            interval=interval,
            logger=self.logger,
            on_status_change=self._update_status_indicator,
            on_file_processed=self._on_file_processed,
        )
        self.monitor.start()

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.interval_slider.configure(state="normal")

    def _stop_monitoring(self):
        if self.monitor:
            self.monitor.stop()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def _update_status_indicator(self, running: bool):
        if running:
            self.status_label.configure(
                text="● MONITORING / 監視中", foreground="green"
            )
        else:
            self.status_label.configure(
                text="● STOPPED / 停止中", foreground="gray"
            )

    def _on_file_processed(self, filepath: str, success: bool, pdfs: list[str]):
        name = os.path.basename(filepath)
        if success:
            self.logger.info(f"Completed / 完了: {name} ({len(pdfs)} PDF(s))")
        else:
            self.logger.error(f"Failed / 失敗: {name}")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _on_close(self):
        if self.monitor and self.monitor.is_running():
            if not messagebox.askyesno(
                self.APP_TITLE,
                "Monitoring is running. Quit anyway?\n監視中です。終了しますか？",
            ):
                return
            self.monitor.stop()
        self.root.destroy()

    # -- Entry point ---------------------------------------------------------

    def run(self):
        self.logger.info(f"{self.APP_TITLE} started.")
        self.logger.info("Select folders and click 'Start Monitoring'.")
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    app = ExcelPdfProcessorApp()
    app.run()


if __name__ == "__main__":
    main()
