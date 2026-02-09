"""
PLC-linked Excel PDF Converter – GUI Application (16-sheet edition)
===================================================================
Windows resident application with tkinter GUI.

Features:
    - Folder path configuration (Input / Output / Archive)
    - PLC connection settings (IP, port, device addresses)
    - Start/Stop control for PLC monitoring
    - Live status indicators (Heartbeat, Ready, Busy, Error)
    - Scrollable log viewer
    - Settings persistence via JSON config file
"""

from __future__ import annotations

import logging
import os
import sys
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from config import AppConfig
from core import PLCExcelController


# ---------------------------------------------------------------------------
# GUI logging handler  (thread-safe → schedules writes on main thread)
# ---------------------------------------------------------------------------

class TextHandlerWidget(logging.Handler):
    """Send log records to a Tkinter ScrolledText widget (thread-safe)."""

    def __init__(self, text_widget: scrolledtext.ScrolledText):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
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
    """PLC-linked Excel PDF Converter – main window."""

    APP_TITLE = "PLC Excel PDF Converter (16-Sheet)"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry("850x720")
        self.root.minsize(800, 680)
        self.root.resizable(True, True)

        # Load persisted config (or defaults)
        self.cfg = AppConfig.load()

        # Tkinter variables bound to config
        self.input_folder_var = tk.StringVar(value=self.cfg.input_folder)
        self.output_folder_var = tk.StringVar(value=self.cfg.output_folder)
        self.archive_folder_var = tk.StringVar(value=self.cfg.archive_folder)
        self.plc_ip_var = tk.StringVar(value=self.cfg.plc_ip)
        self.plc_port_var = tk.IntVar(value=self.cfg.plc_port)
        self.cmd_device_var = tk.StringVar(value=self.cfg.command_device)
        self.cmp_device_var = tk.StringVar(value=self.cfg.complete_device)
        self.mon_device_var = tk.StringVar(value=self.cfg.monitor_device)

        # Logger
        self.logger = logging.getLogger("PLCExcelPDF")
        self.logger.setLevel(logging.DEBUG)

        # Controller
        self.controller: PLCExcelController | None = None

        self._build_ui()
        self._setup_logging()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # =================================================================
    # UI construction
    # =================================================================

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Status.TLabel", font=("", 10, "bold"))
        style.configure("Indicator.TLabel", font=("", 10))

        pad = {"padx": 6, "pady": 3}

        # ---- Folder settings ----
        folder_frame = ttk.LabelFrame(self.root, text="フォルダ設定 / Folder Settings")
        folder_frame.pack(fill="x", **pad)

        self._folder_row(folder_frame, 0, "作業中フォルダ (Input):", self.input_folder_var)
        self._folder_row(folder_frame, 1, "PDF出力先 (Output):", self.output_folder_var)
        self._folder_row(folder_frame, 2, "Excel保存用 (Archive):", self.archive_folder_var)
        folder_frame.columnconfigure(1, weight=1)

        # ---- PLC connection settings ----
        plc_frame = ttk.LabelFrame(self.root, text="PLC接続設定 / PLC Connection")
        plc_frame.pack(fill="x", **pad)

        ttk.Label(plc_frame, text="IP:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(plc_frame, textvariable=self.plc_ip_var, width=18).grid(
            row=0, column=1, sticky="w", **pad)

        ttk.Label(plc_frame, text="Port:").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(plc_frame, textvariable=self.plc_port_var, width=8).grid(
            row=0, column=3, sticky="w", **pad)

        ttk.Label(plc_frame, text="指令 (Cmd):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(plc_frame, textvariable=self.cmd_device_var, width=8).grid(
            row=1, column=1, sticky="w", **pad)

        ttk.Label(plc_frame, text="完了 (Cmp):").grid(row=1, column=2, sticky="w", **pad)
        ttk.Entry(plc_frame, textvariable=self.cmp_device_var, width=8).grid(
            row=1, column=3, sticky="w", **pad)

        ttk.Label(plc_frame, text="モニタ (Mon):").grid(row=1, column=4, sticky="w", **pad)
        ttk.Entry(plc_frame, textvariable=self.mon_device_var, width=8).grid(
            row=1, column=5, sticky="w", **pad)

        # ---- Status indicators ----
        status_frame = ttk.LabelFrame(self.root, text="ステータス / Status")
        status_frame.pack(fill="x", **pad)

        self.lbl_heartbeat = ttk.Label(status_frame, text="Heartbeat: --",
                                        style="Indicator.TLabel")
        self.lbl_heartbeat.grid(row=0, column=0, **pad)

        self.lbl_ready = ttk.Label(status_frame, text="Ready: --",
                                    style="Indicator.TLabel")
        self.lbl_ready.grid(row=0, column=1, **pad)

        self.lbl_busy = ttk.Label(status_frame, text="Busy: --",
                                   style="Indicator.TLabel")
        self.lbl_busy.grid(row=0, column=2, **pad)

        self.lbl_error = ttk.Label(status_frame, text="Error: --",
                                    style="Indicator.TLabel")
        self.lbl_error.grid(row=0, column=3, **pad)

        self.lbl_connection = ttk.Label(status_frame, text="PLC: 未接続",
                                         foreground="gray", style="Status.TLabel")
        self.lbl_connection.grid(row=0, column=4, **pad)

        self.lbl_file_count = ttk.Label(status_frame, text="Files: --",
                                         style="Indicator.TLabel")
        self.lbl_file_count.grid(row=0, column=5, **pad)

        # ---- Control buttons ----
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.pack(fill="x", **pad)

        self.start_btn = ttk.Button(
            ctrl_frame, text="接続・監視開始 / Start", command=self._start)
        self.start_btn.pack(side="left", **pad)

        self.stop_btn = ttk.Button(
            ctrl_frame, text="停止 / Stop", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", **pad)

        self.save_btn = ttk.Button(
            ctrl_frame, text="設定保存 / Save Config", command=self._save_config)
        self.save_btn.pack(side="right", **pad)

        # ---- Log ----
        log_frame = ttk.LabelFrame(self.root, text="ログ / Log")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, **pad)

        ttk.Button(log_frame, text="Clear Log", command=self._clear_log).pack(
            anchor="e", **pad)

    def _folder_row(self, parent, row: int, label: str, var: tk.StringVar):
        pad = {"padx": 6, "pady": 3}
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(parent, textvariable=var, width=55).grid(
            row=row, column=1, sticky="ew", **pad)
        ttk.Button(parent, text="参照...",
                   command=lambda: self._browse_folder(var)).grid(
            row=row, column=2, **pad)

    def _browse_folder(self, var: tk.StringVar):
        d = filedialog.askdirectory()
        if d:
            var.set(d)

    # =================================================================
    # Logging setup
    # =================================================================

    def _setup_logging(self):
        handler = TextHandlerWidget(self.log_text)
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                              datefmt="%H:%M:%S"))
        self.logger.addHandler(handler)

        stderr_handler = logging.StreamHandler(sys.stderr)
        stderr_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        self.logger.addHandler(stderr_handler)

    # =================================================================
    # Config sync
    # =================================================================

    def _sync_config_from_gui(self):
        """Copy GUI variable values into the AppConfig object."""
        self.cfg.input_folder = self.input_folder_var.get().strip()
        self.cfg.output_folder = self.output_folder_var.get().strip()
        self.cfg.archive_folder = self.archive_folder_var.get().strip()
        self.cfg.plc_ip = self.plc_ip_var.get().strip()
        try:
            self.cfg.plc_port = int(self.plc_port_var.get())
        except (ValueError, tk.TclError):
            self.cfg.plc_port = 5000
        self.cfg.command_device = self.cmd_device_var.get().strip()
        self.cfg.complete_device = self.cmp_device_var.get().strip()
        self.cfg.monitor_device = self.mon_device_var.get().strip()

    def _save_config(self):
        self._sync_config_from_gui()
        path = self.cfg.save()
        self.logger.info("Config saved: %s", path)

    # =================================================================
    # Start / Stop
    # =================================================================

    def _validate_folders(self) -> bool:
        for name, var in [
            ("作業中フォルダ (Input)", self.input_folder_var),
            ("PDF出力先 (Output)", self.output_folder_var),
            ("Excel保存用 (Archive)", self.archive_folder_var),
        ]:
            val = var.get().strip()
            if not val:
                messagebox.showwarning(self.APP_TITLE,
                                       f"{name} を設定してください。")
                return False
        # Create output & archive if they don't exist
        os.makedirs(self.output_folder_var.get().strip(), exist_ok=True)
        os.makedirs(self.archive_folder_var.get().strip(), exist_ok=True)
        return True

    def _start(self):
        if not self._validate_folders():
            return

        self._sync_config_from_gui()

        try:
            self.controller = PLCExcelController(
                config=self.cfg,
                logger=self.logger,
                on_error=self._on_plc_error,
                on_status=self._on_status_update,
            )
            self.controller.start()
        except Exception as e:
            self.logger.exception("Failed to start controller")
            messagebox.showerror(self.APP_TITLE,
                                 f"PLC接続に失敗しました:\n{e}")
            return

        self.lbl_connection.configure(text="PLC: 接続中", foreground="green")
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._save_config()
        self.logger.info("System started.")

        # Start periodic status refresh
        self._refresh_status()

    def _stop(self):
        if self.controller:
            self.controller.stop()
            self.controller = None

        self.lbl_connection.configure(text="PLC: 未接続", foreground="gray")
        self.lbl_heartbeat.configure(text="Heartbeat: --")
        self.lbl_ready.configure(text="Ready: --")
        self.lbl_busy.configure(text="Busy: --")
        self.lbl_error.configure(text="Error: --")
        self.lbl_file_count.configure(text="Files: --")
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.logger.info("System stopped.")

    # =================================================================
    # Callbacks (called from worker threads – must schedule on main)
    # =================================================================

    def _on_plc_error(self, message: str):
        """Show error popup from worker thread."""
        self.root.after(0, lambda: messagebox.showerror("Error / エラー", message))

    def _on_status_update(self, status: dict):
        """Update status labels from status dict."""
        def _update():
            if "ready" in status:
                self.lbl_ready.configure(
                    text=f"Ready: {'ON' if status['ready'] else 'OFF'}",
                    foreground="green" if status["ready"] else "gray")
            if "file_count" in status:
                self.lbl_file_count.configure(
                    text=f"Files: {status['file_count']}")
        self.root.after(0, _update)

    def _refresh_status(self):
        """Periodically read D2 from PLC and update indicator labels."""
        if self.controller and self.controller.is_running:
            try:
                plc = self.controller._plc
                if plc and plc.is_connected:
                    val = plc.read_word(self.cfg.monitor_device)

                    hb = bool(val & (1 << 0))
                    rdy = bool(val & (1 << 1))
                    bsy = bool(val & (1 << 2))
                    err = bool(val & (1 << 3))

                    self.lbl_heartbeat.configure(
                        text=f"Heartbeat: {'ON' if hb else 'OFF'}",
                        foreground="green" if hb else "gray")
                    self.lbl_ready.configure(
                        text=f"Ready: {'ON' if rdy else 'OFF'}",
                        foreground="green" if rdy else "gray")
                    self.lbl_busy.configure(
                        text=f"Busy: {'ON' if bsy else 'OFF'}",
                        foreground="orange" if bsy else "gray")
                    self.lbl_error.configure(
                        text=f"Error: {'ON' if err else 'OFF'}",
                        foreground="red" if err else "gray")
            except Exception:
                pass  # PLC read may fail transiently

            self.root.after(500, self._refresh_status)

    # =================================================================
    # Misc
    # =================================================================

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _on_close(self):
        if self.controller and self.controller.is_running:
            if not messagebox.askyesno(
                self.APP_TITLE,
                "PLC監視中です。終了しますか？\nMonitoring is active. Quit?"):
                return
            self.controller.stop()
        self.root.destroy()

    # =================================================================
    # Entry point
    # =================================================================

    def run(self):
        self.logger.info("%s started.", self.APP_TITLE)
        self.logger.info("フォルダとPLC設定を行い、「接続・監視開始」を押してください。")
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    app = ExcelPdfProcessorApp()
    app.run()


if __name__ == "__main__":
    main()
