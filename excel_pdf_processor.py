"""
PLC-linked Excel PDF Converter – GUI Application (16-sheet edition)
===================================================================
Three-screen Windows resident application with tkinter GUI.

Screens:
    1. Operation screen (動作画面)  – default at launch, auto-starts monitoring
    2. Settings screen  (設定画面)  – password-protected, folder/PLC config
    3. Test mode screen (テスト画面) – run without PLC using mock connection

Navigation:
    Operation → Settings : requires password (default "0000")
    Operation → Test     : open from operation screen
    Settings / Test → Operation : direct return
"""

from __future__ import annotations

import logging
import os
import sys
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog

from config import AppConfig
from core import PLCExcelController
from plc_comm import MockPLCConnection


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
# Password dialog
# ---------------------------------------------------------------------------

class PasswordDialog(simpledialog.Dialog):
    """Modal dialog asking for a password with masked input."""

    def body(self, master):
        self.title("パスワード / Password")
        ttk.Label(master, text="設定画面のパスワードを入力:").pack(padx=10, pady=(10, 4))
        self.pw_entry = ttk.Entry(master, show="*", width=20)
        self.pw_entry.pack(padx=10, pady=(0, 10))
        self.result_pw: str | None = None
        return self.pw_entry

    def apply(self):
        self.result_pw = self.pw_entry.get()


# =====================================================================
# Main application
# =====================================================================

class ExcelPdfProcessorApp:
    """Three-screen PLC Excel PDF Converter."""

    APP_TITLE = "PLC Excel PDF Converter (16-Sheet)"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry("880x750")
        self.root.minsize(820, 700)
        self.root.resizable(True, True)

        # Config
        self.cfg = AppConfig.load()

        # Logger
        self.logger = logging.getLogger("PLCExcelPDF")
        self.logger.setLevel(logging.DEBUG)

        # Controller reference
        self.controller: PLCExcelController | None = None

        # Container for all screens (stacked frames)
        self._container = ttk.Frame(self.root)
        self._container.pack(fill="both", expand=True)

        # Build the three screens
        self._screens: dict[str, ttk.Frame] = {}
        self._build_operation_screen()
        self._build_settings_screen()
        self._build_test_screen()

        # Logging setup (connects to operation screen log widget)
        self._setup_logging()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Show operation screen first
        self._show_screen("operation")

        # Schedule auto-start after mainloop begins
        if self.cfg.auto_start and self._folders_valid():
            self.root.after(500, self._auto_start)

    # =================================================================
    # Screen switching
    # =================================================================

    def _show_screen(self, name: str):
        """Raise the named screen to the front."""
        for frame in self._screens.values():
            frame.pack_forget()
        self._screens[name].pack(fill="both", expand=True)

    # =================================================================
    # 1. OPERATION SCREEN  (動作画面)
    # =================================================================

    def _build_operation_screen(self):
        frame = ttk.Frame(self._container)
        self._screens["operation"] = frame

        style = ttk.Style()
        style.configure("Status.TLabel", font=("", 10, "bold"))
        style.configure("Indicator.TLabel", font=("", 10))
        style.configure("Title.TLabel", font=("", 14, "bold"))

        pad = {"padx": 6, "pady": 3}

        # ---- Title bar ----
        title_bar = ttk.Frame(frame)
        title_bar.pack(fill="x", **pad)
        ttk.Label(title_bar, text="動作画面 / Operation",
                  style="Title.TLabel").pack(side="left", **pad)

        nav_frame = ttk.Frame(title_bar)
        nav_frame.pack(side="right")
        ttk.Button(nav_frame, text="設定画面",
                   command=self._goto_settings).pack(side="left", padx=3)
        ttk.Button(nav_frame, text="テストモード",
                   command=self._goto_test).pack(side="left", padx=3)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", padx=6)

        # ---- Status indicators ----
        status_frame = ttk.LabelFrame(frame, text="ステータス / Status")
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

        # ---- Folder display (read-only) ----
        info_frame = ttk.LabelFrame(frame, text="現在の設定 / Current Config")
        info_frame.pack(fill="x", **pad)

        self.lbl_input_folder = ttk.Label(info_frame, text="Input: --")
        self.lbl_input_folder.pack(anchor="w", **pad)
        self.lbl_output_folder = ttk.Label(info_frame, text="Output: --")
        self.lbl_output_folder.pack(anchor="w", **pad)
        self.lbl_archive_folder = ttk.Label(info_frame, text="Archive: --")
        self.lbl_archive_folder.pack(anchor="w", **pad)
        self.lbl_plc_info = ttk.Label(info_frame, text="PLC: --")
        self.lbl_plc_info.pack(anchor="w", **pad)

        # ---- Control buttons ----
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill="x", **pad)

        self.op_start_btn = ttk.Button(
            ctrl_frame, text="監視開始 / Start", command=self._start_monitoring)
        self.op_start_btn.pack(side="left", **pad)

        self.op_stop_btn = ttk.Button(
            ctrl_frame, text="監視停止 / Stop", command=self._stop_monitoring,
            state="disabled")
        self.op_stop_btn.pack(side="left", **pad)

        # ---- Log ----
        log_frame = ttk.LabelFrame(frame, text="ログ / Log")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, **pad)

        ttk.Button(log_frame, text="Clear Log", command=self._clear_log).pack(
            anchor="e", **pad)

        self._refresh_config_labels()

    # =================================================================
    # 2. SETTINGS SCREEN  (設定画面) – password-protected
    # =================================================================

    def _build_settings_screen(self):
        frame = ttk.Frame(self._container)
        self._screens["settings"] = frame

        pad = {"padx": 6, "pady": 3}

        # ---- Title bar ----
        title_bar = ttk.Frame(frame)
        title_bar.pack(fill="x", **pad)
        ttk.Label(title_bar, text="設定画面 / Settings",
                  style="Title.TLabel").pack(side="left", **pad)
        ttk.Button(title_bar, text="動作画面に戻る",
                   command=self._goto_operation).pack(side="right", padx=3)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", padx=6)

        # ---- Folder settings ----
        folder_frame = ttk.LabelFrame(frame, text="フォルダ設定 / Folder Settings")
        folder_frame.pack(fill="x", **pad)

        self.input_folder_var = tk.StringVar(value=self.cfg.input_folder)
        self.output_folder_var = tk.StringVar(value=self.cfg.output_folder)
        self.archive_folder_var = tk.StringVar(value=self.cfg.archive_folder)

        self._folder_row(folder_frame, 0, "作業中フォルダ (Input):", self.input_folder_var)
        self._folder_row(folder_frame, 1, "PDF出力先 (Output):", self.output_folder_var)
        self._folder_row(folder_frame, 2, "Excel保存用 (Archive):", self.archive_folder_var)
        folder_frame.columnconfigure(1, weight=1)

        # ---- PLC connection settings ----
        plc_frame = ttk.LabelFrame(frame, text="PLC接続設定 / PLC Connection")
        plc_frame.pack(fill="x", **pad)

        self.plc_ip_var = tk.StringVar(value=self.cfg.plc_ip)
        self.plc_port_var = tk.IntVar(value=self.cfg.plc_port)
        self.cmd_device_var = tk.StringVar(value=self.cfg.command_device)
        self.cmp_device_var = tk.StringVar(value=self.cfg.complete_device)
        self.mon_device_var = tk.StringVar(value=self.cfg.monitor_device)

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

        # ---- Auto-start ----
        auto_frame = ttk.LabelFrame(frame, text="自動起動 / Auto Start")
        auto_frame.pack(fill="x", **pad)
        self.auto_start_var = tk.BooleanVar(value=self.cfg.auto_start)
        ttk.Checkbutton(auto_frame,
                        text="起動時に自動で監視開始する / Start monitoring on launch",
                        variable=self.auto_start_var).pack(anchor="w", **pad)

        # ---- Password change ----
        pw_frame = ttk.LabelFrame(frame, text="パスワード変更 / Change Password")
        pw_frame.pack(fill="x", **pad)

        pw_inner = ttk.Frame(pw_frame)
        pw_inner.pack(fill="x", **pad)

        ttk.Label(pw_inner, text="現在のPW:").grid(row=0, column=0, sticky="w", **pad)
        self.pw_current_entry = ttk.Entry(pw_inner, show="*", width=16)
        self.pw_current_entry.grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(pw_inner, text="新しいPW:").grid(row=0, column=2, sticky="w", **pad)
        self.pw_new_entry = ttk.Entry(pw_inner, show="*", width=16)
        self.pw_new_entry.grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(pw_inner, text="確認:").grid(row=0, column=4, sticky="w", **pad)
        self.pw_confirm_entry = ttk.Entry(pw_inner, show="*", width=16)
        self.pw_confirm_entry.grid(row=0, column=5, sticky="w", **pad)

        ttk.Button(pw_inner, text="変更", command=self._change_password).grid(
            row=0, column=6, **pad)

        # ---- Save button ----
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", **pad)

        ttk.Button(btn_frame, text="設定を保存して戻る / Save & Return",
                   command=self._save_and_return).pack(side="left", **pad)
        ttk.Button(btn_frame, text="保存せず戻る / Cancel",
                   command=self._goto_operation).pack(side="left", **pad)

    # =================================================================
    # 3. TEST MODE SCREEN  (テストモード画面)
    # =================================================================

    def _build_test_screen(self):
        frame = ttk.Frame(self._container)
        self._screens["test"] = frame

        pad = {"padx": 6, "pady": 3}

        # ---- Title bar ----
        title_bar = ttk.Frame(frame)
        title_bar.pack(fill="x", **pad)
        ttk.Label(title_bar, text="テストモード / Test Mode (PLC不要)",
                  style="Title.TLabel").pack(side="left", **pad)
        ttk.Button(title_bar, text="動作画面に戻る",
                   command=self._test_return).pack(side="right", padx=3)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", padx=6)

        # ---- Status indicators (test) ----
        ts_frame = ttk.LabelFrame(frame, text="モックPLCステータス / Mock PLC Status")
        ts_frame.pack(fill="x", **pad)

        self.test_lbl_heartbeat = ttk.Label(ts_frame, text="Heartbeat: --",
                                             style="Indicator.TLabel")
        self.test_lbl_heartbeat.grid(row=0, column=0, **pad)
        self.test_lbl_ready = ttk.Label(ts_frame, text="Ready: --",
                                         style="Indicator.TLabel")
        self.test_lbl_ready.grid(row=0, column=1, **pad)
        self.test_lbl_busy = ttk.Label(ts_frame, text="Busy: --",
                                        style="Indicator.TLabel")
        self.test_lbl_busy.grid(row=0, column=2, **pad)
        self.test_lbl_error = ttk.Label(ts_frame, text="Error: --",
                                         style="Indicator.TLabel")
        self.test_lbl_error.grid(row=0, column=3, **pad)
        self.test_lbl_conn = ttk.Label(ts_frame, text="Mock: 停止",
                                        foreground="gray", style="Status.TLabel")
        self.test_lbl_conn.grid(row=0, column=4, **pad)

        # ---- Sheet bit selection ----
        bit_frame = ttk.LabelFrame(frame, text="D0 シートビット選択 / Sheet Bit Selection")
        bit_frame.pack(fill="x", **pad)

        self._test_bit_vars: list[tk.BooleanVar] = []
        for i in range(16):
            var = tk.BooleanVar(value=False)
            self._test_bit_vars.append(var)
            cb = ttk.Checkbutton(bit_frame, text=f"Bit{i}\n(Sheet{i+1})",
                                 variable=var)
            cb.grid(row=i // 8, column=i % 8, padx=4, pady=4, sticky="w")

        # ---- Quick select ----
        qs_frame = ttk.Frame(frame)
        qs_frame.pack(fill="x", **pad)
        ttk.Button(qs_frame, text="全選択 / Select All",
                   command=lambda: self._test_set_all(True)).pack(side="left", **pad)
        ttk.Button(qs_frame, text="全解除 / Clear All",
                   command=lambda: self._test_set_all(False)).pack(side="left", **pad)

        # ---- D0 hex display ----
        d0_frame = ttk.Frame(frame)
        d0_frame.pack(fill="x", **pad)
        ttk.Label(d0_frame, text="D0値 (hex):").pack(side="left", **pad)
        self.test_d0_display = ttk.Label(d0_frame, text="0x0000",
                                          font=("Consolas", 12, "bold"))
        self.test_d0_display.pack(side="left", **pad)

        # Update display when any checkbox changes
        for var in self._test_bit_vars:
            var.trace_add("write", self._update_d0_display)

        # ---- Control buttons ----
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill="x", **pad)

        self.test_start_btn = ttk.Button(
            ctrl_frame, text="テスト監視開始 / Start Test",
            command=self._test_start)
        self.test_start_btn.pack(side="left", **pad)

        self.test_stop_btn = ttk.Button(
            ctrl_frame, text="テスト停止 / Stop Test",
            command=self._test_stop, state="disabled")
        self.test_stop_btn.pack(side="left", **pad)

        self.test_trigger_btn = ttk.Button(
            ctrl_frame, text="指令送信 / Send D0 Command",
            command=self._test_send_command, state="disabled")
        self.test_trigger_btn.pack(side="left", **pad)

        # ---- Log (test) ----
        log_frame = ttk.LabelFrame(frame, text="テストログ / Test Log")
        log_frame.pack(fill="both", expand=True, **pad)

        self.test_log_text = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word", font=("Consolas", 9))
        self.test_log_text.pack(fill="both", expand=True, **pad)

        ttk.Button(log_frame, text="Clear",
                   command=self._test_clear_log).pack(anchor="e", **pad)

        # Test mode state
        self._test_controller: PLCExcelController | None = None
        self._mock_plc: MockPLCConnection | None = None
        self._test_logger: logging.Logger | None = None

    # =================================================================
    # Helpers
    # =================================================================

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

    def _refresh_config_labels(self):
        """Update read-only config display on operation screen."""
        self.lbl_input_folder.configure(
            text=f"Input: {self.cfg.input_folder or '(未設定)'}")
        self.lbl_output_folder.configure(
            text=f"Output: {self.cfg.output_folder or '(未設定)'}")
        self.lbl_archive_folder.configure(
            text=f"Archive: {self.cfg.archive_folder or '(未設定)'}")
        self.lbl_plc_info.configure(
            text=f"PLC: {self.cfg.plc_ip}:{self.cfg.plc_port}  "
                 f"Cmd={self.cfg.command_device}  Cmp={self.cfg.complete_device}  "
                 f"Mon={self.cfg.monitor_device}")

    def _folders_valid(self) -> bool:
        """Return True if all three folder paths are non-empty."""
        return bool(self.cfg.input_folder and
                    self.cfg.output_folder and
                    self.cfg.archive_folder)

    def _validate_folders_interactive(self) -> bool:
        for name, val in [
            ("作業中フォルダ (Input)", self.cfg.input_folder),
            ("PDF出力先 (Output)", self.cfg.output_folder),
            ("Excel保存用 (Archive)", self.cfg.archive_folder),
        ]:
            if not val.strip():
                messagebox.showwarning(self.APP_TITLE,
                                       f"{name} を設定してください。\n設定画面で設定できます。")
                return False
        os.makedirs(self.cfg.output_folder, exist_ok=True)
        os.makedirs(self.cfg.archive_folder, exist_ok=True)
        return True

    # =================================================================
    # Logging
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
    # Navigation
    # =================================================================

    def _goto_settings(self):
        """Prompt password then switch to settings screen."""
        dlg = PasswordDialog(self.root)
        if dlg.result_pw is None:
            return
        if not self.cfg.verify_password(dlg.result_pw):
            messagebox.showerror(self.APP_TITLE, "パスワードが違います。\nIncorrect password.")
            return
        # Sync current config values into settings entry widgets
        self.input_folder_var.set(self.cfg.input_folder)
        self.output_folder_var.set(self.cfg.output_folder)
        self.archive_folder_var.set(self.cfg.archive_folder)
        self.plc_ip_var.set(self.cfg.plc_ip)
        self.plc_port_var.set(self.cfg.plc_port)
        self.cmd_device_var.set(self.cfg.command_device)
        self.cmp_device_var.set(self.cfg.complete_device)
        self.mon_device_var.set(self.cfg.monitor_device)
        self.auto_start_var.set(self.cfg.auto_start)
        self._show_screen("settings")

    def _goto_operation(self):
        self._refresh_config_labels()
        self._show_screen("operation")

    def _goto_test(self):
        self._show_screen("test")

    # =================================================================
    # Settings screen actions
    # =================================================================

    def _sync_config_from_settings(self):
        """Pull values from settings GUI into self.cfg."""
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
        self.cfg.auto_start = self.auto_start_var.get()

    def _save_and_return(self):
        self._sync_config_from_settings()
        path = self.cfg.save()
        self.logger.info("Config saved: %s", path)
        self._refresh_config_labels()
        self._show_screen("operation")

    def _change_password(self):
        cur = self.pw_current_entry.get()
        new = self.pw_new_entry.get()
        confirm = self.pw_confirm_entry.get()

        if not self.cfg.verify_password(cur):
            messagebox.showerror(self.APP_TITLE, "現在のパスワードが違います。")
            return
        if not new:
            messagebox.showwarning(self.APP_TITLE, "新しいパスワードを入力してください。")
            return
        if new != confirm:
            messagebox.showerror(self.APP_TITLE, "新しいパスワードが一致しません。")
            return

        self.cfg.change_password(new)
        self.cfg.save()
        self.logger.info("Password changed successfully.")
        messagebox.showinfo(self.APP_TITLE, "パスワードを変更しました。")

        # Clear entries
        self.pw_current_entry.delete(0, tk.END)
        self.pw_new_entry.delete(0, tk.END)
        self.pw_confirm_entry.delete(0, tk.END)

    # =================================================================
    # Operation screen: Start / Stop (real PLC)
    # =================================================================

    def _auto_start(self):
        """Called once after mainloop starts if auto_start is enabled."""
        self.logger.info("自動起動: 監視を開始します...")
        self._start_monitoring()

    def _start_monitoring(self):
        if not self._validate_folders_interactive():
            return

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
        self.op_start_btn.configure(state="disabled")
        self.op_stop_btn.configure(state="normal")
        self.logger.info("System started.")
        self._refresh_op_status()

    def _stop_monitoring(self):
        if self.controller:
            self.controller.stop()
            self.controller = None

        self.lbl_connection.configure(text="PLC: 未接続", foreground="gray")
        self._reset_status_labels()
        self.op_start_btn.configure(state="normal")
        self.op_stop_btn.configure(state="disabled")
        self.logger.info("System stopped.")

    def _reset_status_labels(self):
        self.lbl_heartbeat.configure(text="Heartbeat: --", foreground="gray")
        self.lbl_ready.configure(text="Ready: --", foreground="gray")
        self.lbl_busy.configure(text="Busy: --", foreground="gray")
        self.lbl_error.configure(text="Error: --", foreground="gray")
        self.lbl_file_count.configure(text="Files: --")

    # =================================================================
    # Operation screen: callbacks from controller
    # =================================================================

    def _on_plc_error(self, message: str):
        self.root.after(0, lambda: messagebox.showerror("Error / エラー", message))

    def _on_status_update(self, status: dict):
        def _update():
            if "ready" in status:
                self.lbl_ready.configure(
                    text=f"Ready: {'ON' if status['ready'] else 'OFF'}",
                    foreground="green" if status["ready"] else "gray")
            if "file_count" in status:
                self.lbl_file_count.configure(
                    text=f"Files: {status['file_count']}")
        self.root.after(0, _update)

    def _refresh_op_status(self):
        """Periodically read D2 from PLC and update operation screen."""
        if self.controller and self.controller.is_running:
            try:
                plc = self.controller._plc
                if plc and plc.is_connected:
                    val = plc.read_word(self.cfg.monitor_device)
                    self._apply_status_to_labels(
                        val,
                        self.lbl_heartbeat, self.lbl_ready,
                        self.lbl_busy, self.lbl_error)
            except Exception:
                pass
            self.root.after(500, self._refresh_op_status)

    def _apply_status_to_labels(self, val, lbl_hb, lbl_rdy, lbl_bsy, lbl_err):
        hb = bool(val & (1 << 0))
        rdy = bool(val & (1 << 1))
        bsy = bool(val & (1 << 2))
        err = bool(val & (1 << 3))

        lbl_hb.configure(
            text=f"Heartbeat: {'ON' if hb else 'OFF'}",
            foreground="green" if hb else "gray")
        lbl_rdy.configure(
            text=f"Ready: {'ON' if rdy else 'OFF'}",
            foreground="green" if rdy else "gray")
        lbl_bsy.configure(
            text=f"Busy: {'ON' if bsy else 'OFF'}",
            foreground="orange" if bsy else "gray")
        lbl_err.configure(
            text=f"Error: {'ON' if err else 'OFF'}",
            foreground="red" if err else "gray")

    # =================================================================
    # Test mode: Start / Stop / Send
    # =================================================================

    def _test_set_all(self, on: bool):
        for var in self._test_bit_vars:
            var.set(on)

    def _update_d0_display(self, *_args):
        val = 0
        for i, var in enumerate(self._test_bit_vars):
            if var.get():
                val |= (1 << i)
        self.test_d0_display.configure(text=f"0x{val:04X}")

    def _test_start(self):
        """Start monitoring with MockPLCConnection."""
        if not self._validate_folders_interactive():
            return

        # Setup test logger to test log widget
        if self._test_logger is None:
            self._test_logger = logging.getLogger("PLCExcelPDF.Test")
            self._test_logger.setLevel(logging.DEBUG)
            handler = TextHandlerWidget(self.test_log_text)
            handler.setFormatter(
                logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                  datefmt="%H:%M:%S"))
            self._test_logger.addHandler(handler)

        self._mock_plc = MockPLCConnection()

        self._test_controller = PLCExcelController(
            config=self.cfg,
            logger=self._test_logger,
            on_error=self._test_on_error,
            on_status=self._test_on_status,
            plc_connection=self._mock_plc,
        )
        self._test_controller.start()

        self.test_lbl_conn.configure(text="Mock: 稼働中", foreground="green")
        self.test_start_btn.configure(state="disabled")
        self.test_stop_btn.configure(state="normal")
        self.test_trigger_btn.configure(state="normal")
        self._test_logger.info("Test mode started (MockPLC).")
        self._refresh_test_status()

    def _test_stop(self):
        if self._test_controller:
            self._test_controller.stop()
            self._test_controller = None
        self._mock_plc = None

        self.test_lbl_conn.configure(text="Mock: 停止", foreground="gray")
        self.test_lbl_heartbeat.configure(text="Heartbeat: --", foreground="gray")
        self.test_lbl_ready.configure(text="Ready: --", foreground="gray")
        self.test_lbl_busy.configure(text="Busy: --", foreground="gray")
        self.test_lbl_error.configure(text="Error: --", foreground="gray")
        self.test_start_btn.configure(state="normal")
        self.test_stop_btn.configure(state="disabled")
        self.test_trigger_btn.configure(state="disabled")
        if self._test_logger:
            self._test_logger.info("Test mode stopped.")

    def _test_return(self):
        """Stop test mode if running and return to operation screen."""
        if self._test_controller and self._test_controller.is_running:
            self._test_stop()
        self._goto_operation()

    def _test_send_command(self):
        """Write the selected bits to D0 of the mock PLC."""
        if self._mock_plc is None:
            return
        val = 0
        for i, var in enumerate(self._test_bit_vars):
            if var.get():
                val |= (1 << i)
        if val == 0:
            messagebox.showwarning(self.APP_TITLE, "シートを1つ以上選択してください。")
            return
        self._mock_plc.write_word(self.cfg.command_device, val)
        if self._test_logger:
            self._test_logger.info(
                "D0 command sent: 0x%04X (bin: %s)", val, format(val, '016b'))

    def _test_on_error(self, message: str):
        self.root.after(0, lambda: messagebox.showerror("Test Error", message))

    def _test_on_status(self, status: dict):
        def _update():
            if "ready" in status:
                self.test_lbl_ready.configure(
                    text=f"Ready: {'ON' if status['ready'] else 'OFF'}",
                    foreground="green" if status["ready"] else "gray")
        self.root.after(0, _update)

    def _refresh_test_status(self):
        """Periodically read D2 from mock PLC and update test screen."""
        if self._test_controller and self._test_controller.is_running:
            try:
                if self._mock_plc and self._mock_plc.is_connected:
                    val = self._mock_plc.read_word(self.cfg.monitor_device)
                    self._apply_status_to_labels(
                        val,
                        self.test_lbl_heartbeat, self.test_lbl_ready,
                        self.test_lbl_busy, self.test_lbl_error)
            except Exception:
                pass
            self.root.after(500, self._refresh_test_status)

    def _test_clear_log(self):
        self.test_log_text.configure(state="normal")
        self.test_log_text.delete("1.0", tk.END)
        self.test_log_text.configure(state="disabled")

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
        if self._test_controller and self._test_controller.is_running:
            self._test_controller.stop()
        self.root.destroy()

    # =================================================================
    # Entry point
    # =================================================================

    def run(self):
        self.logger.info("%s started.", self.APP_TITLE)
        if self.cfg.auto_start and self._folders_valid():
            self.logger.info("自動起動が有効です。PLC接続を試みます...")
        else:
            self.logger.info("設定画面でフォルダ・PLC設定を行ってください。")
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    app = ExcelPdfProcessorApp()
    app.run()


if __name__ == "__main__":
    main()
