"""
Core logic for PLC-linked Excel PDF Converter (16-sheet edition).
================================================================
Monitors Mitsubishi PLC D0 (16-bit command word), converts the
requested Excel sheets to PDF, archives the original, and signals
completion/status back to the PLC.

Key design decisions
--------------------
* **Bit-loop for sheet selection** – D0 bits 0-15 are scanned with a
  bitmask ``(1 << i)`` inside a ``for i in range(16)`` loop.  Each ON
  bit maps to sheet index ``i`` (0-based) which corresponds to the
  ``(i+1)``-th worksheet in the Excel file.
* **Quit-before-Move guarantee** – The Excel COM object is always
  ``Quit()``-ed (and ``CoUninitialize``-d) in a ``finally`` block
  *before* ``shutil.move()`` is called, preventing file-lock errors.
* **Single-file enforcement** – The input folder is validated to contain
  exactly one Excel file; 0 or ≥2 raises ``FileCountError``.
"""

from __future__ import annotations

import logging
import os
import shutil
import threading
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

from plc_comm import PLCConnection, HeartbeatThread, send_oneshot
from config import AppConfig

# ---------------------------------------------------------------------------
# JST timezone helper
# ---------------------------------------------------------------------------
JST = timezone(timedelta(hours=9))


def jst_now() -> datetime:
    """Return the current datetime in JST."""
    return datetime.now(JST)


# ---------------------------------------------------------------------------
# Filename utilities
# ---------------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """Remove or replace characters that are illegal in Windows file names."""
    illegal = r'<>:"/\|?*'
    for ch in illegal:
        name = name.replace(ch, "_")
    return name.strip()


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}


# ---------------------------------------------------------------------------
# Custom exceptions
# ---------------------------------------------------------------------------

class FileCountError(Exception):
    """Raised when the input folder does not contain exactly 1 Excel file."""
    pass


# ---------------------------------------------------------------------------
# Input folder validation
# ---------------------------------------------------------------------------

def list_excel_files(folder: str) -> list[str]:
    """Return paths of Excel files in *folder* (non-recursive, skips temp)."""
    results: list[str] = []
    if not os.path.isdir(folder):
        return results
    for entry in os.scandir(folder):
        if not entry.is_file():
            continue
        if entry.name.startswith("~$"):
            continue
        if Path(entry.name).suffix.lower() in EXCEL_EXTENSIONS:
            results.append(entry.path)
    return results


def validate_single_file(folder: str) -> str:
    """Return the single Excel file path or raise ``FileCountError``."""
    files = list_excel_files(folder)
    if len(files) == 0:
        raise FileCountError(
            "作業中フォルダにExcelファイルがありません (0個)。\n"
            "No Excel file found in the input folder."
        )
    if len(files) >= 2:
        names = ", ".join(os.path.basename(f) for f in files)
        raise FileCountError(
            f"作業中フォルダにファイルが{len(files)}個あります（1個のみ許可）。\n"
            f"Multiple files detected: {names}"
        )
    return files[0]


# ---------------------------------------------------------------------------
# Bit-field helpers
# ---------------------------------------------------------------------------

def decode_sheet_bits(command_word: int) -> list[int]:
    """Return a sorted list of 0-based sheet indices whose bits are ON.

    >>> decode_sheet_bits(0b0000_0000_0000_0101)
    [0, 2]
    """
    indices: list[int] = []
    for i in range(16):
        if command_word & (1 << i):
            indices.append(i)
    return indices


# ---------------------------------------------------------------------------
# Excel → PDF conversion  (per-sheet, pywin32)
# ---------------------------------------------------------------------------

def convert_sheets_to_pdf(
    excel_path: str,
    sheet_indices: list[int],
    output_dir: str,
    logger: logging.Logger,
) -> list[str]:
    """Open *excel_path*, export each worksheet at *sheet_indices*
    (0-based) to a separate PDF, then **close and Quit Excel completely**
    before returning.

    Returns the list of generated PDF paths.

    IMPORTANT – The caller MUST NOT attempt to move/delete the Excel file
    until this function has returned, because the finally block guarantees
    Excel is fully released.
    """
    import win32com.client  # type: ignore
    import pythoncom        # type: ignore

    pythoncom.CoInitialize()

    generated: list[str] = []
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        abs_path = os.path.abspath(excel_path)
        logger.info("Opening: %s", abs_path)
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True)

        sheet_count = wb.Worksheets.Count
        base_name = Path(excel_path).stem
        timestamp = jst_now().strftime("%Y%m%d_%H%M%S")

        for idx in sheet_indices:
            # Worksheets collection is 1-based in COM
            sheet_num = idx + 1
            if sheet_num > sheet_count:
                logger.warning(
                    "  Sheet index %d (sheet %d) out of range "
                    "(workbook has %d sheets) – skipped.",
                    idx, sheet_num, sheet_count,
                )
                continue

            sheet = wb.Worksheets(sheet_num)
            safe_sheet = sanitize_filename(sheet.Name)
            pdf_name = f"{base_name}_{safe_sheet}_{timestamp}.pdf"
            pdf_path = os.path.join(output_dir, pdf_name)

            logger.info("  Exporting sheet %d '%s' -> %s", sheet_num, sheet.Name, pdf_name)
            # xlTypePDF = 0
            sheet.ExportAsFixedFormat(0, pdf_path)
            generated.append(pdf_path)

    except Exception:
        logger.exception("Error during Excel conversion")
        raise
    finally:
        # ============================================================
        # CRITICAL: Close workbook & Quit Excel BEFORE any file move.
        # This releases the file lock held by the Excel COM process.
        # ============================================================
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        # Small delay to let the OS fully release handles
        time.sleep(0.3)
        pythoncom.CoUninitialize()

    return generated


# ---------------------------------------------------------------------------
# Archive (move) helper
# ---------------------------------------------------------------------------

def archive_excel_file(excel_path: str, archive_folder: str, logger: logging.Logger) -> str:
    """Move *excel_path* into *archive_folder* with a timestamp suffix
    to prevent overwriting.  Returns the destination path.

    MUST be called AFTER Excel has been Quit()-ed.
    """
    os.makedirs(archive_folder, exist_ok=True)

    stem = Path(excel_path).stem
    ext = Path(excel_path).suffix
    ts = jst_now().strftime("%Y%m%d_%H%M%S")
    dest_name = f"{stem}_{ts}{ext}"
    dest_path = os.path.join(archive_folder, dest_name)

    # Extra safety: if timestamp collision, add counter
    counter = 1
    while os.path.exists(dest_path):
        dest_name = f"{stem}_{ts}_{counter}{ext}"
        dest_path = os.path.join(archive_folder, dest_name)
        counter += 1

    shutil.move(excel_path, dest_path)
    logger.info("Archived -> %s", dest_path)
    return dest_path


# =====================================================================
# Main controller – ties PLC, conversion, and file management together
# =====================================================================

class PLCExcelController:
    """Runs the PLC polling loop, triggers conversion, and manages
    the monitor/completion device signals.

    Lifecycle::

        ctrl = PLCExcelController(config, logger, error_callback)
        ctrl.start()   # connects PLC, starts heartbeat & polling threads
        ...
        ctrl.stop()    # stops everything cleanly
    """

    def __init__(
        self,
        config: AppConfig,
        logger: logging.Logger,
        on_error: callable | None = None,
        on_status: callable | None = None,
    ):
        self.cfg = config
        self.logger = logger
        self.on_error = on_error      # callback(message: str)
        self.on_status = on_status    # callback(status_dict)

        self._plc: PLCConnection | None = None
        self._heartbeat: HeartbeatThread | None = None
        self._poll_thread: threading.Thread | None = None
        self._stop_event = threading.Event()

    # -----------------------------------------------------------------
    # Public API
    # -----------------------------------------------------------------

    def start(self) -> None:
        """Connect to PLC and launch background threads."""
        self._plc = PLCConnection(self.cfg.plc_ip, self.cfg.plc_port)
        self._plc.connect()

        # Clear all output devices on start
        self._plc.write_word(self.cfg.complete_device, 0)
        self._plc.write_word(self.cfg.monitor_device, 0)

        # Heartbeat thread (D2.0 toggle)
        self._heartbeat = HeartbeatThread(
            self._plc, self.cfg.monitor_device, self.cfg.heartbeat_interval
        )
        self._heartbeat.start()

        # Polling thread
        self._stop_event.clear()
        self._poll_thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._poll_thread.start()

        self.logger.info("Controller started – polling D0 every %.1fs", self.cfg.plc_poll_interval)

    def stop(self) -> None:
        """Disconnect and stop all threads."""
        self._stop_event.set()

        if self._heartbeat:
            self._heartbeat.stop()
            self._heartbeat = None

        if self._poll_thread:
            self._poll_thread.join(timeout=5)
            self._poll_thread = None

        if self._plc and self._plc.is_connected:
            # Clear monitor word before disconnect
            try:
                self._plc.write_word(self.cfg.monitor_device, 0)
            except Exception:
                pass
            self._plc.disconnect()
        self._plc = None

        self.logger.info("Controller stopped.")

    @property
    def is_running(self) -> bool:
        return self._poll_thread is not None and self._poll_thread.is_alive()

    # -----------------------------------------------------------------
    # Main polling loop
    # -----------------------------------------------------------------

    def _poll_loop(self) -> None:
        """Runs on a daemon thread.  Reads D0 every *plc_poll_interval*
        seconds.  When D0 ≠ 0, triggers the conversion pipeline.
        """
        while not self._stop_event.is_set():
            try:
                self._update_ready_status()

                command = self._plc.read_word(self.cfg.command_device)
                if command != 0:
                    self._handle_command(command)

            except Exception:
                self.logger.exception("Poll loop error")
                self._set_error(True)

            self._stop_event.wait(self.cfg.plc_poll_interval)

    # -----------------------------------------------------------------
    # Status helpers (D2)
    # -----------------------------------------------------------------

    def _update_ready_status(self) -> None:
        """Set or clear Ready (D2.1) depending on folder state."""
        try:
            files = list_excel_files(self.cfg.input_folder)
            is_ready = len(files) == 1

            if is_ready:
                self._plc.write_bits(
                    self.cfg.monitor_device,
                    bits_to_set=[1],       # Ready ON
                    bits_to_clear=[3],     # Error OFF
                )
            else:
                # Not ready – set error if files ≠ 1 and files > 0
                if len(files) >= 2:
                    self._plc.write_bits(
                        self.cfg.monitor_device,
                        bits_to_set=[3],       # Error ON
                        bits_to_clear=[1],     # Ready OFF
                    )
                else:
                    # 0 files – just not ready, not necessarily error
                    self._plc.clear_bit(self.cfg.monitor_device, 1)  # Ready OFF

            if self.on_status:
                self.on_status({
                    "ready": is_ready,
                    "file_count": len(files),
                })

        except Exception:
            self.logger.exception("Error updating ready status")

    def _set_busy(self, on: bool) -> None:
        if on:
            self._plc.set_bit(self.cfg.monitor_device, 2)
        else:
            self._plc.clear_bit(self.cfg.monitor_device, 2)

    def _set_error(self, on: bool) -> None:
        if on:
            self._plc.set_bit(self.cfg.monitor_device, 3)
        else:
            self._plc.clear_bit(self.cfg.monitor_device, 3)

    # -----------------------------------------------------------------
    # Command handler
    # -----------------------------------------------------------------

    def _handle_command(self, command_word: int) -> None:
        """Process a non-zero D0 value.

        Steps:
        1. Busy ON
        2. Validate single file
        3. Decode bits → sheet indices
        4. Convert sheets to PDF  (Excel Quit inside finally)
        5. Move Excel to archive  (AFTER Quit)
        6. Completion one-shot (D1.1  100 ms)
        7. Busy OFF
        """
        self.logger.info("=== Command received: D0 = 0x%04X (bin: %s) ===",
                         command_word, format(command_word, '016b'))

        # 1. Busy ON
        self._set_busy(True)

        try:
            # 2. Validate exactly 1 file
            try:
                excel_path = validate_single_file(self.cfg.input_folder)
            except FileCountError as e:
                msg = str(e)
                self.logger.error(msg)
                self._set_error(True)
                if self.on_error:
                    self.on_error(msg)
                return

            # 3. Decode sheet bits (bit 0 → sheet index 0, ...)
            sheet_indices = decode_sheet_bits(command_word)
            self.logger.info(
                "Sheets to convert (0-based): %s",
                [i for i in sheet_indices],
            )

            if not sheet_indices:
                self.logger.warning("No bits set in command word – nothing to do.")
                return

            # 4. Convert selected sheets to PDF
            #    Excel is fully Quit()-ed inside convert_sheets_to_pdf's finally.
            pdfs = convert_sheets_to_pdf(
                excel_path,
                sheet_indices,
                self.cfg.output_folder,
                self.logger,
            )
            self.logger.info("Generated %d PDF(s).", len(pdfs))

            # 5. Move Excel to archive (AFTER Excel process is gone)
            archive_excel_file(excel_path, self.cfg.archive_folder, self.logger)

            # 6. Completion one-shot: D1 bit 1 ON for 100 ms
            send_oneshot(self._plc, self.cfg.complete_device, bit=1, duration_s=0.1)
            self.logger.info("Completion pulse sent (D1.1, 100 ms).")

            # Clear error flag on success
            self._set_error(False)

        except Exception:
            self.logger.exception("Error during command handling")
            self._set_error(True)
            if self.on_error:
                self.on_error("変換処理中にエラーが発生しました。ログを確認してください。")
        finally:
            # 7. Busy OFF
            self._set_busy(False)
