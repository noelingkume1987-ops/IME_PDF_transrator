"""
Core logic for Excel to PDF Individual Sheet Processor (Pro Edition).
=====================================================================
Contains conversion, monitoring, and utility functions.
Separated from the GUI so it can be tested and reused independently.
"""

import os
import time
import shutil
import logging
import threading
from datetime import datetime, timezone, timedelta
from pathlib import Path

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


# ---------------------------------------------------------------------------
# Excel -> PDF conversion (pywin32)
# ---------------------------------------------------------------------------

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}


def convert_excel_to_pdfs(
    excel_path: str, output_dir: str, logger: logging.Logger
) -> list[str]:
    """Open *excel_path* in a hidden Excel instance, export every visible
    worksheet as its own PDF into *output_dir*, and return the list of
    generated PDF paths.

    Naming rule:
        {original_filename}_{sheet_name}_{YYYYMMDD_HHMMSS}.pdf
    """
    # Late import so the module can be loaded on non-Windows for tests
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore

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
        logger.info(f"Opening: {abs_path}")
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True)

        base_name = Path(excel_path).stem  # e.g. "見積書"
        timestamp = jst_now().strftime("%Y%m%d_%H%M%S")

        for sheet in wb.Worksheets:
            if sheet.Visible != -1:  # -1 = xlSheetVisible
                continue

            safe_sheet = sanitize_filename(sheet.Name)
            pdf_name = f"{base_name}_{safe_sheet}_{timestamp}.pdf"
            pdf_path = os.path.join(output_dir, pdf_name)

            logger.info(f"  Exporting sheet '{sheet.Name}' -> {pdf_name}")
            # xlTypePDF = 0
            sheet.ExportAsFixedFormat(0, pdf_path)
            generated.append(pdf_path)

    except Exception:
        logger.exception("Error during Excel conversion")
        raise
    finally:
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
        pythoncom.CoUninitialize()

    return generated


# ---------------------------------------------------------------------------
# Folder monitor (polling-based)
# ---------------------------------------------------------------------------

class FolderMonitor:
    """Polls *watch_dir* every *interval* seconds for new / modified Excel
    files, converts them, and moves originals into an ``Archive`` subfolder.
    """

    def __init__(
        self,
        watch_dir: str,
        output_dir: str,
        interval: float,
        logger: logging.Logger,
        on_status_change=None,
        on_file_processed=None,
    ):
        self.watch_dir = watch_dir
        self.output_dir = output_dir
        self.interval = interval
        self.logger = logger
        self.on_status_change = on_status_change
        self.on_file_processed = on_file_processed

        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._processed: set[tuple[str, int]] = set()

    # -- public API ----------------------------------------------------------

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self.logger.info("Monitoring started.")
        if self.on_status_change:
            self.on_status_change(True)

    def stop(self):
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=5)
        self._thread = None
        self.logger.info("Monitoring stopped.")
        if self.on_status_change:
            self.on_status_change(False)

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def update_interval(self, interval: float):
        self.interval = interval

    # -- internals -----------------------------------------------------------

    def _run(self):
        while not self._stop_event.is_set():
            try:
                self._scan_once()
            except Exception:
                self.logger.exception("Error during folder scan")
            self._stop_event.wait(self.interval)

    def _scan_once(self):
        if not os.path.isdir(self.watch_dir):
            return

        for entry in os.scandir(self.watch_dir):
            if not entry.is_file():
                continue
            ext = Path(entry.name).suffix.lower()
            if ext not in EXCEL_EXTENSIONS:
                continue
            # Skip temporary Excel files (start with ~$)
            if entry.name.startswith("~$"):
                continue

            file_key = (entry.path, entry.stat().st_mtime_ns)
            if file_key in self._processed:
                continue

            self._process_file(entry.path)
            self._processed.add(file_key)

    def _process_file(self, filepath: str):
        self.logger.info(f"Detected: {filepath}")

        # Wait briefly to ensure the file is fully written
        time.sleep(1)

        try:
            pdfs = convert_excel_to_pdfs(filepath, self.output_dir, self.logger)
            self.logger.info(f"  Generated {len(pdfs)} PDF(s)")
        except Exception:
            self.logger.exception(f"  Failed to convert: {filepath}")
            if self.on_file_processed:
                self.on_file_processed(filepath, False, [])
            return

        # Move original to Archive
        archive_dir = os.path.join(self.watch_dir, "Archive")
        os.makedirs(archive_dir, exist_ok=True)
        archive_dest = os.path.join(archive_dir, os.path.basename(filepath))
        # Avoid overwrite in archive
        if os.path.exists(archive_dest):
            stem = Path(filepath).stem
            ext = Path(filepath).suffix
            ts = jst_now().strftime("%Y%m%d_%H%M%S")
            archive_dest = os.path.join(archive_dir, f"{stem}_{ts}{ext}")
        try:
            shutil.move(filepath, archive_dest)
            self.logger.info(f"  Archived -> {archive_dest}")
        except Exception:
            self.logger.exception(f"  Failed to archive: {filepath}")

        if self.on_file_processed:
            self.on_file_processed(filepath, True, pdfs)
