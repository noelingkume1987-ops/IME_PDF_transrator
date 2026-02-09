"""
Unit tests for Excel PDF Processor core logic (non-GUI, non-COM components).
"""

import os
import sys
import tempfile
import unittest
from datetime import timezone, timedelta
from unittest.mock import MagicMock

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import (
    sanitize_filename,
    jst_now,
    JST,
    EXCEL_EXTENSIONS,
    FolderMonitor,
)


class TestSanitizeFilename(unittest.TestCase):
    def test_clean_name(self):
        self.assertEqual(sanitize_filename("Sheet1"), "Sheet1")

    def test_japanese_name(self):
        self.assertEqual(sanitize_filename("見積書_A社"), "見積書_A社")

    def test_illegal_characters(self):
        self.assertEqual(sanitize_filename('a<b>c:d"e'), "a_b_c_d_e")

    def test_backslash_and_pipe(self):
        self.assertEqual(sanitize_filename("a\\b|c"), "a_b_c")

    def test_question_and_star(self):
        self.assertEqual(sanitize_filename("sheet?*"), "sheet__")

    def test_slash(self):
        self.assertEqual(sanitize_filename("a/b"), "a_b")

    def test_empty_string(self):
        self.assertEqual(sanitize_filename(""), "")

    def test_trailing_spaces(self):
        self.assertEqual(sanitize_filename("  hello  "), "hello")


class TestJstNow(unittest.TestCase):
    def test_returns_jst_timezone(self):
        now = jst_now()
        self.assertEqual(now.tzinfo, JST)

    def test_jst_offset(self):
        self.assertEqual(JST.utcoffset(None), timedelta(hours=9))

    def test_timestamp_format(self):
        now = jst_now()
        fmt = now.strftime("%Y%m%d_%H%M%S")
        # Should be 15 chars: YYYYMMDD_HHMMSS
        self.assertEqual(len(fmt), 15)
        self.assertEqual(fmt[8], "_")


class TestExcelExtensions(unittest.TestCase):
    def test_supported_extensions(self):
        for ext in [".xlsx", ".xlsm", ".xls", ".xlsb"]:
            self.assertIn(ext, EXCEL_EXTENSIONS)

    def test_unsupported(self):
        self.assertNotIn(".csv", EXCEL_EXTENSIONS)
        self.assertNotIn(".pdf", EXCEL_EXTENSIONS)
        self.assertNotIn(".doc", EXCEL_EXTENSIONS)


class TestFolderMonitor(unittest.TestCase):
    def setUp(self):
        self.logger = MagicMock()
        self.tmpdir = tempfile.mkdtemp()
        self.outdir = tempfile.mkdtemp()

    def test_init(self):
        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        self.assertEqual(monitor.watch_dir, self.tmpdir)
        self.assertEqual(monitor.output_dir, self.outdir)
        self.assertEqual(monitor.interval, 5)
        self.assertFalse(monitor.is_running())

    def test_start_stop(self):
        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=1,
            logger=self.logger,
        )
        monitor.start()
        self.assertTrue(monitor.is_running())
        monitor.stop()
        self.assertFalse(monitor.is_running())

    def test_double_start(self):
        """Starting twice should not create a second thread."""
        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=1,
            logger=self.logger,
        )
        monitor.start()
        thread1 = monitor._thread
        monitor.start()
        thread2 = monitor._thread
        self.assertIs(thread1, thread2)
        monitor.stop()

    def test_update_interval(self):
        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        monitor.update_interval(20)
        self.assertEqual(monitor.interval, 20)

    def test_scan_skips_temp_files(self):
        """Temporary Excel files (~$...) should be ignored."""
        temp_file = os.path.join(self.tmpdir, "~$test.xlsx")
        with open(temp_file, "w") as f:
            f.write("fake")

        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        monitor._scan_once()
        self.assertEqual(len(monitor._processed), 0)

    def test_scan_skips_non_excel(self):
        """Non-Excel files should be ignored."""
        txt_file = os.path.join(self.tmpdir, "readme.txt")
        with open(txt_file, "w") as f:
            f.write("hello")

        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        monitor._scan_once()
        self.assertEqual(len(monitor._processed), 0)

    def test_scan_skips_directories(self):
        """Subdirectories should be ignored."""
        subdir = os.path.join(self.tmpdir, "subfolder.xlsx")
        os.makedirs(subdir)

        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        monitor._scan_once()
        self.assertEqual(len(monitor._processed), 0)

    def test_scan_nonexistent_dir(self):
        """Scanning a non-existent directory should not raise."""
        monitor = FolderMonitor(
            watch_dir="/nonexistent/path",
            output_dir=self.outdir,
            interval=5,
            logger=self.logger,
        )
        monitor._scan_once()  # Should not raise
        self.assertEqual(len(monitor._processed), 0)

    def test_status_change_callback(self):
        callback = MagicMock()
        monitor = FolderMonitor(
            watch_dir=self.tmpdir,
            output_dir=self.outdir,
            interval=1,
            logger=self.logger,
            on_status_change=callback,
        )
        monitor.start()
        callback.assert_called_with(True)
        monitor.stop()
        callback.assert_called_with(False)


if __name__ == "__main__":
    unittest.main()
