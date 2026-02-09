"""
Unit tests for PLC-linked Excel PDF Converter core logic.
==========================================================
Tests cover non-COM, non-PLC components that can run on any OS.
"""

import os
import sys
import tempfile
import unittest
from datetime import timezone, timedelta
from unittest.mock import MagicMock, patch, PropertyMock

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import (
    sanitize_filename,
    jst_now,
    JST,
    EXCEL_EXTENSIONS,
    FileCountError,
    list_excel_files,
    validate_single_file,
    decode_sheet_bits,
    archive_excel_file,
)

from config import AppConfig


# =====================================================================
# Filename sanitization
# =====================================================================

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


# =====================================================================
# JST timezone
# =====================================================================

class TestJstNow(unittest.TestCase):
    def test_returns_jst_timezone(self):
        now = jst_now()
        self.assertEqual(now.tzinfo, JST)

    def test_jst_offset(self):
        self.assertEqual(JST.utcoffset(None), timedelta(hours=9))

    def test_timestamp_format(self):
        now = jst_now()
        fmt = now.strftime("%Y%m%d_%H%M%S")
        self.assertEqual(len(fmt), 15)
        self.assertEqual(fmt[8], "_")


# =====================================================================
# Excel extensions
# =====================================================================

class TestExcelExtensions(unittest.TestCase):
    def test_supported_extensions(self):
        for ext in [".xlsx", ".xlsm", ".xls", ".xlsb"]:
            self.assertIn(ext, EXCEL_EXTENSIONS)

    def test_unsupported(self):
        self.assertNotIn(".csv", EXCEL_EXTENSIONS)
        self.assertNotIn(".pdf", EXCEL_EXTENSIONS)
        self.assertNotIn(".doc", EXCEL_EXTENSIONS)


# =====================================================================
# Bit-field decoding  (D0 command word → sheet indices)
# =====================================================================

class TestDecodeSheetBits(unittest.TestCase):
    def test_no_bits(self):
        self.assertEqual(decode_sheet_bits(0x0000), [])

    def test_single_bit_0(self):
        """Bit 0 → sheet index 0 (= Sheet 1)"""
        self.assertEqual(decode_sheet_bits(0x0001), [0])

    def test_single_bit_15(self):
        """Bit 15 → sheet index 15 (= Sheet 16)"""
        self.assertEqual(decode_sheet_bits(0x8000), [15])

    def test_multiple_bits(self):
        # bits 0, 2, 4 → indices [0, 2, 4]
        self.assertEqual(decode_sheet_bits(0b0000_0000_0001_0101), [0, 2, 4])

    def test_all_bits(self):
        """All 16 bits ON → indices 0..15"""
        self.assertEqual(decode_sheet_bits(0xFFFF), list(range(16)))

    def test_high_byte_only(self):
        # bits 8-15
        self.assertEqual(decode_sheet_bits(0xFF00), list(range(8, 16)))

    def test_returns_sorted(self):
        result = decode_sheet_bits(0b1010_0000_0000_0101)
        self.assertEqual(result, sorted(result))


# =====================================================================
# Input folder validation  (single-file rule)
# =====================================================================

class TestListExcelFiles(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def test_empty_folder(self):
        self.assertEqual(list_excel_files(self.tmpdir), [])

    def test_one_xlsx(self):
        path = os.path.join(self.tmpdir, "test.xlsx")
        with open(path, "w") as f:
            f.write("fake")
        result = list_excel_files(self.tmpdir)
        self.assertEqual(len(result), 1)
        self.assertTrue(result[0].endswith("test.xlsx"))

    def test_skips_temp_files(self):
        path = os.path.join(self.tmpdir, "~$test.xlsx")
        with open(path, "w") as f:
            f.write("fake")
        self.assertEqual(list_excel_files(self.tmpdir), [])

    def test_skips_non_excel(self):
        path = os.path.join(self.tmpdir, "readme.txt")
        with open(path, "w") as f:
            f.write("hello")
        self.assertEqual(list_excel_files(self.tmpdir), [])

    def test_skips_directories(self):
        os.makedirs(os.path.join(self.tmpdir, "subfolder.xlsx"))
        self.assertEqual(list_excel_files(self.tmpdir), [])

    def test_nonexistent_dir(self):
        self.assertEqual(list_excel_files("/nonexistent/path"), [])

    def test_multiple_excel_files(self):
        for name in ["a.xlsx", "b.xlsm", "c.xls"]:
            with open(os.path.join(self.tmpdir, name), "w") as f:
                f.write("fake")
        self.assertEqual(len(list_excel_files(self.tmpdir)), 3)


class TestValidateSingleFile(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def test_zero_files_raises(self):
        with self.assertRaises(FileCountError):
            validate_single_file(self.tmpdir)

    def test_one_file_ok(self):
        path = os.path.join(self.tmpdir, "report.xlsx")
        with open(path, "w") as f:
            f.write("fake")
        result = validate_single_file(self.tmpdir)
        self.assertTrue(result.endswith("report.xlsx"))

    def test_two_files_raises(self):
        for name in ["a.xlsx", "b.xlsx"]:
            with open(os.path.join(self.tmpdir, name), "w") as f:
                f.write("fake")
        with self.assertRaises(FileCountError):
            validate_single_file(self.tmpdir)


# =====================================================================
# Archive (move) helper
# =====================================================================

class TestArchiveExcelFile(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.archive_dir = os.path.join(self.tmpdir, "archive")
        self.logger = MagicMock()

    def test_basic_move(self):
        src = os.path.join(self.tmpdir, "test.xlsx")
        with open(src, "w") as f:
            f.write("data")
        dest = archive_excel_file(src, self.archive_dir, self.logger)
        self.assertFalse(os.path.exists(src))
        self.assertTrue(os.path.exists(dest))
        self.assertTrue(dest.startswith(self.archive_dir))

    def test_creates_archive_dir(self):
        src = os.path.join(self.tmpdir, "test.xlsx")
        with open(src, "w") as f:
            f.write("data")
        self.assertFalse(os.path.isdir(self.archive_dir))
        archive_excel_file(src, self.archive_dir, self.logger)
        self.assertTrue(os.path.isdir(self.archive_dir))

    def test_timestamp_in_name(self):
        src = os.path.join(self.tmpdir, "report.xlsx")
        with open(src, "w") as f:
            f.write("data")
        dest = archive_excel_file(src, self.archive_dir, self.logger)
        basename = os.path.basename(dest)
        # Should contain original stem + timestamp + extension
        self.assertTrue(basename.startswith("report_"))
        self.assertTrue(basename.endswith(".xlsx"))


# =====================================================================
# AppConfig  (JSON persistence)
# =====================================================================

class TestAppConfig(unittest.TestCase):
    def test_defaults(self):
        cfg = AppConfig()
        self.assertEqual(cfg.plc_ip, "192.168.1.10")
        self.assertEqual(cfg.plc_port, 5000)
        self.assertEqual(cfg.command_device, "D0")
        self.assertEqual(cfg.complete_device, "D1")
        self.assertEqual(cfg.monitor_device, "D2")

    def test_save_and_load(self):
        cfg = AppConfig(plc_ip="10.0.0.1", input_folder="/tmp/input")
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            path = f.name
        try:
            cfg.save(path)
            loaded = AppConfig.load(path)
            self.assertEqual(loaded.plc_ip, "10.0.0.1")
            self.assertEqual(loaded.input_folder, "/tmp/input")
            # Unchanged defaults should persist
            self.assertEqual(loaded.plc_port, 5000)
        finally:
            os.unlink(path)

    def test_load_missing_file(self):
        cfg = AppConfig.load("/nonexistent/path.json")
        self.assertEqual(cfg.plc_ip, "192.168.1.10")

    def test_ignores_unknown_keys(self):
        """Extra keys in JSON should not cause errors."""
        import json
        with tempfile.NamedTemporaryFile(
            suffix=".json", delete=False, mode="w"
        ) as f:
            json.dump({"plc_ip": "1.2.3.4", "unknown_key": 999}, f)
            path = f.name
        try:
            cfg = AppConfig.load(path)
            self.assertEqual(cfg.plc_ip, "1.2.3.4")
        finally:
            os.unlink(path)


# =====================================================================
# PLCConnection  (mock-based tests)
# =====================================================================

class TestPLCConnectionMock(unittest.TestCase):
    """Test PLCConnection methods using mocked pymcprotocol."""

    def _make_connection(self):
        from plc_comm import PLCConnection
        conn = PLCConnection("192.168.1.10", 5000)
        # Replace internal state to simulate connected
        mock_pymc = MagicMock()
        conn._pymc = mock_pymc
        conn._connected = True
        return conn, mock_pymc

    def test_read_word(self):
        conn, mock = self._make_connection()
        mock.batchread_wordunits.return_value = [0x0005]
        result = conn.read_word("D0")
        self.assertEqual(result, 5)
        mock.batchread_wordunits.assert_called_once_with(headdevice="D0", readsize=1)

    def test_write_word(self):
        conn, mock = self._make_connection()
        conn.write_word("D1", 0x1234)
        mock.batchwrite_wordunits.assert_called_once_with(
            headdevice="D1", values=[0x1234])

    def test_set_bit(self):
        conn, mock = self._make_connection()
        mock.batchread_wordunits.return_value = [0x0000]
        conn.set_bit("D2", 3)
        mock.batchwrite_wordunits.assert_called_once_with(
            headdevice="D2", values=[0x0008])

    def test_clear_bit(self):
        conn, mock = self._make_connection()
        mock.batchread_wordunits.return_value = [0x000F]
        conn.clear_bit("D2", 1)
        # 0x000F with bit 1 cleared = 0x000D
        mock.batchwrite_wordunits.assert_called_once_with(
            headdevice="D2", values=[0x000D])

    def test_write_bits_set_and_clear(self):
        conn, mock = self._make_connection()
        mock.batchread_wordunits.return_value = [0x0002]  # bit 1 ON
        conn.write_bits("D2", bits_to_set=[2], bits_to_clear=[1])
        # bit 1 cleared, bit 2 set → 0x0004
        mock.batchwrite_wordunits.assert_called_once_with(
            headdevice="D2", values=[0x0004])

    def test_disconnect(self):
        conn, mock = self._make_connection()
        conn.disconnect()
        mock.close.assert_called_once()
        self.assertFalse(conn.is_connected)


# =====================================================================
# send_oneshot  (mock-based)
# =====================================================================

class TestSendOneshot(unittest.TestCase):
    def test_oneshot_sets_then_clears(self):
        from plc_comm import send_oneshot
        mock_plc = MagicMock()
        send_oneshot(mock_plc, "D1", bit=1, duration_s=0.01)
        mock_plc.set_bit.assert_called_once_with("D1", 1)
        mock_plc.clear_bit.assert_called_once_with("D1", 1)


if __name__ == "__main__":
    unittest.main()
