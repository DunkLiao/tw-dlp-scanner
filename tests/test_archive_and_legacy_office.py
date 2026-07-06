import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

from dlp_scanner import (
    ENCRYPTED_FILE_RULE,
    read_ole_text_from_bytes,
    read_xls,
    scan_path,
)


class LegacyOfficeTests(unittest.TestCase):
    def test_read_xls_collects_sheet_cells(self):
        sheet = Mock()
        sheet.name = "Customers"
        sheet.nrows = 2
        sheet.row_values.side_effect = [
            ["Name", "Email"],
            ["Alice", "alice@example.com"],
        ]

        workbook = Mock()
        workbook.sheets.return_value = [sheet]

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.xls"
            path.write_bytes(b"not a real xls because xlrd is mocked")

            with patch("dlp_scanner.xlrd") as xlrd:
                xlrd.open_workbook.return_value = workbook

                text = read_xls(path)

        self.assertIn("[Sheet Customers]", text)
        self.assertIn("alice@example.com", text)

    def test_read_ole_text_from_bytes_extracts_utf16_text(self):
        payload = b"\x00\x01" + "Contact alice@example.com".encode("utf-16le") + b"\x02\x03"

        text = read_ole_text_from_bytes(payload)

        self.assertIn("alice@example.com", text)


class ZipScanTests(unittest.TestCase):
    def test_scan_path_reports_hits_inside_zip(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = Path(temp_dir) / "archive.zip"
            with zipfile.ZipFile(archive_path, "w") as zf:
                zf.writestr("folder/customer.txt", "contact alice@example.com")

            rows = scan_path(archive_path)

        self.assertEqual(1, len(rows))
        self.assertEqual("customer.txt", rows[0]["file_name"])
        self.assertEqual(f"{archive_path}::folder/customer.txt", rows[0]["full_path"])

    def test_scan_path_recurses_into_nested_zip(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            inner_path = Path(temp_dir) / "inner.zip"
            with zipfile.ZipFile(inner_path, "w") as zf:
                zf.writestr("secret.txt", "contact bob@example.com")

            outer_path = Path(temp_dir) / "outer.zip"
            with zipfile.ZipFile(outer_path, "w") as zf:
                zf.write(inner_path, "nested/inner.zip")

            rows = scan_path(outer_path)

        self.assertEqual(1, len(rows))
        self.assertEqual("secret.txt", rows[0]["file_name"])
        self.assertEqual(f"{outer_path}::nested/inner.zip::secret.txt", rows[0]["full_path"])

    def test_scan_path_reports_encrypted_zip_entries_without_reading_them(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = Path(temp_dir) / "archive.zip"
            archive_path.write_bytes(b"placeholder")

            info = zipfile.ZipInfo("secret.txt")
            info.flag_bits |= 0x1
            zf = Mock()
            zf.__enter__ = Mock(return_value=zf)
            zf.__exit__ = Mock(return_value=False)
            zf.infolist.return_value = [info]

            with patch("dlp_scanner.zipfile.ZipFile", return_value=zf):
                rows = scan_path(archive_path)

        self.assertEqual(1, len(rows))
        self.assertEqual(ENCRYPTED_FILE_RULE, rows[0]["hit_type"])
        self.assertEqual(f"{archive_path}::secret.txt", rows[0]["full_path"])
        zf.read.assert_not_called()


if __name__ == "__main__":
    unittest.main()
