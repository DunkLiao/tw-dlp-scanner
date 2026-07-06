import tempfile
import unittest
from pathlib import Path

from dlp_scanner import read_text_file


class TextEncodingTests(unittest.TestCase):
    def test_reads_big5_traditional_chinese_without_dropping_characters(self):
        expected = "台北市中正區個資掃描"
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "big5.txt"
            path.write_bytes(expected.encode("big5"))

            self.assertEqual(read_text_file(path), expected)

    def test_reads_cp950_traditional_chinese_without_dropping_characters(self):
        expected = "臺中市銀行存摺根號測試"
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "cp950.txt"
            path.write_bytes(expected.encode("cp950"))

            self.assertEqual(read_text_file(path), expected)


if __name__ == "__main__":
    unittest.main()
