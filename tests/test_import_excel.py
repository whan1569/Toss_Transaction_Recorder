import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.etree.ElementTree import Element, SubElement, tostring
from zipfile import ZipFile

from import_excel import _get_excel_password, import_excel

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def synthetic_excel(path: Path, transactions: list[tuple], account: str = "****-****-0123") -> None:
    root = Element(f"{{{NS}}}worksheet")
    sheet = SubElement(root, f"{{{NS}}}sheetData")
    data = [
        [None, "계좌번호", account],
        [],
        [None, "거래 일시", "적요", "거래 유형", "거래 기관", "계좌번호", "거래 금액", "거래 후 잔액", "메모"],
    ]
    data.extend([[None, *values] for values in transactions])
    for row_no, cells in enumerate(data, 1):
        row = SubElement(sheet, f"{{{NS}}}row", r=str(row_no))
        for index, value in enumerate(cells):
            if value is None:
                continue
            ref = chr(ord('A')+index)+str(row_no)
            if isinstance(value, int):
                cell = SubElement(row, f"{{{NS}}}c", r=ref)
                SubElement(cell, f"{{{NS}}}v").text = str(value)
            else:
                cell = SubElement(row, f"{{{NS}}}c", r=ref, t="inlineStr")
                SubElement(SubElement(cell, f"{{{NS}}}is"), f"{{{NS}}}t").text = value
    with ZipFile(path, "w") as zf:
        zf.writestr("xl/worksheets/sheet1.xml", tostring(root, encoding="utf-8"))


class ImportTransactionsTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.db = self.root / "transactions.db"
        self.first = self.root / "first.xlsx"
        self.second = self.root / "second.xlsx"
        self.row1 = ("2026.10.28 09:01:00", "가상입금자 A", "입금", "토스뱅크", None, 30000, 80000, None)
        self.row2 = ("2026.10.28 09:02:00", "가상입금자 B", "입금", "토스뱅크", None, 50000, 130000, None)
        self.row3 = ("2026.10.28 09:03:00", "카드가맹점", "체크카드결제", None, None, -2000, 128000, None)

    def count(self):
        with sqlite3.connect(self.db) as db:
            return db.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]

    def test_dotenv_password_is_loaded_without_prompt(self):
        (self.root / ".env").write_text("TOSS_EXCEL_PASSWORD=local-test-password\n", encoding="utf-8")
        with patch.dict(os.environ, {}, clear=True), \
             patch("import_excel.__file__", str(self.root / "import_excel.py")), \
             patch("import_excel.getpass", side_effect=AssertionError("unexpected password prompt")):
            self.assertEqual(_get_excel_password(), "local-test-password")

    def test_environment_variable_takes_precedence_over_dotenv(self):
        (self.root / ".env").write_text("TOSS_EXCEL_PASSWORD=local-test-password\n", encoding="utf-8")
        with patch.dict(os.environ, {"TOSS_EXCEL_PASSWORD": "configured-password"}, clear=True), \
             patch("import_excel.__file__", str(self.root / "import_excel.py")), \
             patch("import_excel.getpass", side_effect=AssertionError("unexpected password prompt")):
            self.assertEqual(_get_excel_password(), "configured-password")

    def test_same_file_twice_is_noop(self):
        synthetic_excel(self.first, [self.row1, self.row2])
        self.assertEqual(import_excel(self.first, self.db), (2, 2, 0))
        self.assertEqual(import_excel(self.first, self.db), (2, 0, 2))
        self.assertEqual(self.count(), 2)

    def test_overlapping_date_windows_keep_only_new_row(self):
        synthetic_excel(self.first, [self.row1, self.row2])
        synthetic_excel(self.second, [self.row2, self.row3])
        self.assertEqual(import_excel(self.first, self.db), (2, 2, 0))
        self.assertEqual(import_excel(self.second, self.db), (2, 1, 1))
        self.assertEqual(self.count(), 3)

    def test_same_second_different_balance_not_merged(self):
        other = ("2026.10.28 09:01:00", "가상입금자 A", "입금", "토스뱅크", None, 30000, 110000, None)
        synthetic_excel(self.first, [self.row1, other])
        self.assertEqual(import_excel(self.first, self.db), (2, 2, 0))

    def test_identical_rows_inside_one_file_keep_occurrences(self):
        synthetic_excel(self.first, [self.row1, self.row1])
        synthetic_excel(self.second, [self.row1])
        self.assertEqual(import_excel(self.first, self.db), (2, 2, 0))
        self.assertEqual(import_excel(self.second, self.db), (1, 0, 1))
        self.assertEqual(self.count(), 2)

    def test_separate_source_account_not_merged(self):
        synthetic_excel(self.first, [self.row1], account="****-****-0123")
        synthetic_excel(self.second, [self.row1], account="****-****-0987")
        self.assertEqual(import_excel(self.first, self.db), (1, 1, 0))
        self.assertEqual(import_excel(self.second, self.db), (1, 1, 0))
        self.assertEqual(self.count(), 2)

    def test_invalid_excel_does_not_change_existing_db(self):
        synthetic_excel(self.first, [self.row1])
        bad = ("invalid date", *self.row3[1:])
        synthetic_excel(self.second, [self.row2, bad])
        import_excel(self.first, self.db)
        with self.assertRaises(ValueError):
            import_excel(self.second, self.db)
        self.assertEqual(self.count(), 1)


if __name__ == "__main__":
    unittest.main()
