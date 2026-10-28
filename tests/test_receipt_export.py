"""No real bank details, customer identifiers or Hometax credentials in tests."""
import sqlite3
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from receipt_export import (ensure_receipt_schema, prepare_receipt, export_receipts,
                            confirm_issued, SHEET_NS)

M = "{" + SHEET_NS + "}"


def sample_template(path):
    headings = ["용도구분", "발급수단번호", "총거래금액(합계)", "공급가액", "부가가치세", "메모"]
    root = ET.Element(f"{M}worksheet")
    data = ET.SubElement(root, f"{M}sheetData")
    row = ET.SubElement(data, f"{M}row", r="6")
    for col, name in zip("ABCDEF", headings):
        cell = ET.SubElement(row, f"{M}c", r=f"{col}6", t="inlineStr")
        ET.SubElement(ET.SubElement(cell, f"{M}is"), f"{M}t").text = name
    old = ET.SubElement(data, f"{M}row", r="7")
    c = ET.SubElement(old, f"{M}c", r="A7", s="2")
    ET.SubElement(c, f"{M}f").text = "1+2"
    ET.SubElement(c, f"{M}v").text = "3"
    with ZipFile(path, "w") as zf:
        zf.writestr("xl/worksheets/sheet1.xml", ET.tostring(root))
        zf.writestr("xl/worksheets/sheet2.xml", b"<worksheet>keep intact</worksheet>")
        zf.writestr("[Content_Types].xml", b"<Types/>")


def row_values(path, number):
    with ZipFile(path) as zf:
        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        row = next(r for r in root.findall(f".//{M}row") if r.get("r") == str(number))
        cells = {}
        for c in row.findall(f"{M}c"):
            if c.get("t") == "inlineStr":
                t = c.find(f"{M}is/{M}t")
                cells[c.get("r")] = t.text if t is not None else None
            else:
                v = c.find(f"{M}v")
                cells[c.get("r")] = v.text if v is not None else None
        return cells


class ReceiptExportTest(unittest.TestCase):
    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.root = Path(d.name)
        self.db = self.root / "transactions.db"
        self.template = self.root / "official_example.xlsx"
        self.output = self.root / "hometax_output.xlsx"
        with sqlite3.connect(self.db) as conn:
            conn.execute("CREATE TABLE transactions (id INTEGER PRIMARY KEY, transaction_at TEXT, amount_krw INTEGER NOT NULL)")
            conn.executemany("INSERT INTO transactions VALUES (?, ?, ?)", [
                (1, "2026-10-27T17:11:00", 30000),
                (2, "2026-10-27T18:11:00", 50000),
                (3, "2026-10-27T19:11:00", -2000),
                (4, "2026-10-28T09:11:00", 80000),
            ])
        sample_template(self.template)

    def _fill_first(self):
        prepare_receipt(self.db, 1, gmt_sale=1, purpose="1", number="010-1234-5678",
                        supply=27273, vat=2727, memo="주문 1")

    def test_non_destructive_schema_and_defaults(self):
        with sqlite3.connect(self.db) as conn:
            ensure_receipt_schema(conn)
            ensure_receipt_schema(conn)
            row = conn.execute("SELECT receipt_requested,is_gmt_sale,receipt_status FROM transactions WHERE id=1").fetchone()
        self.assertEqual(row, (1, None, "PENDING"))

    def test_gmt_sale_and_number_required(self):
        with self.assertRaisesRegex(ValueError, "발급 준비된 GMT 매출"):
            export_receipts(self.db, self.template, self.output, through="2026-10-27")
        prepare_receipt(self.db, 1, gmt_sale=1)
        with self.assertRaisesRegex(ValueError, "용도구분"):
            export_receipts(self.db, self.template, self.output, through="2026-10-27")
        self.assertFalse(self.output.exists())

    def test_excel_output_preserves_template_other_sheets_and_status(self):
        self._fill_first()
        self.assertEqual(export_receipts(self.db, self.template, self.output, through="2026-10-27"), 1)
        self.assertEqual(row_values(self.output, 7), {"A7": "1", "B7": "010-1234-5678",
                                                       "C7": "30000", "D7": "27273", "E7": "2727", "F7": "주문 1"})
        with ZipFile(self.output) as zf:
            self.assertEqual(zf.read("xl/worksheets/sheet2.xml"), b"<worksheet>keep intact</worksheet>")
            self.assertNotIn(b"<f>", zf.read("xl/worksheets/sheet1.xml"))
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute("SELECT receipt_status,receipt_issued_at,receipt_approval_no FROM transactions WHERE id=1").fetchone(), ("EXPORTED", None, None))
        with self.assertRaisesRegex(ValueError, "발급 준비된"):
            export_receipts(self.db, self.template, self.root / "again.xlsx", through="2026-10-27")
        confirm_issued(self.db, 1, "SYNTHETIC-APPROVAL", "2026-10-28")
        with sqlite3.connect(self.db) as conn:
            self.assertEqual(conn.execute("SELECT receipt_status,receipt_approval_no,receipt_issued_at FROM transactions WHERE id=1").fetchone(), ("ISSUED", "SYNTHETIC-APPROVAL", "2026-10-28"))

    def test_request_false_excluded_even_if_gmt_sale(self):
        prepare_receipt(self.db, 1, gmt_sale=1, requested=0, purpose="1",
                        number="010-1234-5678", supply=30000, vat=0)
        with self.assertRaisesRegex(ValueError, "발급 준비된"):
            export_receipts(self.db, self.template, self.output, through="2026-10-27")

    def test_no_reverse_date_and_no_non_gmt_deposits(self):
        self._fill_first()
        prepare_receipt(self.db, 2, gmt_sale=0)
        prepare_receipt(self.db, 4, gmt_sale=1, purpose="1", number="010-2222-2222", supply=80000, vat=0)
        self.assertEqual(export_receipts(self.db, self.template, self.output, through="2026-10-27"), 1)
        self.assertEqual(row_values(self.output, 7)["C7"], "30000")

    def test_invalid_tax_breakdown_refuses_export(self):
        prepare_receipt(self.db, 1, gmt_sale=1, purpose="1", number="010-1234-5678", supply=30000, vat=3000)
        with self.assertRaisesRegex(ValueError, r"공급가액\+부가가치세"):
            export_receipts(self.db, self.template, self.output, through="2026-10-27")
        self.assertFalse(self.output.exists())

    def test_self_issue_number_refused_when_requested(self):
        prepare_receipt(self.db, 1, gmt_sale=1, purpose="0", number="010-000-1234", supply=30000, vat=0)
        with self.assertRaisesRegex(ValueError, "자진발급 번호"):
            export_receipts(self.db, self.template, self.output, through="2026-10-27")

    def test_reexport_must_be_explicit(self):
        self._fill_first()
        export_receipts(self.db, self.template, self.output, through="2026-10-27")
        with self.assertRaisesRegex(ValueError, "이미 엑셀로 내보낸"):
            prepare_receipt(self.db, 1, requested=0)
        self.assertEqual(export_receipts(self.db, self.template, self.root / "retry.xlsx",
                                         through="2026-10-27", retry_exported=True), 1)

    def test_missing_or_wrong_template_rejected(self):
        self._fill_first()
        wrong = self.root / "wrong.xlsx"
        with ZipFile(wrong, "w") as zf:
            zf.writestr("xl/worksheets/sheet1.xml", b"<worksheet/>")
        with self.assertRaisesRegex(ValueError, "공식 홈택스"):
            export_receipts(self.db, wrong, self.output, through="2026-10-27")


if __name__ == "__main__":
    unittest.main()
