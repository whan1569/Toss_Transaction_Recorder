"""Generate Hometax cash-receipt XLSX from confirmed GMT payments.

This does NOT issue a cash receipt. It edits a locally supplied *official* Hometax
XLSX template in place, preserving every other worksheet and workbook component.
No payment is assumed to be GMT revenue until explicitly confirmed.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone
import os
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile, is_zipfile

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", SHEET_NS)
M = "{" + SHEET_NS + "}"
HEADERS = ("용도구분", "발급수단번호", "총거래금액(합계)", "공급가액", "부가가치세", "메모")
MAX_ROWS = 1000
# Kept as independent transaction columns so GMT Hub can later populate them.
RECEIPT_COLUMNS = {
    "receipt_requested": "INTEGER NOT NULL DEFAULT 1 CHECK(receipt_requested IN (0, 1))",
    "is_gmt_sale": "INTEGER CHECK(is_gmt_sale IN (0, 1))",
    "receipt_purpose_code": "TEXT",
    "receipt_number": "TEXT",
    "receipt_supply_krw": "INTEGER",
    "receipt_vat_krw": "INTEGER",
    "receipt_memo": "TEXT",
    "receipt_order_id": "TEXT",
    "receipt_status": "TEXT NOT NULL DEFAULT 'PENDING'",
    "receipt_exported_at": "TEXT",
    "receipt_issued_at": "TEXT",
    "receipt_confirmed_at": "TEXT",
    "receipt_approval_no": "TEXT",
}


def ensure_receipt_schema(db: sqlite3.Connection) -> None:
    """Idempotent non-destructive ALTER migration, compatible with existing SQLite DBs."""
    existing = {row[1] for row in db.execute("PRAGMA table_info(transactions)")}
    if "id" not in existing or "amount_krw" not in existing:
        raise ValueError("기존 거래 DB의 transactions 테이블을 찾지 못했습니다. 먼저 import_excel.py를 실행하세요.")
    for col, definition in RECEIPT_COLUMNS.items():
        if col not in existing:
            db.execute(f"ALTER TABLE transactions ADD COLUMN {col} {definition}")
    db.execute("CREATE INDEX IF NOT EXISTS idx_receipt_export ON transactions(is_gmt_sale, receipt_requested, receipt_status, transaction_at)")


def _db_connect(db_path: Path) -> sqlite3.Connection:
    if not db_path.is_file():
        raise FileNotFoundError(f"거래 DB가 없습니다: {db_path}. import_excel.py를 먼저 실행하세요.")
    db = sqlite3.connect(db_path, timeout=30)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA busy_timeout = 30000")
    try:
        ensure_receipt_schema(db)
        db.commit()
    except Exception:
        db.close()
        raise
    return db


def _yn(value: str | None) -> int | None:
    if value is None:
        return None
    return 1 if value == "yes" else 0


def prepare_receipt(
    db_path: Path, transaction_id: int, *, gmt_sale: int | None = None,
    requested: int | None = None, purpose: str | None = None,
    number: str | None = None, supply: int | None = None,
    vat: int | None = None, memo: str | None = None,
    order_id: str | None = None,
) -> None:
    """Manual placeholder for future GMT Hub data sync, not a receipt issue action."""
    changes = {
        "is_gmt_sale": gmt_sale, "receipt_requested": requested,
        "receipt_purpose_code": purpose, "receipt_number": number,
        "receipt_supply_krw": supply, "receipt_vat_krw": vat,
        "receipt_memo": memo, "receipt_order_id": order_id,
    }
    changes = {k: v for k, v in changes.items() if v is not None}
    if not changes:
        raise ValueError("변경할 발급정보가 없습니다.")
    if supply is not None and supply < 0 or vat is not None and vat < 0:
        raise ValueError("공급가액과 부가가치세는 0 이상이어야 합니다.")
    with _db_connect(db_path) as db:
        row = db.execute("SELECT amount_krw, receipt_status FROM transactions WHERE id=?", (transaction_id,)).fetchone()
        if row is None:
            raise ValueError(f"거래 ID {transaction_id}를 찾지 못했습니다.")
        if row["receipt_status"] == "ISSUED":
            raise ValueError("이미 발급 완료한 거래는 변경할 수 없습니다.")
        if gmt_sale == 1 and row["amount_krw"] <= 0:
            raise ValueError("출금 거래는 GMT 현금영수증 발급 대상으로 지정할 수 없습니다.")
        # Exported transactions cannot be silently modified and re-issued.
        if row["receipt_status"] == "EXPORTED":
            raise ValueError("이미 엑셀로 내보낸 거래입니다. 홈택스 발급 결과를 확인한 후 처리하세요.")
        assignments = ", ".join(f"{key}=?" for key in changes)
        db.execute(f"UPDATE transactions SET {assignments} WHERE id=?", (*changes.values(), transaction_id))


def _shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(node.itertext()) for node in root.findall(f"{M}si")]


def _cell_text(node: ET.Element, strings: list[str]) -> str:
    t = node.get("t")
    if t == "inlineStr":
        inline = node.find(f"{M}is")
        return "".join(inline.itertext()) if inline is not None else ""
    v = node.find(f"{M}v")
    if v is None or v.text is None:
        return ""
    return strings[int(v.text)] if t == "s" else v.text


def _col_number(ref: str) -> int:
    token = re.match(r"[A-Z]+", ref.upper())
    if token is None:
        raise ValueError(f"잘못된 엑셀 셀 주소: {ref!r}")
    col = 0
    for letter in token.group(0):
        col = col * 26 + ord(letter) - 64
    return col


def _col_name(n: int) -> str:
    result = ""
    while n:
        n, digit = divmod(n - 1, 26)
        result = chr(65 + digit) + result
    return result


def _find_target(zf: ZipFile) -> tuple[str, ET.Element, tuple[int, ...]]:
    strings = _shared_strings(zf)
    worksheet_paths = [name for name in zf.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)]
    for path in worksheet_paths:
        root = ET.fromstring(zf.read(path))
        data = root.find(f"{M}sheetData")
        if data is None:
            continue
        for row in data.findall(f"{M}row"):
            if row.get("r") != "6":
                continue
            names = {_cell_text(c, strings).strip(): _col_number(c.get("r", "")) for c in row.findall(f"{M}c")}
            if all(name in names for name in HEADERS):
                cols = tuple(names[name] for name in HEADERS)
                if sorted(cols) != list(range(cols[0], cols[0] + 6)):
                    raise ValueError("홈택스 필수 6개 열이 연속적이지 않습니다. 공식 양식을 확인하세요.")
                return path, root, cols
    raise ValueError("공식 홈택스 엑셀 6행에서 필수 6개 항목을 찾지 못했습니다. 원본 .xlsx 양식을 사용하세요.")


def _write_cell(row: ET.Element, column: int, row_no: int, value: object) -> None:
    addr = f"{_col_name(column)}{row_no}"
    cells = row.findall(f"{M}c")
    cell = next((c for c in cells if c.get("r") == addr), None)
    if cell is None:
        cell = ET.Element(f"{M}c", {"r": addr})
        before = next((i for i, c in enumerate(cells) if _col_number(c.get("r", "")) > column), len(cells))
        row.insert(before, cell)
    # Remove any previous formula, inline text or values; keep style and attrs.
    for child in list(cell):
        if child.tag in {f"{M}f", f"{M}v", f"{M}is"}:
            cell.remove(child)
    if isinstance(value, int):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, f"{M}v").text = str(value)
    else:
        cell.set("t", "inlineStr")
        ET.SubElement(ET.SubElement(cell, f"{M}is"), f"{M}t").text = str(value)


def _fill_target(root: ET.Element, cols: tuple[int, ...], receipts: list[dict]) -> bytes:
    data = root.find(f"{M}sheetData")
    if data is None:
        raise ValueError("홈택스 양식의 시트 데이터가 없습니다.")
    # Never remove template rows, worksheet styles or explanatory sheets.
    row_map = {int(row.get("r")): row for row in data.findall(f"{M}row") if row.get("r", "").isdigit()}
    # Clean old data in the export region to prevent stale template data leaking.
    for row_no, row in row_map.items():
        if not 7 <= row_no <= 1006:
            continue
        for cell in row.findall(f"{M}c"):
            if _col_number(cell.get("r", "")) not in cols:
                continue
            for child in list(cell):
                if child.tag in {f"{M}f", f"{M}v", f"{M}is"}:
                    cell.remove(child)
            cell.attrib.pop("t", None)
    for i, receipt in enumerate(receipts, 7):
        row = row_map.get(i)
        if row is None:
            row = ET.Element(f"{M}row", {"r": str(i)})
            for index, old_row in enumerate(data):
                if old_row.tag == f"{M}row" and int(old_row.get("r", "0")) > i:
                    data.insert(index, row)
                    break
            else:
                data.append(row)
            row_map[i] = row
        values = (receipt["purpose"], receipt["number"], receipt["total"],
                  receipt["supply"], receipt["vat"], receipt["memo"])
        for col, value in zip(cols, values):
            _write_cell(row, col, i, value)
    # Keep the ZIP format and all other sheets unmodified.
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _validate_receipt(row: sqlite3.Row) -> dict:
    transaction_id = row["id"]
    purpose, number = row["receipt_purpose_code"], row["receipt_number"]
    supply, vat, total = row["receipt_supply_krw"], row["receipt_vat_krw"], row["amount_krw"]
    if purpose is None or not re.fullmatch(r"\d{1,3}", str(purpose)):
        raise ValueError(f"거래 {transaction_id}: 용도구분 코드가 필요합니다. 공식 양식 [항목설명]을 확인하세요.")
    if not number or "\n" in number or "\r" in number:
        raise ValueError(f"거래 {transaction_id}: 고객의 발급수단번호가 필요합니다.")
    if row["receipt_requested"] and number == "010-000-1234":
        raise ValueError(f"거래 {transaction_id}: 고객 요청 거래에 자진발급 번호를 사용할 수 없습니다.")
    if supply is None or vat is None or supply < 0 or vat < 0 or supply + vat != total:
        raise ValueError(f"거래 {transaction_id}: 공급가액+부가가치세가 입금액 {total}원과 일치하도록 설정하세요.")
    memo = row["receipt_memo"] if row["receipt_memo"] else (
        f"GMT:{row['receipt_order_id']}" if row["receipt_order_id"] else f"TTR:{transaction_id}")
    if len(memo) > 100:
        raise ValueError(f"거래 {transaction_id}: 메모는 100자를 넘길 수 없습니다.")
    return {"id": transaction_id, "purpose": int(purpose), "number": number,
            "total": total, "supply": supply, "vat": vat, "memo": memo}


def _yesterday_seoul() -> str:
    # KST currently has no daylight saving time; fixed offset also works on Windows without tzdata.
    return (datetime.now(timezone(timedelta(hours=9))).date() - timedelta(days=1)).isoformat()


def export_receipts(db_path: Path, template_path: Path, output_path: Path, *,
                    through: str | None = None, retry_exported: bool = False) -> int:
    if not template_path.is_file() or not is_zipfile(template_path):
        raise ValueError("홈택스에서 내려받은 .xlsx 원본 양식을 --template로 지정하세요.")
    if template_path.resolve() == output_path.resolve():
        raise ValueError("원본 템플릿을 덮어쓸 수 없습니다. 출력 경로를 따로 지정하세요.")
    through = through or _yesterday_seoul()
    try:
        datetime.strptime(through, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("--through는 YYYY-MM-DD 형식이어야 합니다.") from exc
    statuses = ("PENDING", "EXPORTED") if retry_exported else ("PENDING",)
    placeholders = ",".join("?" for _ in statuses)
    with _db_connect(db_path) as db:
        db.execute("BEGIN IMMEDIATE")
        rows = db.execute(f"""SELECT * FROM transactions
            WHERE amount_krw > 0 AND is_gmt_sale = 1 AND receipt_requested = 1
              AND receipt_status IN ({placeholders}) AND substr(transaction_at, 1, 10) <= ?
            ORDER BY transaction_at, id""", (*statuses, through)).fetchall()
        if not rows:
            raise ValueError("발급 준비된 GMT 매출이 없습니다. 입금 거래의 GMT 매출 여부와 발급정보를 먼저 등록하세요.")
        if len(rows) > MAX_ROWS:
            raise ValueError(f"일괄발급 최대 {MAX_ROWS}건을 초과했습니다. 날짜별로 나누어 처리하세요.")
        receipts = [_validate_receipt(row) for row in rows]
        with ZipFile(template_path, "r") as source:
            target, root, cols = _find_target(source)
            new_xml = _fill_target(root, cols, receipts)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            fd, temp_path = tempfile.mkstemp(prefix=".receipt_", suffix=".xlsx", dir=output_path.parent)
            os.close(fd)
            try:
                with ZipFile(temp_path, "w") as dest:
                    for entry in source.infolist():
                        dest.writestr(entry, new_xml if entry.filename == target else source.read(entry.filename))
                os.replace(temp_path, output_path)
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        # EXPORTED != ISSUED: only Hometax confirmation can advance to ISSUED.
        exported_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
        db.executemany("""UPDATE transactions SET receipt_status='EXPORTED', receipt_exported_at=?
            WHERE id=?""", [(exported_at, r["id"]) for r in receipts])
    return len(receipts)


def confirm_issued(db_path: Path, transaction_id: int, approval_no: str, issued_on: str) -> None:
    """Record a verified Hometax issuance date, not the time of local confirmation."""
    if not approval_no.strip():
        raise ValueError("홈택스에서 확인한 승인번호를 입력하세요.")
    try:
        if datetime.strptime(issued_on, "%Y-%m-%d").date().isoformat() != issued_on:
            raise ValueError
    except ValueError as exc:
        raise ValueError("홈택스 실제 발급일은 YYYY-MM-DD 형식으로 입력하세요.") from exc
    with _db_connect(db_path) as db:
        cursor = db.execute("""UPDATE transactions
            SET receipt_status='ISSUED', receipt_issued_at=?, receipt_confirmed_at=?, receipt_approval_no=?
            WHERE id=? AND receipt_status='EXPORTED'""",
            (issued_on, datetime.now(timezone.utc).isoformat(timespec="seconds"),
             approval_no.strip(), transaction_id))
        if cursor.rowcount != 1:
            raise ValueError("엑셀로 내보낸 거래만 홈택스 발급 확인 후 ISSUED로 변경할 수 있습니다.")


def list_deposits(db_path: Path) -> None:
    with _db_connect(db_path) as db:
        rows = db.execute("""SELECT id, transaction_at, amount_krw, receipt_requested,
              is_gmt_sale, receipt_status FROM transactions WHERE amount_krw > 0
              ORDER BY transaction_at DESC LIMIT 100""").fetchall()
        print("ID | 거래일시 | 입금액 | 요청 | GMT매출 | 발급상태")
        for r in rows:
            print(f"{r['id']} | {r['transaction_at']} | {r['amount_krw']} | "
                  f"{r['receipt_requested']} | {r['is_gmt_sale']} | {r['receipt_status']}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="GMT 거래의 홈택스 현금영수증 엑셀 작성 (발급 자체는 하지 않음)")
    parser.add_argument("--db", type=Path, default=Path("toss_transactions.db"))
    cmds = parser.add_subparsers(dest="action", required=True)
    cmds.add_parser("list", help="최근 입금 거래 ID와 발급 상태 조회")
    p = cmds.add_parser("prepare", help="거래별 고객/세액 정보를 준비 (추후 GMT Hub 데이터로 대체)")
    p.add_argument("--id", type=int, required=True)
    p.add_argument("--gmt-sale", choices=["yes", "no"])
    p.add_argument("--requested", choices=["yes", "no"])
    p.add_argument("--purpose", help="홈택스 [항목설명] 시트의 용도구분 코드")
    p.add_argument("--number", help="고객이 지정한 발급수단번호")
    p.add_argument("--supply", type=int, help="공급가액(원); 임의로 계산하지 않음")
    p.add_argument("--vat", type=int, help="부가세(원); 임의로 계산하지 않음")
    p.add_argument("--memo")
    p.add_argument("--order-id")
    e = cmds.add_parser("export", help="홈택스 공식 XLSX 양식 기반 일괄 발급용 엑셀 생성")
    e.add_argument("--template", type=Path, required=True)
    e.add_argument("--output", type=Path, required=True)
    e.add_argument("--through", help="거래일자 마감(기본값: 한국시간 어제), YYYY-MM-DD")
    e.add_argument("--retry-exported", action="store_true", help="기존에 내보낸 미확인 거래도 포함 (중복발급 주의)")
    c = cmds.add_parser("confirm", help="홈택스 발급 결과 확인 후에만 발급완료 처리")
    c.add_argument("--id", type=int, required=True)
    c.add_argument("--approval-no", required=True)
    c.add_argument("--issued-on", required=True, help="홈택스 실제 발급일 YYYY-MM-DD")
    args = parser.parse_args(argv)
    try:
        if args.action == "list":
            list_deposits(args.db)
        elif args.action == "prepare":
            prepare_receipt(args.db, args.id, gmt_sale=_yn(args.gmt_sale), requested=_yn(args.requested),
                            purpose=args.purpose, number=args.number, supply=args.supply,
                            vat=args.vat, memo=args.memo, order_id=args.order_id)
            print(f"[OK] 거래 {args.id}번의 발급정보 저장")
        elif args.action == "export":
            count = export_receipts(args.db, args.template, args.output, through=args.through,
                                    retry_exported=args.retry_exported)
            print(f"[OK] {count}건 홈택스 엑셀 작성: {args.output.resolve()}")
            print("[주의] 엑셀 생성만 완료되었습니다. 홈택스 실제 발급 완료가 아닙니다.")
        else:
            confirm_issued(args.db, args.id, args.approval_no, args.issued_on)
            print(f"[OK] 거래 {args.id}번 발급완료 기록")
        return 0
    except (OSError, ValueError, sqlite3.Error, ET.ParseError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
