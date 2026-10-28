"""Import a Toss Bank transaction-history Excel workbook into local SQLite.

- Imports ONE .xlsx file per invocation (password-protected Toss exports supported).
- Preserves signed KRW amounts and all eight transaction columns.
- Deduplicates overlapping exports via full-row fingerprint + occurrence number.
- Changes to the database commit atomically.
"""
from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from getpass import getpass
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile, is_zipfile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REQUIRED = ["거래 일시", "적요", "거래 유형", "거래 기관", "계좌번호", "거래 금액", "거래 후 잔액", "메모"]

CREATE_SQL = """
CREATE TABLE IF NOT EXISTS transactions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source_account TEXT NOT NULL,
  transaction_at TEXT NOT NULL,
  description TEXT,
  transaction_type TEXT,
  institution TEXT,
  counterparty_account TEXT,
  amount_krw INTEGER NOT NULL,
  balance_after_krw INTEGER,
  memo TEXT,
  fingerprint TEXT NOT NULL,
  occurrence_no INTEGER NOT NULL CHECK(occurrence_no > 0),
  first_imported_at TEXT NOT NULL,
  UNIQUE(source_account, fingerprint, occurrence_no)
);
CREATE INDEX IF NOT EXISTS idx_transactions_at ON transactions(transaction_at);
CREATE TABLE IF NOT EXISTS import_runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source_filename TEXT NOT NULL,
  source_file_sha256 TEXT NOT NULL,
  source_account TEXT NOT NULL,
  checked_rows INTEGER NOT NULL,
  inserted_rows INTEGER NOT NULL,
  duplicate_rows INTEGER NOT NULL,
  imported_at TEXT NOT NULL
);
"""


def _cell_col(cell_ref: str) -> int:
    match = re.match(r"[A-Z]+", cell_ref.upper())
    if not match:
        raise ValueError(f"잘못된 셀 주소: {cell_ref!r}")
    value = 0
    for letter in match.group():
        value = value * 26 + ord(letter) - 64
    return value - 1


def _text(value: object) -> str | None:
    if value is None:
        return None
    stripped = str(value).strip()
    return stripped or None


def _get_excel_password() -> str:
    """Load the local .env without overriding a separately configured environment variable."""
    try:
        from dotenv import load_dotenv
    except ImportError as exc:
        raise RuntimeError(".env 사용을 위해 'pip install -r requirements.txt'를 실행하세요.") from exc
    load_dotenv(dotenv_path=Path(__file__).resolve().parent / ".env", override=False)
    return os.environ.get("TOSS_EXCEL_PASSWORD") or getpass("엑셀 암호 (화면에 표시되지 않음): ")


def _read_excel_zip(input_path: Path) -> bytes:
    """Return raw XLSX bytes; decrypt only if the file is Office-encrypted."""
    data = input_path.read_bytes()
    if is_zipfile(BytesIO(data)):
        return data
    if not data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        raise ValueError("지원하지 않는 파일입니다. 토스뱅크 .xlsx 파일을 사용하세요.")
    try:
        import msoffcrypto
    except ImportError as exc:
        raise RuntimeError("암호화된 엑셀은 'pip install -r requirements.txt' 실행 후 가져올 수 있습니다.") from exc
    password = _get_excel_password()
    if not password:
        raise ValueError("암호가 입력되지 않았습니다.")
    source = BytesIO(data)
    output = BytesIO()
    try:
        office = msoffcrypto.OfficeFile(source)
        office.load_key(password=password)
        office.decrypt(output)
    except Exception as exc:
        raise ValueError("엑셀 복호화에 실패했습니다. 암호와 파일을 확인하세요.") from exc
    finally:
        password = None
    result = output.getvalue()
    if not is_zipfile(BytesIO(result)):
        raise ValueError("복호화 결과가 유효한 .xlsx 파일이 아닙니다.")
    return result


def _parse_sheet(data: bytes) -> list[dict[int, object]]:
    """Read the first worksheet from an OOXML .xlsx, no workbook modification."""
    with ZipFile(BytesIO(data)) as zf:
        names = set(zf.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            shared = ["".join(el.itertext()) for el in root.findall("m:si", NS)]
        # Toss Bank's exported transaction history is a single-sheet workbook.
        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in names:
            raise ValueError("첫 번째 워크시트를 찾을 수 없습니다.")
        root = ET.fromstring(zf.read(sheet_name))
        rows: list[dict[int, object]] = []
        for row in root.findall(".//m:sheetData/m:row", NS):
            result: dict[int, object] = {}
            for cell in row.findall("m:c", NS):
                col = _cell_col(cell.attrib.get("r", ""))
                cell_type = cell.attrib.get("t")
                value = cell.find("m:v", NS)
                if cell_type == "inlineStr":
                    inline = cell.find("m:is", NS)
                    parsed = "".join(inline.itertext()) if inline is not None else None
                elif value is None:
                    parsed = None
                elif cell_type == "s":
                    parsed = shared[int(value.text)]
                elif cell_type in {"str", "e"}:
                    parsed = value.text
                elif value.text is None:
                    parsed = None
                else:
                    try:
                        parsed = Decimal(value.text)
                    except InvalidOperation:
                        parsed = value.text
                result[col] = parsed
            rows.append(result)
        return rows


def _integer(value: object, *, field: str, row: int, nullable: bool) -> int | None:
    if value is None or _text(value) is None:
        if nullable:
            return None
        raise ValueError(f"{row}행 {field} 값이 없습니다.")
    try:
        number = Decimal(str(value).replace(",", "").replace("원", "").strip())
    except InvalidOperation as exc:
        raise ValueError(f"{row}행 {field} 숫자 형식 오류") from exc
    if not number.is_finite() or number != number.to_integral_value():
        raise ValueError(f"{row}행 {field}은 정수 원 단위여야 합니다.")
    return int(number)


def _date(value: object, row: int) -> str:
    raw = _text(value)
    if not raw:
        raise ValueError(f"{row}행 거래 일시가 없습니다.")
    for pattern in ("%Y.%m.%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(raw, pattern).isoformat(timespec="seconds")
        except ValueError:
            continue
    raise ValueError(f"{row}행 거래 일시 형식을 해석하지 못했습니다: {raw!r}")


def parse_toss_excel(path: Path) -> tuple[str, list[dict[str, object]]]:
    rows = _parse_sheet(_read_excel_zip(path))
    header_index = -1
    columns: dict[str, int] = {}
    for i, row in enumerate(rows):
        locations = {_text(v): k for k, v in row.items() if _text(v)}
        if all(name in locations for name in REQUIRED):
            header_index, columns = i, {name: locations[name] for name in REQUIRED}
            break
    if header_index == -1:
        raise ValueError("토스뱅크 거래내역 8개 컬럼을 찾지 못했습니다.")
    source_account = None
    for row in rows[:header_index]:
        if any(_text(v) == "계좌번호" for v in row.values()):
            key = next(k for k, v in row.items() if _text(v) == "계좌번호")
            source_account = _text(row.get(key + 1))
            break
    if not source_account:
        raise ValueError("원본 계좌 식별정보를 찾지 못했습니다. 계좌별 거래가 섞이지 않도록 가져오기를 중단합니다.")

    records: list[dict[str, object]] = []
    for index, row in enumerate(rows[header_index + 1:], header_index + 2):
        values = {name: row.get(column) for name, column in columns.items()}
        if all(_text(value) is None for value in values.values()):
            continue
        record = {
            "transaction_at": _date(values["거래 일시"], index),
            "description": _text(values["적요"]),
            "transaction_type": _text(values["거래 유형"]),
            "institution": _text(values["거래 기관"]),
            "counterparty_account": _text(values["계좌번호"]),
            "amount_krw": _integer(values["거래 금액"], field="거래 금액", row=index, nullable=False),
            "balance_after_krw": _integer(values["거래 후 잔액"], field="거래 후 잔액", row=index, nullable=True),
            "memo": _text(values["메모"]),
        }
        records.append(record)
    if not records:
        raise ValueError("거래내역이 없습니다.")
    return source_account, records


def _fingerprint(record: dict[str, object]) -> str:
    fields = (
        "transaction_at", "description", "transaction_type", "institution",
        "counterparty_account", "amount_krw", "balance_after_krw", "memo",
    )
    canonical = json.dumps([record[key] for key in fields], ensure_ascii=False, separators=(",", ":"))
    return sha256(canonical.encode("utf-8")).hexdigest()


def import_excel(excel_path: Path, db_path: Path) -> tuple[int, int, int]:
    if not excel_path.is_file():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
    source_account, records = parse_toss_excel(excel_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    imported_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    seen: dict[str, int] = defaultdict(int)
    inserted = 0
    duplicates = 0
    with sqlite3.connect(db_path, timeout=30) as db:
        db.execute("PRAGMA busy_timeout = 30000")
        db.executescript(CREATE_SQL)
        db.execute("BEGIN IMMEDIATE")
        try:
            for record in records:
                fingerprint = _fingerprint(record)
                seen[fingerprint] += 1
                cursor = db.execute(
                    """INSERT INTO transactions (
                      source_account, transaction_at, description, transaction_type,
                      institution, counterparty_account, amount_krw, balance_after_krw,
                      memo, fingerprint, occurrence_no, first_imported_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(source_account, fingerprint, occurrence_no) DO NOTHING""",
                    (source_account, record["transaction_at"], record["description"],
                     record["transaction_type"], record["institution"],
                     record["counterparty_account"], record["amount_krw"],
                     record["balance_after_krw"], record["memo"], fingerprint,
                     seen[fingerprint], imported_at),
                )
                inserted += int(cursor.rowcount == 1)
                duplicates += int(cursor.rowcount == 0)
            db.execute(
                """INSERT INTO import_runs
                (source_filename, source_file_sha256, source_account, checked_rows,
                 inserted_rows, duplicate_rows, imported_at)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (excel_path.name, sha256(excel_path.read_bytes()).hexdigest(), source_account,
                 len(records), inserted, duplicates, imported_at),
            )
            db.commit()
        except Exception:
            db.rollback()
            raise
    return len(records), inserted, duplicates


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="토스뱅크 거래내역 엑셀 1개를 SQLite에 중복 없이 저장합니다.")
    parser.add_argument("excel", type=Path, help="토스뱅크에서 내려받은 .xlsx 파일")
    parser.add_argument("--db", type=Path, default=Path("toss_transactions.db"), help="SQLite 경로 (기본: ./toss_transactions.db)")
    args = parser.parse_args(argv)
    try:
        checked, inserted, duplicate = import_excel(args.excel, args.db)
    except (OSError, ValueError, RuntimeError, BadZipFile, sqlite3.Error) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    print(f"[OK] 전체 {checked}건 | 신규 {inserted}건 | 중복 제외 {duplicate}건")
    print(f"[DB] {args.db.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
