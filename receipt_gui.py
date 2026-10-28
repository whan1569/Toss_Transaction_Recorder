"""Small Windows-friendly Tkinter UI for preparing Hometax Excel exports.

This app does NOT log into Hometax or issue cash receipts.
"""
from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from receipt_export import _db_connect, _yesterday_seoul, export_receipts, prepare_receipt, reset_unissued


class ReceiptApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("GMT | 현금영수증 엑셀 준비")
        self.geometry("1060x715")
        self.minsize(840, 600)
        self.db_path = tk.StringVar(value=str(Path("toss_transactions.db").resolve()))
        self.template_path = tk.StringVar()
        self.through = tk.StringVar(value=_yesterday_seoul())
        self.sale = tk.BooleanVar(value=False)
        self.requested = tk.BooleanVar(value=True)
        self.purpose = tk.StringVar()
        self.number = tk.StringVar()
        self.supply = tk.StringVar()
        self.vat = tk.StringVar()
        self.order_id = tk.StringVar()
        self.memo = tk.StringVar()
        self.selected_id: int | None = None
        self._build()
        if Path(self.db_path.get()).is_file():
            self.reload()

    def _build(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        file_frame = ttk.LabelFrame(root, text="거래 데이터", padding=8)
        file_frame.pack(fill="x")
        ttk.Entry(file_frame, textvariable=self.db_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="DB 선택", command=self.choose_db).pack(side="left", padx=6)
        ttk.Button(file_frame, text="새로고침", command=self.reload).pack(side="left")

        ttk.Label(root, text="GMT 매출 여부는 수동으로 확인해야 합니다. 현금영수증 요청 여부는 현재 기본값 '예'이며, 나중에 GMT Hub에서 동기화할 수 있습니다.",
                  wraplength=970).pack(fill="x", pady=(9, 5))
        table_frame = ttk.Frame(root)
        table_frame.pack(fill="both", expand=True)
        cols = ("id", "at", "amount", "sale", "requested", "status", "ready")
        self.table = ttk.Treeview(table_frame, columns=cols, show="headings", height=11, selectmode="browse")
        labels = ("거래 ID", "토스 입금일시", "입금액(원)", "GMT 매출", "발급 요청", "진행 상태", "필수정보")
        widths = (70, 170, 130, 110, 110, 120, 130)
        for col, label, width in zip(cols, labels, widths):
            self.table.heading(col, text=label)
            self.table.column(col, width=width, minwidth=60, anchor="center")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scrollbar.set)
        self.table.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.table.bind("<<TreeviewSelect>>", self._on_select)

        edit = ttk.LabelFrame(root, text="선택한 거래의 발급정보 (홈택스 제출 전 확인)", padding=10)
        edit.pack(fill="x", pady=(12, 8))
        self.active_label = ttk.Label(edit, text="거래를 선택하세요.")
        self.active_label.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 6))
        ttk.Checkbutton(edit, text="GMT 매출 확인", variable=self.sale).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(edit, text="현금영수증 요청", variable=self.requested).grid(row=1, column=1, sticky="w")
        entries = [
            ("용도구분", self.purpose, 2, 0), ("발급수단번호", self.number, 2, 2),
            ("공급가액", self.supply, 3, 0), ("부가가치세", self.vat, 3, 2),
            ("GMT 주문 ID (선택)", self.order_id, 4, 0), ("메모 (선택)", self.memo, 4, 2),
        ]
        for title, var, row, col in entries:
            ttk.Label(edit, text=title).grid(row=row, column=col, sticky="w", padx=(0, 7), pady=3)
            ttk.Entry(edit, textvariable=var, width=24).grid(row=row, column=col+1, sticky="ew", pady=3, padx=(0, 10))
        edit.columnconfigure(1, weight=1)
        edit.columnconfigure(3, weight=1)
        ttk.Button(edit, text="미발급 거래 재작업", command=self.reset_export).grid(row=4, column=4, sticky="e")
        ttk.Button(edit, text="선택 거래 저장", command=self.save_receipt).grid(row=4, column=5, sticky="e", padx=(15, 0))
        ttk.Label(edit, text="세액은 자동 계산하지 않습니다. 홈택스 공식 양식과 사업자의 과세기준을 확인한 금액만 입력하세요.",
                  wraplength=970).grid(row=5, column=0, columnspan=6, sticky="w", pady=(9, 0))

        output = ttk.LabelFrame(root, text="홈택스 제출용 엑셀 생성", padding=8)
        output.pack(fill="x")
        ttk.Entry(output, textvariable=self.template_path).grid(row=0, column=0, columnspan=4, sticky="ew")
        ttk.Button(output, text="홈택스 공식 양식 선택", command=self.choose_template).grid(row=0, column=4, padx=(7, 0))
        ttk.Label(output, text="거래일자 마감").grid(row=1, column=0, sticky="w", pady=(9, 0))
        ttk.Entry(output, textvariable=self.through, width=14).grid(row=1, column=1, sticky="w", pady=(9, 0))
        ttk.Label(output, text="(기본값: 한국시간 어제까지의 미출력·발급요청 거래)").grid(row=1, column=2, columnspan=2, sticky="w", padx=(8, 0), pady=(9, 0))
        ttk.Button(output, text="제출용 엑셀 만들기", command=self.export).grid(row=1, column=4, sticky="e", pady=(9, 0))
        output.columnconfigure(0, weight=1)
        output.columnconfigure(2, weight=1)
        ttk.Label(root, text="주의: 파일 생성은 현금영수증 발급이 아닙니다. 홈택스에서 파일검증 → 일괄발급 → 결과확인까지 별도로 진행해야 합니다.",
                  wraplength=980).pack(fill="x", pady=(8, 0))

    def choose_db(self) -> None:
        result = filedialog.askopenfilename(filetypes=[("SQLite database", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")])
        if result:
            self.db_path.set(result)
            self.reload()

    def choose_template(self) -> None:
        result = filedialog.askopenfilename(filetypes=[("Excel Workbook", "*.xlsx")])
        if result:
            self.template_path.set(result)

    def reload(self) -> None:
        self.selected_id = None
        for item in self.table.get_children():
            self.table.delete(item)
        try:
            with _db_connect(Path(self.db_path.get())) as db:
                rows = db.execute("""SELECT id,transaction_at,amount_krw,receipt_requested,
                    is_gmt_sale,receipt_status,receipt_purpose_code,receipt_number,
                    receipt_supply_krw,receipt_vat_krw FROM transactions
                    WHERE amount_krw > 0 ORDER BY transaction_at DESC LIMIT 1000""").fetchall()
            for row in rows:
                ready = (row["receipt_purpose_code"] is not None and bool(row["receipt_number"])
                         and row["receipt_supply_krw"] is not None and row["receipt_vat_krw"] is not None
                         and row["receipt_supply_krw"] + row["receipt_vat_krw"] == row["amount_krw"])
                self.table.insert("", "end", iid=str(row["id"]), values=(
                    row["id"], row["transaction_at"], f"{row['amount_krw']:,}",
                    "확인" if row["is_gmt_sale"] == 1 else "미확인" if row["is_gmt_sale"] is None else "아님",
                    "예" if row["receipt_requested"] else "아니오", row["receipt_status"],
                    "완료" if ready else "미입력"))
        except Exception as exc:
            # Tk callbacks must not crash the event loop.
            messagebox.showerror("DB 조회 실패", str(exc))

    def _on_select(self, _event) -> None:
        selected = self.table.selection()
        if not selected:
            return
        self.selected_id = int(selected[0])
        try:
            with _db_connect(Path(self.db_path.get())) as db:
                row = db.execute("SELECT * FROM transactions WHERE id=?", (self.selected_id,)).fetchone()
            if row is None:
                return
            self.active_label.config(text=f"선택 거래 ID {self.selected_id} | 입금 {row['amount_krw']:,}원 | {row['transaction_at']}")
            self.sale.set(row["is_gmt_sale"] == 1)
            self.requested.set(bool(row["receipt_requested"]))
            for var, column in ((self.purpose,"receipt_purpose_code"),(self.number,"receipt_number"),
                                (self.supply,"receipt_supply_krw"),(self.vat,"receipt_vat_krw"),
                                (self.order_id,"receipt_order_id"),(self.memo,"receipt_memo")):
                var.set("" if row[column] is None else str(row[column]))
        except (OSError, ValueError) as exc:
            messagebox.showerror("거래 조회 실패", str(exc))

    def save_receipt(self) -> None:
        if self.selected_id is None:
            messagebox.showwarning("거래 선택", "거래를 먼저 선택하세요.")
            return
        try:
            kwargs = dict(gmt_sale=int(self.sale.get()), requested=int(self.requested.get()))
            for field, var in (("purpose",self.purpose),("number",self.number),
                               ("memo",self.memo),("order_id",self.order_id)):
                kwargs[field] = var.get().strip()
            for field, var in (("supply",self.supply),("vat",self.vat)):
                if var.get().strip():
                    kwargs[field] = int(var.get().replace(",", ""))
            prepare_receipt(Path(self.db_path.get()), self.selected_id, **kwargs)
            self.reload()
            messagebox.showinfo("저장 완료", "발급정보를 DB에 저장했습니다. 아직 현금영수증을 발급하지 않았습니다.")
        except (OSError, ValueError) as exc:
            messagebox.showerror("저장 실패", str(exc))

    def reset_export(self) -> None:
        if self.selected_id is None:
            messagebox.showwarning("거래 선택", "거래를 먼저 선택하세요.")
            return
        if not messagebox.askyesno("미발급 확인", "홈택스에서 이 거래의 현금영수증이 실제로 발급되지 않았음을 확인했나요?\n발급 완료된 거래를 다시 제출하면 중복발급이 발생할 수 있습니다."):
            return
        reason = simpledialog.askstring("재작업 사유", "홈택스 검증 실패 등 미발급 재작업 사유를 입력하세요.", parent=self)
        if not reason:
            return
        try:
            reset_unissued(Path(self.db_path.get()), self.selected_id, reason)
            self.reload()
            messagebox.showinfo("재작업 준비", "미발급 확인 거래를 PENDING 상태로 되돌렸습니다. 발급정보를 수정한 뒤 다시 내보내세요.")
        except (OSError, ValueError) as exc:
            messagebox.showerror("재작업 실패", str(exc))

    def export(self) -> None:
        if not self.template_path.get():
            messagebox.showwarning("양식 선택", "홈택스에서 내려받은 원본 .xlsx 양식을 먼저 선택하세요.")
            return
        output = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=f"현금영수증_제출용_{self.through.get().replace('-', '')}.xlsx")
        if not output:
            return
        try:
            count = export_receipts(Path(self.db_path.get()), Path(self.template_path.get()),
                                    Path(output), through=self.through.get())
            self.reload()
            messagebox.showinfo("제출용 엑셀 생성", f"{count}건을 저장했습니다.\n{output}\n\n홈택스 발급은 아직 진행되지 않았습니다.")
        except (OSError, ValueError) as exc:
            messagebox.showerror("엑셀 생성 실패", str(exc))


if __name__ == "__main__":
    ReceiptApp().mainloop()
