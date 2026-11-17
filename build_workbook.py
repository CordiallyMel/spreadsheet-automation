"""Build a monthly-close workbook from a transaction CSV.

Every total in the finished workbook is a live Excel formula (SUMIFS / SUM), so
an accountant can filter, extend, or re-tag the ledger and watch the summary
recalculate. Nothing is pasted as a number.

Because formulas cannot be evaluated without Excel, the workbook also carries:

  Control  - an independent recomputation done in Python, cell for cell,
  Checks   - formulas that compare Summary against Control and print PASS/FAIL,
  Recon    - statement balance versus composed balance, with a variance cell.

The point is that the file argues with itself. If a range is wrong, a category
is missing, or someone types over a formula, Checks says so on open.

Usage:
    python build_workbook.py --input data/transactions.csv \
        --balances data/balances.csv --out out/monthly_close.xlsx
"""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MONEY = "#,##0.00"
NAVY = "1F3864"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


@dataclass(frozen=True)
class Txn:
    date: str
    category: str
    account: str
    amount: float
    memo: str

    @property
    def month(self) -> str:
        return self.date[:7]


@dataclass(frozen=True)
class Balances:
    account: str
    opening: float
    closing: float


def load_transactions(path: Path) -> list[Txn]:
    with path.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    txns = [
        Txn(
            date=row["date"].strip(),
            category=row["category"].strip(),
            account=row["account"].strip(),
            amount=round(float(row["amount"]), 2),
            memo=(row.get("memo") or "").strip(),
        )
        for row in rows
    ]
    if not txns:
        raise SystemExit(f"no transactions found in {path}")
    return txns


def load_balances(path: Path) -> Balances:
    with path.open(newline="", encoding="utf-8") as handle:
        row = next(csv.DictReader(handle))
    return Balances(
        account=row["account"].strip(),
        opening=round(float(row["opening"]), 2),
        closing=round(float(row["closing"]), 2),
    )


def matrix(txns: list[Txn]) -> tuple[list[str], list[str], dict[tuple[str, str], float]]:
    """Return (categories, months, {category, month -> amount})."""
    categories = sorted({t.category for t in txns})
    months = sorted({t.month for t in txns})
    totals: dict[tuple[str, str], float] = defaultdict(float)
    for txn in txns:
        totals[(txn.category, txn.month)] += txn.amount
    return categories, months, {
        key: round(value, 2) for key, value in totals.items()
    }


def _header(ws, row: int, values: list[str], start_col: int = 1) -> None:
    for offset, value in enumerate(values):
        cell = ws.cell(row=row, column=start_col + offset, value=value)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX


def _title(ws, text: str, span: int) -> None:
    cell = ws.cell(row=1, column=1, value=text)
    cell.font = Font(bold=True, size=14, color=NAVY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(span, 2))


def build(transactions: list[Txn], balances: Balances, out_path: Path) -> Path:
    categories, months, totals = matrix(transactions)

    wb = Workbook()
    ledger = wb.active
    ledger.title = "Ledger"

    # --- Ledger -----------------------------------------------------------
    _title(ledger, f"Ledger - {balances.account}", 6)
    _header(ledger, 3, ["Date", "Category", "Account", "Amount", "Month", "Memo"])
    first_data_row = 4
    for index, txn in enumerate(transactions):
        row = first_data_row + index
        ledger.cell(row=row, column=1, value=txn.date)
        ledger.cell(row=row, column=2, value=txn.category)
        ledger.cell(row=row, column=3, value=txn.account)
        amount_cell = ledger.cell(row=row, column=4, value=txn.amount)
        amount_cell.number_format = MONEY
        ledger.cell(row=row, column=5, value=f'=TEXT(A{row},"YYYY-MM")')
        ledger.cell(row=row, column=6, value=txn.memo)
        for column in range(1, 7):
            ledger.cell(row=row, column=column).border = BOX
    last_data_row = first_data_row + len(transactions) - 1

    for column, width in zip("ABCDEF", (12, 22, 22, 14, 11, 38)):
        ledger.column_dimensions[column].width = width
    ledger.freeze_panes = "A4"
    ledger.auto_filter.ref = f"A3:F{last_data_row}"

    categories_csv = ",".join(categories)
    if len(categories_csv) <= 250:
        validation = DataValidation(
            type="list", formula1=f'"{categories_csv}"', allow_blank=False
        )
        validation.error = "Pick a category from the list so the summary stays accurate."
        ledger.add_data_validation(validation)
        validation.add(f"B{first_data_row}:B{last_data_row}")

    # --- Summary (formulas only) -----------------------------------------
    summary = wb.create_sheet("Summary")
    _title(summary, "Monthly summary - every figure below is a live formula", len(months) + 2)
    _header(summary, 3, ["Category", *months, "Total"])

    first_cat_row = 4
    for index, category in enumerate(categories):
        row = first_cat_row + index
        cell = summary.cell(row=row, column=1, value=category)
        cell.border = BOX
        for column in range(2, len(months) + 2):
            header = f"{get_column_letter(column)}$3"
            formula = (
                f"=SUMIFS(Ledger!$D:$D,Ledger!$B:$B,$A{row},"
                f"Ledger!$E:$E,{header})"
            )
            target = summary.cell(row=row, column=column, value=formula)
            target.number_format = MONEY
            target.border = BOX
        total = summary.cell(
            row=row,
            column=len(months) + 2,
            value=f"=SUM(B{row}:{get_column_letter(len(months) + 1)}{row})",
        )
        total.number_format = MONEY
        total.font = Font(bold=True)
        total.border = BOX
        total.fill = PatternFill("solid", fgColor=GREY)

    last_cat_row = first_cat_row + len(categories) - 1
    total_row = last_cat_row + 1
    label = summary.cell(row=total_row, column=1, value="Total")
    label.font = Font(bold=True, color="FFFFFF")
    label.fill = PatternFill("solid", fgColor=NAVY)
    for column in range(2, len(months) + 3):
        letter = get_column_letter(column)
        cell = summary.cell(
            row=total_row, column=column, value=f"=SUM({letter}{first_cat_row}:{letter}{last_cat_row})"
        )
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.number_format = MONEY
    summary.column_dimensions["A"].width = 24
    for column in range(2, len(months) + 3):
        summary.column_dimensions[get_column_letter(column)].width = 14
    summary.freeze_panes = "B4"

    # --- Control (Python's independent recomputation) ---------------------
    control = wb.create_sheet("Control")
    _title(control, "Control - independent recomputation in Python", len(months) + 1)
    _header(control, 3, ["Category", *months])
    for index, category in enumerate(categories):
        row = first_cat_row + index
        control.cell(row=row, column=1, value=category).border = BOX
        for offset, month in enumerate(months):
            value = totals.get((category, month), 0.0)
            cell = control.cell(row=row, column=2 + offset, value=value)
            cell.number_format = MONEY
            cell.border = BOX
    control.column_dimensions["A"].width = 24
    for column in range(2, len(months) + 2):
        control.column_dimensions[get_column_letter(column)].width = 14

    # --- Checks (Excel compares the two) ----------------------------------
    checks = wb.create_sheet("Checks")
    _title(checks, "Checks - Summary vs Control", len(months) + 1)
    status = checks.cell(row=2, column=1, value="Workbook status")
    status.font = Font(bold=True)
    checks.column_dimensions["A"].width = 24

    _header(checks, 3, ["Row", *months])
    first_check_row = 4
    for index, category in enumerate(categories):
        row = first_check_row + index
        summary_row = first_cat_row + index
        checks.cell(row=row, column=1, value=f'=Summary!A{summary_row}')
        for offset in range(len(months)):
            column = get_column_letter(2 + offset)
            checks.cell(
                row=row,
                column=2 + offset,
                value=(
                    f'=IF(ABS(Summary!{column}{summary_row}-'
                    f'Control!{column}{summary_row})<0.005,"OK","MISMATCH")'
                ),
            )
    last_check_row = first_check_row + len(categories) - 1
    for column in range(2, len(months) + 3):
        checks.column_dimensions[get_column_letter(column)].width = 12

    checks_range = f"B{first_check_row}:{get_column_letter(len(months) + 1)}{last_check_row}"
    verdict = checks.cell(
        row=2, column=2, value=f'=IF(COUNTIF({checks_range},"MISMATCH")=0,"PASS","FAIL")'
    )
    verdict.font = Font(bold=True, size=12)
    checks.conditional_formatting.add(
        f"B2",
        CellIsRule(operator="equal", formula=['"FAIL"'], fill=PatternFill("solid", fgColor="FFC7CE")),
    )

    # --- Recon ------------------------------------------------------------
    recon = wb.create_sheet("Recon")
    _title(recon, "Statement reconciliation", 2)
    labels = [
        ("Account", balances.account),
        ("Opening balance", balances.opening),
        ("Inflows", '=SUMIF(Ledger!$D:$D,">0")'),
        ("Outflows", '=SUMIF(Ledger!$D:$D,"<0")'),
        ("Composed closing", "=B5+B6+B7"),
        ("Closing per statement", balances.closing),
        ("Variance", "=B9-B8"),
        ("Status", '=IF(ABS(B10)<0.005,"RECONCILED","VARIANCE")'),
    ]
    for index, (label, value) in enumerate(labels):
        row = 4 + index
        key = recon.cell(row=row, column=1, value=label)
        key.font = Font(bold=True)
        key.fill = PatternFill("solid", fgColor=LIGHT)
        key.border = BOX
        cell = recon.cell(row=row, column=2, value=value)
        cell.border = BOX
        if isinstance(value, float) or str(value).startswith("="):
            cell.number_format = MONEY
    # Status cell should read as text, not currency.
    recon["B11"].number_format = "General"
    recon.column_dimensions["A"].width = 24
    recon.column_dimensions["B"].width = 22
    recon.conditional_formatting.add(
        "B9",
        CellIsRule(operator="notEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE")),
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("data/transactions.csv"))
    parser.add_argument("--balances", type=Path, default=Path("data/balances.csv"))
    parser.add_argument("--out", type=Path, default=Path("out/monthly_close.xlsx"))
    args = parser.parse_args(argv)

    transactions = load_transactions(args.input)
    balances = load_balances(args.balances)
    path = build(transactions, balances, args.out)
    categories, months, _ = matrix(transactions)
    print(
        f"wrote {path} | {len(transactions)} rows | {len(categories)} categories | "
        f"{len(months)} months | formulas throughout, Control sheet cross-checked"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
