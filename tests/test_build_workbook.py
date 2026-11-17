"""Tests for the monthly-close builder.

The workbook is formula driven, so these tests verify what a test can verify
honestly without Excel evaluating the file:

  * the totals really are formulas, not pasted numbers,
  * the Control sheet agrees with an independent recomputation from the CSV,
  * the reconciliation variance is a live formula keyed to the statement figure,
  * a deliberately broken statement balance moves the statement cell, not the
    composed closing.

The workbook's own Checks sheet does the value comparison inside Excel, where a
formula can actually be calculated.
"""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

import pytest
from openpyxl import load_workbook

import build_workbook as bw

HERE = Path(__file__).resolve().parents[1]
TXN = HERE / "data" / "transactions.csv"
BAL = HERE / "data" / "balances.csv"


def independent_totals() -> tuple[list[str], list[str], dict[tuple[str, str], float]]:
    """Recompute the summary straight from the CSV, without build_workbook."""
    per_pair: dict[tuple[str, str], float] = defaultdict(float)
    categories: set[str] = set()
    months: set[str] = set()
    with TXN.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            month = row["date"][:7]
            category = row["category"]
            per_pair[(category, month)] += float(row["amount"])
            categories.add(category)
            months.add(month)
    return (
        sorted(categories),
        sorted(months),
        {key: round(value, 2) for key, value in per_pair.items()},
    )


@pytest.fixture(scope="module")
def workbook(tmp_path_factory: pytest.TempPathFactory):
    out = tmp_path_factory.mktemp("wb") / "monthly_close.xlsx"
    bw.build(bw.load_transactions(TXN), bw.load_balances(BAL), out)
    return load_workbook(out)


def test_summary_totals_are_formulas_not_values(workbook):
    summary = workbook["Summary"]
    categories, months, _ = independent_totals()

    for index in range(len(categories)):
        row = 4 + index
        for offset in range(len(months)):
            cell = summary.cell(row=row, column=2 + offset)
            assert isinstance(cell.value, str) and cell.value.startswith("=SUMIFS("), (
                f"Summary {cell.coordinate} should be a SUMIFS formula, got {cell.value!r}"
            )
        total_cell = summary.cell(row=row, column=2 + len(months))
        assert str(total_cell.value).startswith("=SUM(")


def test_summary_formulas_point_at_ledger_columns(workbook):
    formula = workbook["Summary"].cell(row=4, column=2).value
    assert "Ledger!$D:$D" in formula  # amount
    assert "Ledger!$B:$B" in formula  # category
    assert "Ledger!$E:$E" in formula  # month


def test_control_sheet_matches_independent_recomputation(workbook):
    control = workbook["Control"]
    categories, months, totals = independent_totals()

    for index, category in enumerate(categories):
        row = 4 + index
        assert control.cell(row=row, column=1).value == category
        for offset, month in enumerate(months):
            expected = totals.get((category, month), 0.0)
            actual = control.cell(row=row, column=2 + offset).value
            assert actual == pytest.approx(expected, abs=0.005)


def test_control_grand_total_equals_ledger_total(workbook):
    control = workbook["Control"]
    categories, months, _ = independent_totals()
    grand = sum(
        control.cell(row=4 + index, column=2 + offset).value
        for index in range(len(categories))
        for offset in range(len(months))
    )
    with TXN.open(newline="", encoding="utf-8") as handle:
        ledger_total = sum(float(row["amount"]) for row in csv.DictReader(handle))
    assert grand == pytest.approx(round(ledger_total, 2), abs=0.01)


def test_ledger_row_count_matches_input(workbook):
    ledger = workbook["Ledger"]
    with TXN.open(newline="", encoding="utf-8") as handle:
        expected = sum(1 for _ in csv.DictReader(handle))
    seen = 0
    row = 4
    while ledger.cell(row=row, column=1).value:
        seen += 1
        row += 1
    assert seen == expected


def test_ledger_month_column_is_a_formula(workbook):
    assert workbook["Ledger"].cell(row=4, column=5).value == '=TEXT(A4,"YYYY-MM")'


def test_checks_sheet_compares_summary_to_control(workbook):
    checks = workbook["Checks"]
    assert checks["B2"].value.startswith("=IF(COUNTIF(")
    first_check = checks.cell(row=4, column=2).value
    assert "Summary!" in first_check and "Control!" in first_check


def test_reconciliation_is_balanced_for_the_committed_data(workbook):
    recon = workbook["Recon"]
    assert recon["B8"].value == "=B5+B6+B7"
    assert recon["B9"].value == pytest.approx(bw.load_balances(BAL).closing)
    assert recon["B10"].value == "=B9-B8"
    assert recon["B11"].value == '=IF(ABS(B10)<0.005,"RECONCILED","VARIANCE")'


def test_broken_statement_balance_moves_only_the_statement_cell(tmp_path: Path):
    balances = bw.load_balances(BAL)
    altered = bw.Balances(
        account=balances.account,
        opening=balances.opening,
        closing=round(balances.closing + 125.00, 2),
    )
    out = tmp_path / "broken.xlsx"
    bw.build(bw.load_transactions(TXN), altered, out)
    recon = load_workbook(out)["Recon"]

    assert recon["B5"].value == pytest.approx(altered.opening)
    assert recon["B9"].value == pytest.approx(balances.closing + 125.00)
    assert recon["B10"].value == "=B9-B8"
