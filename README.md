# Monthly close, built by script — not typed by hand

A ledger in, a reviewed workbook out. This tool turns a transaction CSV into the
monthly close pack a small finance team actually signs off on, and it does one
thing differently from most spreadsheet generators:

**every figure is a live formula, and the file checks itself.**

```bash
python data/make_sample_data.py          # deterministic sample ledger
python build_workbook.py                 # -> out/monthly_close.xlsx
python -m pytest -q                      # 9 passed
```

```
wrote out/monthly_close.xlsx | 61 rows | 10 categories | 4 months
```

## Why formulas instead of pasted numbers

The usual failure mode of "automated" reporting is a workbook full of values
that nobody can trace. The first time a reviewer disagrees with a number, the
file is dead — there is no path from the total back to the rows.

Here the summary is `SUMIFS` against the ledger, the month column is
`=TEXT(A4,"YYYY-MM")`, and the reconciliation composes itself from the flows.
An accountant can add rows, re-tag a category from a dropdown, or filter the
ledger, and the summary recalculates on its own. The script's job is to lay out
structure and formatting, not to pre-chew the arithmetic.

## The workbook

| Sheet | What it holds | How it is built |
|---|---|---|
| `Ledger` | one row per transaction, filtered table, frozen header | written by Python |
| `Summary` | category × month grid + totals | **`SUMIFS` formulas only** |
| `Control` | the same grid, recomputed independently in Python | written by Python |
| `Checks` | `OK` / `MISMATCH` per cell, plus a single `PASS` / `FAIL` | **formulas comparing the two sheets** |
| `Recon` | opening, inflows, outflows, composed closing, statement closing, variance | formulas |

`Checks` is the part worth showing a reviewer:

```
=IF(ABS(Summary!B4-Control!B4)<0.005,"OK","MISMATCH")
=IF(COUNTIF(B4:E13,"MISMATCH")=0,"PASS","FAIL")
```

Two independent derivations of the same numbers, compared inside the file, with
the verdict on screen the moment it opens. If someone types over a formula or a
range slips by a row, it says `FAIL` instead of quietly reporting a number.

`Recon` closes the loop against reality: opening balance plus ledger flows must
equal the statement. The variance cell turns red when it does not, and
`Status` reads `RECONCILED` or `VARIANCE`.

## Sample output

Committed demo data: 61 transactions, 10 categories, 4 months.

| Month | Net |
|---|---|
| 2026-07 | −3,514.18 |
| 2026-08 | +21,335.82 |
| 2026-09 | +1,023.45 |
| 2026-10 | −1,101.24 |

Opening 48,250.00 → closing 65,993.85, reconciled to the statement. The months
deliberately differ: a flat synthetic series would not exercise the variance
logic or look like anything a business would pay to automate.

## What the tests cover, and what they cannot

`openpyxl` writes formulas but does not calculate them, so a test cannot read a
computed total back out of the file. Rather than fake that, the tests prove the
parts that are provable:

- every summary cell is a `SUMIFS` formula, and its ranges point at the amount,
  category and month columns (9 assertions on formula shape and ranges),
- `Control` matches an independent recomputation done straight from the CSV,
- the control grid's grand total equals the CSV total,
- ledger row count equals input row count,
- reconciliation variance is a formula keyed to the statement cell, and moving
  the statement figure moves only that cell.

The value comparison itself happens in Excel, in `Checks`. That split is
deliberate and is the honest division of labour between Python and the file.

## Notes

- Categories are fixed at build time and become a dropdown on the ledger, so new
  rows cannot silently invent a category the summary will not pick up.
- Single account per workbook; the balance file carries one opening/closing pair.
- `data/make_sample_data.py` is seeded, so the committed CSVs regenerate byte for
  byte — the tests are deterministic.
- Requires `openpyxl`; tests need `pytest`.
