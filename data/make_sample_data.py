"""Regenerate the sample ledger and balances used by the demo and the tests.

Deterministic: same seed, same files. Run it only if you want to rebuild the
committed CSVs.

    python data/make_sample_data.py
"""

from __future__ import annotations

import csv
import random
from pathlib import Path

SEED = 20261117
HERE = Path(__file__).resolve().parent
OPENING = 48_250.00

REVENUE = ["Consulting Revenue", "Product Revenue", "Retainer Revenue"]
EXPENSE = [
    "Payroll",
    "Rent",
    "Software",
    "Marketing",
    "Travel",
    "Utilities",
    "Professional Fees",
]

MEMOS = {
    "Consulting Revenue": ["project invoice", "advisory retainer", "implementation",
                            "scoping call", "data audit"],
    "Product Revenue": ["subscription renewals", "seat expansion", "annual plan",
                         "usage overage"],
    "Retainer Revenue": ["monthly retainer", "support retainer"],
    "Payroll": ["semi-monthly payroll", "payroll taxes", "contractor payout"],
    "Rent": ["office lease"],
    "Software": ["cloud hosting", "analytics seats", "source control", "ci minutes"],
    "Marketing": ["paid search", "conference booth", "content retainer"],
    "Travel": ["client visit", "rail fare", "hotel", "meals"],
    "Utilities": ["electricity", "internet"],
    "Professional Fees": ["bookkeeping", "legal review", "tax prep"],
}

MONTHS = ["2026-07", "2026-08", "2026-09", "2026-10"]


def build_rows() -> list[tuple[str, str, str, float, str]]:
    rng = random.Random(SEED)
    rows: list[tuple[str, str, str, float, str]] = []
    for month in MONTHS:
        year, mon = month.split("-")
        for category in REVENUE:
            count = rng.randint(1, 3)
            for _ in range(count):
                day = rng.randint(1, 27)
                amount = round(rng.uniform(2_500, 14_000), 2)
                rows.append(
                    (
                        f"{year}-{mon}-{day:02d}",
                        category,
                        "Operating",
                        amount,
                        rng.choice(MEMOS[category]),
                    )
                )
        for category in EXPENSE:
            if category in ("Rent", "Utilities"):
                count = 1
            elif category == "Payroll":
                count = 2
            elif category == "Travel":
                count = rng.randint(0, 2)
            else:
                count = rng.randint(1, 2)
            for _ in range(count):
                day = rng.randint(1, 27)
                if category == "Rent":
                    amount = -6_500.00
                elif category == "Payroll":
                    amount = -round(rng.uniform(9_500, 12_500), 2)
                elif category == "Software":
                    amount = -round(rng.uniform(180, 2_400), 2)
                elif category == "Travel":
                    amount = -round(rng.uniform(150, 2_000), 2)
                elif category == "Marketing":
                    amount = -round(rng.uniform(500, 4_000), 2)
                elif category == "Utilities":
                    amount = -round(rng.uniform(210, 640), 2)
                else:
                    amount = -round(rng.uniform(400, 2_500), 2)
                rows.append(
                    (
                        f"{year}-{mon}-{day:02d}",
                        category,
                        "Operating",
                        amount,
                        rng.choice(MEMOS[category]),
                    )
                )
    rows.sort(key=lambda row: (row[0], row[1]))
    return rows


def main() -> int:
    rows = build_rows()
    txn_path = HERE / "transactions.csv"
    with txn_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["date", "category", "account", "amount", "memo"])
        writer.writerows(rows)

    closing = round(OPENING + sum(row[3] for row in rows), 2)
    bal_path = HERE / "balances.csv"
    with bal_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["account", "period", "opening", "closing"])
        writer.writerow(
            ["Operating", f"{MONTHS[0]} to {MONTHS[-1]}", f"{OPENING:.2f}", f"{closing:.2f}"]
        )

    print(f"transactions.csv: {len(rows)} rows")
    print(f"balances.csv: opening {OPENING:,.2f} closing {closing:,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
