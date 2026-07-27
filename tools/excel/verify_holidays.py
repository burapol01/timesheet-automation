"""Verify month Excel holidays and work-day selection."""

from __future__ import annotations

import argparse
import calendar
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from excel_report import read_month_entries, read_work_entries  # noqa: E402
from paths import WORKING_DATA  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify holidays and work days in report-data.xlsx")
    parser.add_argument("--month", default="July")
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--data", default=str(WORKING_DATA))
    args = parser.parse_args()

    month_num = datetime.strptime(args.month, "%B").month
    from_d = date(args.year, month_num, 1)
    to_d = date(args.year, month_num, calendar.monthrange(args.year, month_num)[1])

    print(f"=== {args.month} non-work days in Excel ===")
    for e in read_month_entries(Path(args.data), args.month):
        if e.event_date.month != month_num:
            continue
        if not e.is_work_day and e.attendance:
            print(f"  {e.event_date}  {e.attendance}")

    print(f"\n=== {args.month} work days selected for submit ===")
    work = read_work_entries(
        Path(args.data),
        sheet_name=args.month,
        from_date=from_d,
        to_date=to_d,
    )
    for e in work:
        print(f"  {e.event_date}  {e.remark[:70]}...")
    print(f"\nTotal work days: {len(work)}")

    spill = [
        e for e in read_month_entries(Path(args.data), args.month) if e.event_date.month != month_num
    ]
    print(f"Spillover rows outside {args.month}: {len(spill)}")


if __name__ == "__main__":
    main()
