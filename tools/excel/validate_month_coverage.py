"""Validate month work-day coverage in Excel from day 1 through a cutoff.

Detects gaps like 2026-09-04 (weekday, not holiday, no Excel work row).
Queue timesheet jobs MUST run this before today-only submit.

Exit codes:
  0 = complete (no seed gaps)
  2 = seed gaps found (agent must write Excel first, then intranet)
"""

from __future__ import annotations

import argparse
import calendar
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from excel_report import (  # noqa: E402
    NON_WORK_ATTENDANCE,
    read_month_entries,
    read_work_entries,
)
from paths import WORKING_DATA  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")


def _parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def _month_bounds(year: int, month_name: str) -> tuple[int, date, date]:
    month_num = datetime.strptime(month_name, "%B").month
    first = date(year, month_num, 1)
    last = date(year, month_num, calendar.monthrange(year, month_num)[1])
    return month_num, first, last


def _is_weekend(d: date) -> bool:
    return d.weekday() >= 5  # Sat=5 Sun=6


def expected_work_days(
    *,
    year: int,
    month_name: str,
    through: date,
    excel_path: Path,
) -> tuple[list[date], list[date], dict[date, str]]:
    """Return (expected_work, missing_excel, non_work_labels).

    expected_work = weekdays not marked holiday/non-work in Excel attendance,
    from month day-1 through ``through`` (inclusive, clamped to month).
    """
    month_num, first, last = _month_bounds(year, month_name)
    if through < first:
        raise ValueError(f"--through {through} is before {first}")
    cutoff = min(through, last)

    by_date = {
        e.event_date: e
        for e in read_month_entries(excel_path, month_name)
        if e.event_date.month == month_num
    }

    non_work_labels: dict[date, str] = {}
    expected: list[date] = []
    d = first
    while d <= cutoff:
        entry = by_date.get(d)
        attendance = (entry.attendance if entry else "") or ""
        if attendance in NON_WORK_ATTENDANCE:
            non_work_labels[d] = attendance
        elif _is_weekend(d):
            non_work_labels[d] = "weekend"
        else:
            # Weekday not marked holiday → treat as required work day
            expected.append(d)
        d += timedelta(days=1)

    work_present = {
        e.event_date
        for e in read_work_entries(
            excel_path,
            sheet_name=month_name,
            from_date=first,
            to_date=cutoff,
        )
    }
    # Also treat ATTENDANCE_WORK + non-empty detail as present even if helper
    # filters differ; is_work_day already requires both.
    missing = [day for day in expected if day not in work_present]
    return expected, missing, non_work_labels


def main() -> int:
    today = date.today()
    parser = argparse.ArgumentParser(
        description="Validate Excel work-day coverage from month day-1 through cutoff"
    )
    parser.add_argument("--month", default=today.strftime("%B"))
    parser.add_argument("--year", type=int, default=today.year)
    parser.add_argument(
        "--through",
        default=today.isoformat(),
        help="Inclusive cutoff (default: today). Queue jobs: use mission target date.",
    )
    parser.add_argument("--data", default=str(WORKING_DATA))
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print machine-readable JSON summary",
    )
    args = parser.parse_args()

    through = _parse_date(args.through)
    excel_path = Path(args.data)
    expected, missing, non_work = expected_work_days(
        year=args.year,
        month_name=args.month,
        through=through,
        excel_path=excel_path,
    )

    present = [d for d in expected if d not in set(missing)]
    summary = {
        "month": args.month,
        "year": args.year,
        "through": through.isoformat(),
        "excel": str(excel_path),
        "expected_work_days": [d.isoformat() for d in expected],
        "present_in_excel": [d.isoformat() for d in present],
        "missing_excel": [d.isoformat() for d in missing],
        "non_work_skipped": {
            d.isoformat(): label for d, label in sorted(non_work.items())
        },
        "verdict": "COMPLETE" if not missing else "SEED_GAPS",
        "agent_action": (
            "none"
            if not missing
            else (
                "For each missing date: (1) derive FTR evidence → seed "
                f"tools/excel/{args.month.lower()}_{args.year}_entries.py "
                "(2) write_month_data + format_report "
                "(3) submit_timesheet --from-date <month-01> --to-date <through> "
                "(Excel-first, then intranet; submit skips days already on web)"
            )
        ),
    }

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"=== Month coverage validate: {args.month} {args.year} through {through} ===")
        print(f"Excel: {excel_path}")
        print(f"Expected work days: {len(expected)}")
        for d in expected:
            mark = "OK" if d not in missing else "MISSING_EXCEL"
            print(f"  {d.isoformat()}  {mark}")
        if non_work:
            print("\nSkipped non-work / weekend:")
            for d, label in sorted(non_work.items()):
                print(f"  {d.isoformat()}  {label}")
        print(f"\nVerdict: {summary['verdict']}")
        if missing:
            print("Missing Excel work rows:")
            for d in missing:
                print(f"  - {d.isoformat()}")
            print("\nAgent action:")
            print(f"  {summary['agent_action']}")
        else:
            print("Excel seed complete for range.")
            print(
                "Next: dry-run then submit intranet "
                f"--from-date {date(args.year, datetime.strptime(args.month, '%B').month, 1).isoformat()} "
                f"--to-date {through.isoformat()} "
                "(fills any web gaps; skips days already present)"
            )

    return 0 if not missing else 2


if __name__ == "__main__":
    raise SystemExit(main())
