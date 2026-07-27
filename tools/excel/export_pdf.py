"""
Export ชีตรายเดือนจาก report-formatted.xlsx เป็น PDF

Output:
  04-export/pdf/01-employee/{year}/{Month}/Project_Report_Burapol_{Month}{year}_FTR_Timesheet.pdf
  04-export/manifest.json
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from pdf_export import (  # noqa: E402
    PDF_STATUS_PENDING,
    PDF_STATUS_SENT,
    export_months,
    list_exportable_sheets,
)
from paths import (  # noqa: E402
    DEFAULT_REPORT_YEAR,
    MONTH_SHEETS,
    WORKING_FORMATTED,
    ensure_dirs,
)
from report_footer import EMPLOYEE_FILE_TAG  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Export month sheet(s) from report-formatted.xlsx to PDF"
    )
    parser.add_argument("--month", help="Sheet name e.g. July (default: all month sheets)")
    parser.add_argument("--all", action="store_true", help="Export all month sheets")
    parser.add_argument("--year", type=int, default=DEFAULT_REPORT_YEAR)
    parser.add_argument(
        "--source",
        type=Path,
        default=WORKING_FORMATTED,
        help="Excel source (default: 02-working/report-formatted.xlsx)",
    )
    parser.add_argument(
        "--status",
        choices=(PDF_STATUS_PENDING, PDF_STATUS_SENT),
        default=PDF_STATUS_PENDING,
        help="employee = export เริ่มต้น, sent = accounting ส่งแล้ว (legacy alias)",
    )
    parser.add_argument(
        "--employee-tag",
        default=EMPLOYEE_FILE_TAG,
        help="ชื่อในไฟล์ PDF (default: Burapol)",
    )
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--list", action="store_true", help="List exportable sheets only")
    args = parser.parse_args()

    ensure_dirs()

    if args.list:
        for name in list_exportable_sheets(args.source):
            print(name)
        return

    if args.all:
        months = list(MONTH_SHEETS)
    elif args.month:
        months = [args.month]
    else:
        months = list(MONTH_SHEETS)

    export_months(
        months=months,
        source=args.source,
        year=args.year,
        status=args.status,
        employee_tag=args.employee_tag,
        dry_run=args.dry_run,
    )
    if not args.dry_run:
        print("Done.")


if __name__ == "__main__":
    main()
