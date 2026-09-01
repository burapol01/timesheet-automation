"""
เขียนข้อมูลรายเดือนลง 02-working/report-data.xlsx เท่านั้น

ห้ามแตะ 01-original/ — ใช้ openpyxl ได้ (ไม่มีโลโก้ในไฟล์ data ก็ไม่เป็นไร)
หลังเขียนเสร็จ รัน tools/excel/format_report.py เพื่อสร้าง PDF-ready file
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EXCEL_TOOLS = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "lib"))
sys.path.insert(0, str(EXCEL_TOOLS))

from excel_report import init_working_data, write_entries  # noqa: E402
from august_2026_entries import AUGUST_2026_ENTRIES  # noqa: E402
from july_2026_entries import JULY_2026_ENTRIES  # noqa: E402
from september_2026_entries import SEPTEMBER_2026_ENTRIES  # noqa: E402
from paths import WORKING_DATA  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser(description="Write month data to report-data.xlsx")
    parser.add_argument("--month", default="July", choices=["July", "August", "September"])
    args = parser.parse_args()

    init_working_data()
    if args.month == "July":
        entries = JULY_2026_ENTRIES
    elif args.month == "August":
        entries = AUGUST_2026_ENTRIES
    elif args.month == "September":
        entries = SEPTEMBER_2026_ENTRIES
    else:
        raise ValueError(f"No seed data for {args.month}")

    n = write_entries(WORKING_DATA, args.month, entries)
    print(f"Wrote {n} rows to {WORKING_DATA}")
    print("Next: python tools/excel/format_report.py --month", args.month)
    print("(format_report ตั้ง WrapText + ความสูงแถวให้ remark หลายบรรทัดไม่ถูกหุบ)")


if __name__ == "__main__":
    main()
