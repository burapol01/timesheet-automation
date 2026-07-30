"""
Sync layout ยืนยัน (summary + footer + logo/ลายเซ็น) เข้า template และไฟล์ทำงาน

ใช้หลังยืนยัน layout ใน report-formatted.xlsx แล้วต้องการอัปเดต:
- 01-original/project-report-template.xlsx
- 02-working/report-data.xlsx
- 02-working/report-formatted.xlsx
"""

from __future__ import annotations

import argparse
import shutil
import stat
import sys
from datetime import datetime
from pathlib import Path

import xlwings as xw

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))
sys.path.insert(0, str(ROOT / "tools" / "excel"))

from format_report import (  # noqa: E402
    apply_signature_footer,
    apply_template_structure,
    clear_readonly,
    month_year_for_sheet,
    sync_january_master,
)
from paths import (  # noqa: E402
    ARCHIVE_DIR,
    DEFAULT_REPORT_YEAR,
    FORMAT_REFERENCE_SHEET,
    MONTH_SHEETS,
    ORIGINAL_TEMPLATE,
    WORKING_DATA,
    WORKING_FORMATTED,
    ensure_dirs,
)


def archive_file(path: Path) -> Path:
    ensure_dirs()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = ARCHIVE_DIR / f"{path.stem}_{stamp}{path.suffix}"
    clear_readonly(path)
    shutil.copy2(path, dest)
    print(f"  archived -> {dest.name}")
    return dest


def sync_workbook(
    path: Path,
    *,
    reference_month: str = FORMAT_REFERENCE_SHEET,
    year: int = DEFAULT_REPORT_YEAR,
    archive: bool = True,
) -> None:
    if not path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์: {path}")

    if archive:
        archive_file(path)

    clear_readonly(path)

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(path.resolve()))
        sheet_names = [s.name for s in wb.sheets]
        if reference_month not in sheet_names:
            raise ValueError(f"ไม่พบชีตอ้างอิง: {reference_month}")

        ref = wb.sheets[reference_month]

        sync_january_master(wb.sheets["January"])
        print("  restored January master (F43=Manager, H43=Approver, A44=title)")

        # Bootstrap footer format บนชีตอ้างอิงก่อน
        ref_year, ref_month = month_year_for_sheet(reference_month, year)
        apply_signature_footer(ref, ref, year=ref_year, month=ref_month)
        print(f"  bootstrapped footer on {reference_month}")

        for sheet_name in MONTH_SHEETS:
            if sheet_name not in sheet_names:
                print(f"  skip {sheet_name} (no sheet)")
                continue
            tgt = wb.sheets[sheet_name]
            sheet_year, sheet_month = month_year_for_sheet(sheet_name, year)
            apply_template_structure(
                ref,
                tgt,
                year=sheet_year,
                month=sheet_month,
            )
            print(f"  synced {sheet_name}")

        wb.save()
        wb.close()
    finally:
        app.quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sync confirmed report layout to original template and working files"
    )
    parser.add_argument(
        "--target",
        choices=("original", "working", "all"),
        default="all",
        help="original=01-original, working=data+formatted, all=everything",
    )
    parser.add_argument("--year", type=int, default=DEFAULT_REPORT_YEAR)
    parser.add_argument(
        "--reference",
        default=FORMAT_REFERENCE_SHEET,
        help="Reference sheet for layout (default: June)",
    )
    parser.add_argument(
        "--no-archive",
        action="store_true",
        help="Skip backup to 03-archive/",
    )
    args = parser.parse_args()

    targets: list[Path] = []
    if args.target in ("original", "all"):
        targets.append(ORIGINAL_TEMPLATE)
    if args.target in ("working", "all"):
        targets.extend([WORKING_DATA, WORKING_FORMATTED])

    for path in targets:
        if not path.exists():
            print(f"skip missing: {path}")
            continue
        print(f"Syncing {path} ...")
        sync_workbook(
            path,
            reference_month=args.reference,
            year=args.year,
            archive=not args.no_archive,
        )

    print("Done.")


if __name__ == "__main__":
    main()
