"""
Sync layout จาก report-layout-master.xlsx (ชีต July) เข้า template / ไฟล์ทำงาน

ห้าม sync เข้า report-layout-master.xlsx — ไฟล์นั้นเป็นแม่แบบ read-only
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
    apply_template_structure,
    clear_readonly,
    month_year_for_sheet,
    open_layout_reference,
    sync_january_master,
)
from paths import (  # noqa: E402
    ARCHIVE_DIR,
    DEFAULT_REPORT_YEAR,
    FORMAT_REFERENCE_SHEET,
    LAYOUT_MASTER_WORKBOOK,
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
    year: int = DEFAULT_REPORT_YEAR,
    archive: bool = True,
) -> None:
    if path.resolve() == LAYOUT_MASTER_WORKBOOK.resolve():
        print(f"  skip layout master (read-only): {path.name}")
        return
    if not path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์: {path}")
    if not LAYOUT_MASTER_WORKBOOK.exists():
        raise FileNotFoundError(f"ไม่พบ layout master: {LAYOUT_MASTER_WORKBOOK}")

    if archive:
        archive_file(path)

    clear_readonly(path)

    app = xw.App(visible=False, add_book=False)
    master_wb = None
    try:
        ref, master_wb = open_layout_reference(app)
        wb = app.books.open(str(path.resolve()))
        sheet_names = [s.name for s in wb.sheets]

        if "January" in sheet_names:
            sync_january_master(wb.sheets["January"])
            print("  restored January master (F43=Manager, H43=Approver, A44=title)")

        for sheet_name in MONTH_SHEETS:
            if sheet_name not in sheet_names:
                print(f"  skip {sheet_name} (no sheet)")
                continue
            if sheet_name == FORMAT_REFERENCE_SHEET:
                print(f"  skip {sheet_name} (layout master sheet)")
                continue
            sheet_year, sheet_month = month_year_for_sheet(sheet_name, year)
            apply_template_structure(
                ref,
                wb.sheets[sheet_name],
                year=sheet_year,
                month=sheet_month,
            )
            print(f"  synced {sheet_name}")

        wb.save()
        wb.close()
    finally:
        if master_wb is not None:
            master_wb.close()
        app.quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sync layout from report-layout-master.xlsx (July) to working/template files"
    )
    parser.add_argument(
        "--target",
        choices=("original", "working", "all"),
        default="all",
        help="original=project-report-template, working=data+formatted",
    )
    parser.add_argument("--year", type=int, default=DEFAULT_REPORT_YEAR)
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
        sync_workbook(path, year=args.year, archive=not args.no_archive)

    print("Done.")


if __name__ == "__main__":
    main()
