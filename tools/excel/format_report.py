"""
จัด format ชีตรายเดือนให้ตรงต้นฉบับ

- คง highlight ของ template (ไม่ copy สีจาก June ทีละแถว)
- สูตรหัวข้อวันที่ D3 + สรุปท้ายตาราง
- ปรับความสูงแถวอัตโนมัติให้ข้อความไม่ทับ
- วางลายเซ็น/โลโก้ตำแหน่งเดียวกับชีตอ้างอิง
"""

from __future__ import annotations

import argparse
import calendar
import shutil
import stat
import sys
from datetime import date, datetime
from pathlib import Path

import xlwings as xw

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from excel_report import (  # noqa: E402
    COL_ATTENDANCE,
    COL_DETAIL,
    COL_JOB_CODE,
    COL_WORK_TYPE,
    HEADER_ROW,
    NON_WORK_ATTENDANCE,
    ReportEntry,
    read_month_entries,
    row_for_day,
)
from paths import (  # noqa: E402
    ARCHIVE_DIR,
    DEFAULT_REPORT_YEAR,
    FORMAT_REFERENCE_SHEET,
    LAYOUT_MASTER_WORKBOOK,
    LAYOUT_REFERENCE_DAYS,
    MONTH_SHEETS,
    WORKING_DATA,
    WORKING_FORMATTED,
    ensure_dirs,
)
from report_footer import (  # noqa: E402
    APPROVER_FOOTER_NAME,
    EMPLOYEE_TITLE,
    FOOTER_BLOCK_FIRST_ROW,
    FOOTER_BLOCK_LAST_ROW,
    FOOTER_DATE_COLUMNS,
    FOOTER_DATE_ROW,
    JANUARY_MASTER_APPROVER_NAME_CELL,
    JANUARY_MASTER_MANAGER_CELL,
    JANUARY_MASTER_TITLE_CELL,
    MANAGER_FOOTER_NAME,
    SIGNATURE_ANCHOR_ROW,
    SIGNATURE_HEIGHT,
    SIGNATURE_WIDTH,
)

FILL_YELLOW = 13434879
FILL_WHITE = 16777215
XL_TOP = -4160
XL_CENTER = -4108
HOLIDAY_SAMPLE_ROW = 14  # fallback — ใช้ holiday_sample_row() จาก July master ก่อน
WORK_ROW_SAMPLE = 9  # fallback — ใช้ work_row_sample() จาก July master ก่อน

HEADER_LABELS = (
    "วันที่",
    "การปฎิบัติงาน",
    "รหัสงาน",
    "ประเภทงาน",
)
HEADER_DETAIL_LABEL = "รายละเอียดของงานที่ทำ"

THAI_MONTHS = (
    '"มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",'
    '"กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"'
)

SIGNATURE_SHAPE = "Picture 1"
LOGO_SHAPE = "Picture 2"
APPROVER_SHAPE = "ApproverSig"
FOOTER_LABEL_ROW = 43
FOOTER_SIG_ROW = 44
FOOTER_NAME_ROW = 45
FOOTER_TITLE_ROW = 46
SIGNATURE_ROW_HEIGHT = 38.0
FOOTER_TEXT_ROW_HEIGHT = 18.0
MIN_WORK_ROW_HEIGHT = 23.3
ROW_HEIGHT_PAD = 2.0
AUTOFIT_TEMP_HEIGHT = 409.0


def summary_row_for_days(days_in_month: int) -> int:
    return HEADER_ROW + days_in_month + 2


def ref_summary_row() -> int:
    """แถวสรุปบนชีต July master (31 วัน)."""
    return summary_row_for_days(LAYOUT_REFERENCE_DAYS)


def work_row_sample(ref_sheet) -> int:
    """แถวงานตัวอย่างจาก July master — ใช้ copy format แถวข้อมูล."""
    last = HEADER_ROW + LAYOUT_REFERENCE_DAYS
    for row in range(HEADER_ROW + 1, last + 1):
        detail = ref_sheet.range((row, COL_DETAIL)).value
        job = ref_sheet.range((row, COL_JOB_CODE)).value
        if detail and job:
            return row
    return WORK_ROW_SAMPLE


def open_layout_reference(app: xw.App):
    """เปิด report-layout-master.xlsx (read-only) — ห้ามเขียนทับไฟล์นี้."""
    master_path = LAYOUT_MASTER_WORKBOOK.resolve()
    if not master_path.exists():
        raise FileNotFoundError(
            f"ไม่พบ layout master: {master_path}\n"
            "วาง report-layout-master.xlsx ใน 01-original/"
        )
    master_wb = app.books.open(str(master_path), read_only=True)
    return master_wb.sheets[FORMAT_REFERENCE_SHEET], master_wb


def paste_range_all(ref_sheet, tgt_sheet, ref_addr: str, tgt_addr: str | None = None) -> None:
    tgt_addr = tgt_addr or ref_addr
    ref_sheet.range(ref_addr).copy()
    tgt_sheet.range(tgt_addr).paste(paste="all")
    ref_sheet.book.api.Application.CutCopyMode = False


def align_data_area_to_master(ref_sheet, tgt_sheet) -> None:
    """คงความสูงแถว 8..41 จาก master ทุกแถว — ห้าม autofit ดัน footer."""
    for row in range(HEADER_ROW + 1, FOOTER_BLOCK_FIRST_ROW):
        h = ref_sheet.range((row, 1)).row_height
        if h:
            tgt_sheet.range((row, 1)).row_height = h


def copy_master_footer_block(
    ref_sheet,
    tgt_sheet,
    *,
    year: int,
    month: int,
    last_data_row: int,
    summary_row: int,
) -> None:
    """Copy summary+footer จาก July master ทั้ง block — เปลี่ยนแค่สูตรสรุปและวันที่."""
    ref_summary = ref_summary_row()

    copy_formats(
        ref_sheet.range(f"A{ref_summary}:G{ref_summary + 1}"),
        tgt_sheet.range(f"A{summary_row}:G{summary_row + 1}"),
    )
    for offset in range(2):
        for col in (2, 4):
            tgt_sheet.range((summary_row + offset, col)).value = ref_sheet.range(
                (ref_summary + offset, col)
            ).value
    apply_summary_formulas(tgt_sheet, last_data_row, summary_row)

    paste_range_all(
        ref_sheet,
        tgt_sheet,
        f"A{FOOTER_BLOCK_FIRST_ROW}:G{FOOTER_BLOCK_LAST_ROW}",
    )

    date_str = date(year, month, calendar.monthrange(year, month)[1]).strftime("%d/%m/%Y")
    for col in FOOTER_DATE_COLUMNS:
        tgt_sheet.range((FOOTER_DATE_ROW, col)).value = date_str

    for offset in range(2):
        ref_row = ref_summary + offset
        h = ref_sheet.range((ref_row, 1)).row_height
        if h:
            tgt_sheet.range((summary_row + offset, 1)).row_height = h
    for row in range(FOOTER_BLOCK_FIRST_ROW, FOOTER_BLOCK_LAST_ROW + 1):
        h = ref_sheet.range((row, 1)).row_height
        if h:
            tgt_sheet.range((row, 1)).row_height = h


def copy_column_widths(source, target, last_col: int = 9) -> None:
    for col in range(1, last_col + 1):
        w = source.range((1, col)).column_width
        if w:
            target.range((1, col)).column_width = w


def copy_row_heights(source, target, last_row: int = 47) -> None:
    for row in range(1, last_row + 1):
        h = source.range((row, 1)).row_height
        if h:
            target.range((row, 1)).row_height = h


def copy_formats(src: xw.Range, dst: xw.Range) -> None:
    src.copy()
    dst.paste(paste="formats")
    src.api.Application.CutCopyMode = False


def apply_table_header(ref_sheet, tgt_sheet) -> None:
    """แถวหัวตาราง — E7:G7 merge + กึ่งกลาง."""
    copy_formats(ref_sheet.range("A7:G7"), tgt_sheet.range("A7:G7"))
    for addr in ("A7:B7", "C7:D7", "E7:G7"):
        try:
            tgt_sheet.range(addr).unmerge()
        except Exception:
            pass

    for col, fallback in enumerate(HEADER_LABELS, start=1):
        val = ref_sheet.range((7, col)).value or fallback
        tgt_sheet.range((7, col)).value = val

    detail = tgt_sheet.range("E7:G7")
    detail.merge()
    detail.value = ref_sheet.range("E7").value or HEADER_DETAIL_LABEL
    detail.api.WrapText = True
    detail.api.HorizontalAlignment = XL_CENTER
    detail.api.VerticalAlignment = XL_CENTER


def date_header_formula(last_data_row: int) -> str:
    return (
        f'=DAY(A8)&" "&CHOOSE(MONTH(A8),{THAI_MONTHS})&" "&YEAR(A8)+543'
        f'&" ถึง "&DAY(A{last_data_row})&" "&CHOOSE(MONTH(A{last_data_row}),{THAI_MONTHS})'
        f'&" "&YEAR(A{last_data_row})+543'
    )


def summary_formulas(last_data_row: int, summary_row: int) -> dict[tuple[int, int], str]:
    b_range = f"$B$8:$B${last_data_row}"
    return {
        (summary_row, 3): f'=COUNTIF({b_range}, "เข้าปฎิบัติงาน")',
        (summary_row, 5): (
            f'=SUM(COUNTIF({b_range}, {{"วันเสาร์","วันอาทิตย์","วันหยุดนักขัตฤกษ์"}}))'
        ),
        (summary_row + 1, 3): f'=COUNTIF({b_range}, "ไม่เข้าปฎิบัติงาน")',
        (summary_row + 1, 5): f'=COUNTIF({b_range}, "อื่นๆ โปรดระบุรายละเอียด ")',
    }


def month_year_for_sheet(sheet_name: str, year: int = DEFAULT_REPORT_YEAR) -> tuple[int, int]:
    dt = datetime.strptime(sheet_name, "%B").replace(year=year)
    return dt.year, dt.month


def month_year_from_entries(entries: list[ReportEntry], sheet_name: str) -> tuple[int, int]:
    in_month = [e for e in entries if e.event_date.strftime("%B") == sheet_name]
    if in_month:
        return in_month[0].event_date.year, in_month[0].event_date.month
    dt = datetime.strptime(sheet_name, "%B").replace(year=date.today().year)
    return dt.year, dt.month


def is_non_work_attendance(attendance: str) -> bool:
    if attendance in NON_WORK_ATTENDANCE:
        return True
    return any(k in attendance for k in ("เสาร์", "อาทิตย์", "หยุด"))


def get_shape(sheet, shape_name: str):
    for i in range(1, sheet.api.Shapes.Count + 1):
        sh = sheet.api.Shapes(i)
        if sh.Name == shape_name:
            return sh
    return None


def delete_shape(sheet, shape_name: str) -> None:
    sh = get_shape(sheet, shape_name)
    if sh is not None:
        sh.Delete()


def apply_all_entries(sheet, entries: list[ReportEntry], sheet_name: str) -> int:
    work_count = 0
    for entry in entries:
        if entry.event_date.strftime("%B") != sheet_name:
            continue
        row = row_for_day(entry.event_date.day)
        if entry.is_work_day:
            sheet.range((row, COL_ATTENDANCE)).value = entry.attendance
            sheet.range((row, COL_JOB_CODE)).value = entry.job_code
            sheet.range((row, COL_WORK_TYPE)).value = entry.work_type
            sheet.range((row, COL_DETAIL)).value = entry.detail
            work_count += 1
        elif entry.attendance:
            sheet.range((row, COL_ATTENDANCE)).value = entry.attendance
    return work_count


def apply_row_highlights(sheet, entries: list[ReportEntry], sheet_name: str) -> None:
    for entry in entries:
        if entry.event_date.strftime("%B") != sheet_name:
            continue
        row = row_for_day(entry.event_date.day)
        row_rng = sheet.range((row, 1), (row, 7))
        if is_non_work_attendance(entry.attendance):
            row_rng.api.Interior.Color = FILL_YELLOW
            sheet.range((row, COL_JOB_CODE)).value = None
            sheet.range((row, COL_WORK_TYPE)).value = None
            sheet.range((row, COL_DETAIL)).value = None
        else:
            row_rng.api.Interior.Color = FILL_WHITE


def copy_work_row_formats(ref_sheet, tgt_sheet, work_rows: list[int]) -> None:
    """Copy wrap/align/border from a sample work row on July master."""
    sample_row = work_row_sample(ref_sheet)
    for row in work_rows:
        copy_formats(
            ref_sheet.range(f"A{sample_row}:G{sample_row}"),
            tgt_sheet.range(f"A{row}:G{row}"),
        )


def apply_text_layout(sheet, work_rows: list[int]) -> None:
    for row in work_rows:
        for col in (COL_WORK_TYPE, COL_DETAIL):
            cell = sheet.range((row, col))
            cell.api.WrapText = True
            cell.api.VerticalAlignment = XL_TOP
        detail_block = sheet.range((row, COL_DETAIL), (row, 7))
        try:
            detail_block.api.WrapText = True
            detail_block.api.VerticalAlignment = XL_TOP
        except Exception:
            pass


def _autofit_row_with_wrap(sheet, row: int) -> float:
    """AutoFit แถว wrap — วัด D กับ E แยก (E ใช้ความกว้าง E:G รวม ไม่ใช่แค่ col E)."""
    merge_addr = f"E{row}:G{row}"
    e_col_width = sheet.range((1, COL_DETAIL)).column_width
    merged_detail_width = sum(
        sheet.range((1, col)).column_width for col in (COL_DETAIL, 6, 7)
    )

    try:
        sheet.range(merge_addr).unmerge()
    except Exception:
        pass

    d_text = sheet.range((row, COL_WORK_TYPE)).value
    e_text = sheet.range((row, COL_DETAIL)).value

    def _measure(d_val, e_val, *, widen_detail: bool) -> float:
        sheet.range((1, COL_DETAIL)).column_width = (
            merged_detail_width if widen_detail else e_col_width
        )
        sheet.range((row, COL_WORK_TYPE)).value = d_val if d_val is not None else ""
        sheet.range((row, COL_DETAIL)).value = e_val if e_val is not None else ""
        for col in (COL_WORK_TYPE, COL_DETAIL):
            cell = sheet.range((row, col))
            cell.api.WrapText = True
            cell.api.VerticalAlignment = XL_TOP
        row_api = sheet.range((row, 1)).api.EntireRow
        row_api.RowHeight = AUTOFIT_TEMP_HEIGHT
        row_api.AutoFit()
        return float(sheet.range((row, 1)).row_height or MIN_WORK_ROW_HEIGHT)

    h_d = _measure(d_text, "", widen_detail=False)
    h_e = _measure("", e_text, widen_detail=True)
    height = max(MIN_WORK_ROW_HEIGHT, max(h_d, h_e) + ROW_HEIGHT_PAD)

    sheet.range((1, COL_DETAIL)).column_width = e_col_width
    sheet.range((row, COL_WORK_TYPE)).value = d_text
    sheet.range((row, COL_DETAIL)).value = e_text
    for col in (COL_WORK_TYPE, COL_DETAIL):
        cell = sheet.range((row, col))
        cell.api.WrapText = True
        cell.api.VerticalAlignment = XL_TOP
    try:
        sheet.range(merge_addr).merge()
        detail_block = sheet.range((row, COL_DETAIL), (row, 7))
        detail_block.api.WrapText = True
        detail_block.api.VerticalAlignment = XL_TOP
    except Exception:
        pass

    return height


def fit_work_row_heights(sheet, work_rows: list[int], ref_sheet=None) -> None:
    """ปรับความสูงทุกแถวงานตามเนื้อหาจริง (D + E merge) — แถวละค่า."""
    del ref_sheet  # kept for call-site compatibility
    for row in work_rows:
        height = _autofit_row_with_wrap(sheet, row)
        sheet.range((row, 1)).row_height = height


def autofit_work_rows(sheet, work_rows: list[int], ref_sheet=None) -> None:
    fit_work_row_heights(sheet, work_rows, ref_sheet)


def autofit_data_area(sheet, first_row: int, last_row: int) -> None:
    for row in range(first_row, last_row + 1):
        sheet.range((row, 1)).api.EntireRow.AutoFit()
        if sheet.range((row, 1)).row_height < MIN_WORK_ROW_HEIGHT:
            sheet.range((row, 1)).row_height = MIN_WORK_ROW_HEIGHT


def autofit_footer_rows(sheet, rows: list[int]) -> None:
    for row in rows:
        sheet.range((row, 1)).api.EntireRow.AutoFit()


def work_row_numbers(entries: list[ReportEntry], sheet_name: str) -> list[int]:
    rows: list[int] = []
    for entry in entries:
        if entry.event_date.strftime("%B") != sheet_name:
            continue
        if entry.is_work_day:
            rows.append(row_for_day(entry.event_date.day))
    return rows


def holiday_row_numbers(entries: list[ReportEntry], sheet_name: str) -> list[int]:
    rows: list[int] = []
    for entry in entries:
        if entry.event_date.strftime("%B") != sheet_name:
            continue
        if entry.attendance and is_non_work_attendance(entry.attendance):
            rows.append(row_for_day(entry.event_date.day))
    return rows


def scan_holiday_rows(sheet, first_row: int, last_row: int) -> list[int]:
    rows: list[int] = []
    for row in range(first_row, last_row + 1):
        val = sheet.range((row, COL_ATTENDANCE)).value
        if val and is_non_work_attendance(str(val).strip()):
            rows.append(row)
    return rows


def holiday_sample_row(ref_sheet) -> int:
    last = HEADER_ROW + 31
    for row in range(HEADER_ROW + 1, last + 1):
        val = ref_sheet.range((row, COL_ATTENDANCE)).value
        if val and is_non_work_attendance(str(val).strip()):
            return row
    return HOLIDAY_SAMPLE_ROW


def fit_holiday_rows(ref_sheet, tgt_sheet, rows: list[int]) -> None:
    """แถบวันหยุด/เสาร์-อาทิตย์ — สีเหลืองเต็มแถว ความสูงเท่า template."""
    if not rows:
        return
    sample = holiday_sample_row(ref_sheet)
    for row in rows:
        copy_formats(
            ref_sheet.range(f"A{sample}:G{sample}"),
            tgt_sheet.range(f"A{row}:G{row}"),
        )
        tgt_sheet.range((row, 1), (row, 7)).api.Interior.Color = FILL_YELLOW
        for col in range(COL_JOB_CODE, 8):
            cell = tgt_sheet.range((row, col))
            cell.value = None
            cell.api.WrapText = False
        att = tgt_sheet.range((row, COL_ATTENDANCE))
        att.api.WrapText = False
        att.api.VerticalAlignment = XL_CENTER
        att.api.HorizontalAlignment = XL_CENTER
        master_h = ref_sheet.range((row, 1)).row_height
        tgt_sheet.range((row, 1)).row_height = master_h or MIN_WORK_ROW_HEIGHT


def _paste_copied_shape(ref_sheet, tgt_sheet) -> None:
    ref_sheet.activate()
    tgt_sheet.activate()
    tgt_sheet.api.Paste()
    ref_sheet.book.api.Application.CutCopyMode = False


def place_employee_signature(ref_sheet, tgt_sheet) -> bool:
    """วางลายเซ็นพนักงาน — offset จาก July master shape (anchor แถว 42)."""
    return place_shape_like_ref(
        ref_sheet,
        tgt_sheet,
        SIGNATURE_SHAPE,
        anchor_row=SIGNATURE_ANCHOR_ROW,
    )


def place_shape_like_ref(
    ref_sheet,
    tgt_sheet,
    shape_name: str,
    *,
    anchor_row: int = SIGNATURE_ANCHOR_ROW,
) -> bool:
    ref_sh = get_shape(ref_sheet, shape_name)
    if ref_sh is None:
        return False

    delete_shape(tgt_sheet, shape_name)

    ref_sheet.activate()
    ref_sh.Copy()
    _paste_copied_shape(ref_sheet, tgt_sheet)

    new_sh = tgt_sheet.api.Shapes(tgt_sheet.api.Shapes.Count)
    ref_anchor_top = ref_sheet.range((anchor_row, 1)).api.Top
    tgt_anchor_top = tgt_sheet.range((anchor_row, 1)).api.Top
    top_offset = ref_sh.Top - ref_anchor_top

    new_sh.Top = tgt_anchor_top + top_offset
    new_sh.Left = ref_sh.Left
    try:
        new_sh.LockAspectRatio = 0
    except Exception:
        pass
    new_sh.Width = ref_sh.Width
    new_sh.Height = ref_sh.Height
    new_sh.Placement = 1  # xlMoveAndSize
    try:
        new_sh.Name = shape_name
    except Exception:
        pass
    return True


def sync_signature_size_from_master(ref_sheet, tgt_sheet) -> bool:
    """ปรับเฉพาะ Width/Height ของ Picture 1 ให้ตรง July master — ไม่ขยับตำแหน่ง."""
    ref_sh = get_shape(ref_sheet, SIGNATURE_SHAPE)
    tgt_sh = get_shape(tgt_sheet, SIGNATURE_SHAPE)
    if ref_sh is None or tgt_sh is None:
        return False
    try:
        tgt_sh.LockAspectRatio = 0
    except Exception:
        pass
    tgt_sh.Width = ref_sh.Width
    tgt_sh.Height = ref_sh.Height
    return True


def replace_logo_from_ref(ref_sheet, tgt_sheet) -> bool:
    ref_sh = get_shape(ref_sheet, LOGO_SHAPE)
    if ref_sh is None:
        return False
    delete_shape(tgt_sheet, LOGO_SHAPE)
    ref_sheet.activate()
    ref_sh.Copy()
    _paste_copied_shape(ref_sheet, tgt_sheet)
    new_sh = tgt_sheet.api.Shapes(tgt_sheet.api.Shapes.Count)
    try:
        new_sh.Name = LOGO_SHAPE
    except Exception:
        pass
    new_sh.Top = ref_sh.Top
    new_sh.Left = ref_sh.Left
    try:
        new_sh.LockAspectRatio = 0
    except Exception:
        pass
    new_sh.Width = ref_sh.Width
    new_sh.Height = ref_sh.Height
    return True


def sync_logo_from_ref(ref_sheet, tgt_sheet) -> bool:
    ref_sh = get_shape(ref_sheet, LOGO_SHAPE)
    tgt_sh = get_shape(tgt_sheet, LOGO_SHAPE)
    if ref_sh is None or tgt_sh is None:
        return replace_logo_from_ref(ref_sheet, tgt_sheet)
    tgt_sh.Top = ref_sh.Top
    tgt_sh.Left = ref_sh.Left
    try:
        tgt_sh.LockAspectRatio = 0
    except Exception:
        pass
    tgt_sh.Width = ref_sh.Width
    tgt_sh.Height = ref_sh.Height
    try:
        tgt_sh.Placement = 0
    except Exception:
        pass
    return True


def ensure_shapes(ref_sheet, tgt_sheet, *, quiet: bool = False) -> None:
    delete_shape(tgt_sheet, APPROVER_SHAPE)
    if ref_sheet.name == tgt_sheet.name:
        return
    for extra in (LOGO_SHAPE, SIGNATURE_SHAPE, "Picture 3", "Picture 4"):
        delete_shape(tgt_sheet, extra)
    if replace_logo_from_ref(ref_sheet, tgt_sheet):
        if not quiet:
            print(f"  copied logo from {FORMAT_REFERENCE_SHEET}")
    else:
        sync_logo_from_ref(ref_sheet, tgt_sheet)
        if not quiet:
            print("  synced logo size")

    if place_employee_signature(ref_sheet, tgt_sheet):
        sync_signature_size_from_master(ref_sheet, tgt_sheet)
        if not quiet:
            print("  placed employee signature")
    elif not quiet:
        print("  WARNING: employee signature not found on reference sheet")


def copy_page_setup(ref_sheet, tgt_sheet) -> None:
    """Copy page setup จาก July master (margins, fit-to-page, ฯลฯ)."""
    r = ref_sheet.api.PageSetup
    t = tgt_sheet.api.PageSetup
    t.PrintArea = r.PrintArea
    t.FitToPagesWide = r.FitToPagesWide
    t.FitToPagesTall = r.FitToPagesTall
    t.Zoom = r.Zoom
    t.Orientation = r.Orientation
    t.CenterHorizontally = r.CenterHorizontally
    t.CenterVertically = r.CenterVertically
    t.LeftMargin = r.LeftMargin
    t.RightMargin = r.RightMargin
    t.TopMargin = r.TopMargin
    t.BottomMargin = r.BottomMargin
    t.PaperSize = r.PaperSize


def fix_date_header(ref_sheet, tgt_sheet, last_data_row: int) -> None:
    """D3 = สูตรช่วงวันที่เดียว; ล้าง E3/F3 ที่ template July ยังค้าง."""
    for addr in ("E3", "F3", "G3"):
        cell = tgt_sheet.range(addr)
        cell.clear_contents()
        cell.api.Interior.ColorIndex = -4142
    tgt_sheet.range("D3").formula = date_header_formula(last_data_row)
    copy_formats(ref_sheet.range("C3:G3"), tgt_sheet.range("C3:G3"))
    for addr in ("E3", "F3", "G3"):
        tgt_sheet.range(addr).clear_contents()


def _copy_cell_from_ref(
    ref_sheet, tgt_sheet, ref_addr: str, tgt_addr: str | None = None
) -> None:
    tgt_addr = tgt_addr or ref_addr
    ref = ref_sheet.range(ref_addr)
    tgt = tgt_sheet.range(tgt_addr)
    if ref.formula and str(ref.formula).startswith("="):
        tgt.formula = ref.formula
    else:
        tgt.value = ref.value


def sync_january_master(january_sheet) -> None:
    """January เก็บ master — F43=Manager, H43=Approver, A44=Programmer."""
    january_sheet.range(JANUARY_MASTER_MANAGER_CELL).value = MANAGER_FOOTER_NAME
    january_sheet.range(JANUARY_MASTER_APPROVER_NAME_CELL).value = APPROVER_FOOTER_NAME
    january_sheet.range(JANUARY_MASTER_TITLE_CELL).value = EMPLOYEE_TITLE


def apply_summary_formulas(tgt, last_data_row: int, summary_row: int) -> None:
    for (row, col), formula in summary_formulas(last_data_row, summary_row).items():
        tgt.range((row, col)).formula = formula


def apply_template_structure(
    ref_sheet,
    tgt_sheet,
    *,
    year: int,
    month: int,
) -> None:
    """จัด header, summary, footer จาก layout master — ไม่แตะข้อมูลแถวงาน."""
    if tgt_sheet.name == "January":
        sync_january_master(tgt_sheet)
        return
    if tgt_sheet.name == FORMAT_REFERENCE_SHEET:
        return

    days_in_month = calendar.monthrange(year, month)[1]
    last_data_row = HEADER_ROW + days_in_month
    summary_row = summary_row_for_days(days_in_month)

    copy_column_widths(ref_sheet, tgt_sheet)
    copy_row_heights(ref_sheet, tgt_sheet, last_row=FOOTER_BLOCK_LAST_ROW)
    copy_page_setup(ref_sheet, tgt_sheet)
    apply_table_header(ref_sheet, tgt_sheet)
    fix_date_header(ref_sheet, tgt_sheet, last_data_row)
    copy_master_footer_block(
        ref_sheet,
        tgt_sheet,
        year=year,
        month=month,
        last_data_row=last_data_row,
        summary_row=summary_row,
    )
    ensure_shapes(ref_sheet, tgt_sheet, quiet=True)

    holiday_rows = scan_holiday_rows(tgt_sheet, HEADER_ROW + 1, last_data_row)
    fit_holiday_rows(ref_sheet, tgt_sheet, holiday_rows)


def clear_readonly(path: Path) -> None:
    if path.exists():
        path.chmod(stat.S_IWRITE | stat.S_IREAD)


def format_month_sheet(
    *,
    target_month: str,
    reference_month: str | None = None,
) -> Path:
    ensure_dirs()
    reference_month = reference_month or FORMAT_REFERENCE_SHEET

    if not LAYOUT_MASTER_WORKBOOK.exists():
        raise FileNotFoundError(
            f"ไม่พบ layout master: {LAYOUT_MASTER_WORKBOOK}\n"
            "ต้องมี report-formatted.xlsx ที่ยืนยัน July แล้ว (กู้จาก 03-archive/ ได้)"
        )
    if not WORKING_DATA.exists():
        raise FileNotFoundError(
            f"ไม่พบไฟล์ข้อมูล: {WORKING_DATA}\n"
            "รัน tools/excel/write_month_data.py ก่อน"
        )

    entries = read_month_entries(WORKING_DATA, target_month)
    year, month = month_year_from_entries(entries, target_month)
    days_in_month = calendar.monthrange(year, month)[1]
    last_data_row = HEADER_ROW + days_in_month
    summary_row = summary_row_for_days(days_in_month)
    work_rows = work_row_numbers(entries, target_month)
    holiday_rows = holiday_row_numbers(entries, target_month)

    if WORKING_FORMATTED.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        archive = ARCHIVE_DIR / f"report-formatted_{stamp}.xlsx"
        clear_readonly(WORKING_FORMATTED)
        shutil.copy2(WORKING_FORMATTED, archive)
        print(f"Archived -> {archive.name}")

    clear_readonly(WORKING_FORMATTED)

    if target_month == FORMAT_REFERENCE_SHEET:
        raise ValueError(
            f"ห้าม format ชีต {FORMAT_REFERENCE_SHEET} ด้วย script — "
            f"แก้ใน {LAYOUT_MASTER_WORKBOOK.name} แล้ว copy ใหม่"
        )

    app = xw.App(visible=False, add_book=False)
    master_wb = None
    try:
        ref, master_wb = open_layout_reference(app)
        wb = app.books.open(str(WORKING_FORMATTED.resolve()))
        if target_month not in [s.name for s in wb.sheets]:
            raise ValueError(f"ไม่พบชีต {target_month!r} ใน {WORKING_FORMATTED.name}")
        tgt = wb.sheets[target_month]

        copy_column_widths(ref, tgt)
        copy_row_heights(ref, tgt, last_row=FOOTER_BLOCK_LAST_ROW)
        copy_page_setup(ref, tgt)
        apply_table_header(ref, tgt)

        fix_date_header(ref, tgt, last_data_row)

        n = apply_all_entries(tgt, entries, target_month)
        copy_work_row_formats(ref, tgt, work_rows)
        apply_row_highlights(tgt, entries, target_month)
        apply_text_layout(tgt, work_rows)
        fit_holiday_rows(ref, tgt, holiday_rows)
        align_data_area_to_master(ref, tgt)
        print(f"Applied {n} work rows; holiday rows {len(holiday_rows)}")

        copy_master_footer_block(
            ref,
            tgt,
            year=year,
            month=month,
            last_data_row=last_data_row,
            summary_row=summary_row,
        )
        ensure_shapes(ref, tgt)

        wb.save()
        wb.close()
    finally:
        if master_wb is not None:
            master_wb.close()
        app.quit()

    return WORKING_FORMATTED


def fix_signature_sizes(
    months: list[str] | None = None,
    *,
    workbook: Path | None = None,
) -> Path:
    """Sync Picture 1 W/H จาก layout master — ไม่แตะ layout อื่น."""
    months = months or ["September", "October", "November", "December"]
    path = Path(workbook or WORKING_FORMATTED).resolve()
    master_path = LAYOUT_MASTER_WORKBOOK.resolve()
    ensure_dirs()

    app = xw.App(visible=False, add_book=False)
    master_wb = None
    try:
        if path == master_path:
            wb = app.books.open(str(path))
            ref = wb.sheets[FORMAT_REFERENCE_SHEET]
        else:
            ref, master_wb = open_layout_reference(app)
            wb = app.books.open(str(path))
        fixed = 0
        for month in months:
            if month not in [s.name for s in wb.sheets]:
                print(f"  skip {month} (no sheet)")
                continue
            if month == FORMAT_REFERENCE_SHEET:
                continue
            if sync_signature_size_from_master(ref, wb.sheets[month]):
                fixed += 1
                print(f"  fixed signature size: {month}")
        wb.save()
        wb.close()
        print(f"Done — {fixed} sheet(s)")
    finally:
        if master_wb is not None:
            try:
                master_wb.close()
            except Exception:
                pass
        app.quit()
    return path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Format month sheet using July layout master in report-formatted.xlsx"
    )
    parser.add_argument("--month", help="Sheet name e.g. July, August")
    parser.add_argument(
        "--fix-signature-size",
        action="store_true",
        help="Sync Picture 1 W/H only (default months: Sep-Dec)",
    )
    parser.add_argument(
        "--months",
        nargs="*",
        default=["September", "October", "November", "December"],
        help="Sheets for --fix-signature-size",
    )
    parser.add_argument(
        "--reference",
        default=FORMAT_REFERENCE_SHEET,
        help=f"Layout reference sheet (default: {FORMAT_REFERENCE_SHEET})",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=None,
        help="Target xlsx (default: report-formatted; use layout master path to fix master)",
    )
    args = parser.parse_args()
    if args.fix_signature_size:
        fix_signature_sizes(args.months, workbook=args.workbook)
        return
    if not args.month:
        args.month = "July"
    out = format_month_sheet(target_month=args.month, reference_month=args.reference)
    print(f"Done: {out}")


if __name__ == "__main__":
    main()
