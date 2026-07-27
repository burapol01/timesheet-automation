"""Export formatted Excel sheets to PDF for delivery / email."""

from __future__ import annotations

import json
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

import xlwings as xw

from paths import (
    DRAFT_STEP_ATTACH,
    DRAFT_STEP_SNAPSHOT,
    EXPORT_MANIFEST,
    INBOUND_ROLE_STAGE,
    MONTH_SHEETS,
    WORKING_FORMATTED,
    ensure_dirs,
    pdf_stage_dir,
)
from report_footer import EMPLOYEE_FILE_TAG

# stage keys (ใช้กับ pdf_output_path / export)
PDF_STAGE_EMPLOYEE = "employee"
PDF_STAGE_MANAGER_SIGNED = "manager_signed"
PDF_STAGE_APPROVER_SIGNED = "approver_signed"
PDF_STAGE_ACCOUNTING_SENT = "accounting_sent"

# Legacy status names
PDF_STATUS_PENDING = PDF_STAGE_EMPLOYEE
PDF_STATUS_SENT = PDF_STAGE_ACCOUNTING_SENT

_STAGE_SUFFIX = {
    PDF_STAGE_EMPLOYEE: "",
    "manager_draft": "_to_manager",
    PDF_STAGE_MANAGER_SIGNED: "_manager_signed",
    "approver_draft": "_to_approver",
    PDF_STAGE_APPROVER_SIGNED: "_approver_signed",
    "accounting_draft": "_to_accounting",
    PDF_STAGE_ACCOUNTING_SENT: "_accounting_sent",
}


@dataclass
class PdfExportRecord:
    month: str
    year: int
    sheet: str
    source: str
    pdf: str
    stage: str
    exported_at: str


def pdf_basename(*, month: str, year: int, stage: str = PDF_STAGE_EMPLOYEE) -> str:
    suffix = _STAGE_SUFFIX.get(stage)
    if suffix is None:
        raise ValueError(f"Unknown PDF stage: {stage}")
    return (
        f"Project_Report_{EMPLOYEE_FILE_TAG}_{month}{year}_FTR_Timesheet{suffix}.pdf"
    )


def pdf_filename(
    *,
    month: str,
    year: int,
    stage: str = PDF_STAGE_EMPLOYEE,
    employee_tag: str = EMPLOYEE_FILE_TAG,
) -> str:
    if employee_tag != EMPLOYEE_FILE_TAG:
        suffix = _STAGE_SUFFIX.get(stage, "")
        return f"Project_Report_{employee_tag}_{month}{year}_FTR_Timesheet{suffix}.pdf"
    return pdf_basename(month=month, year=year, stage=stage)


def pdf_output_path(
    *,
    month: str,
    year: int,
    stage: str = PDF_STAGE_EMPLOYEE,
    status: str | None = None,
    employee_tag: str = EMPLOYEE_FILE_TAG,
) -> Path:
    """Path สำหรับ PDF ตาม stage ในลำดับชั้น workflow."""
    if status is not None:
        legacy = {"pending": PDF_STAGE_EMPLOYEE, "sent": PDF_STAGE_ACCOUNTING_SENT}
        stage = legacy.get(status, status)
    month_dir = pdf_stage_dir(stage=stage, year=year) / month
    month_dir.mkdir(parents=True, exist_ok=True)
    return month_dir / pdf_filename(
        month=month, year=year, stage=stage, employee_tag=employee_tag
    )


def pdf_for_draft_step(*, step: str, month: str, year: int) -> Path:
    """PDF ที่แนบเมื่อสร้าง draft ของ step (manager|approver|accounting)."""
    stage = DRAFT_STEP_ATTACH[step]
    return pdf_output_path(month=month, year=year, stage=stage)


def pdf_inbound_path(*, role: str, month: str, year: int) -> Path:
    """Path ปลายทางเมื่อดึง PDF จาก inbox reply."""
    stage = INBOUND_ROLE_STAGE[role]
    return pdf_output_path(month=month, year=year, stage=stage)


def snapshot_for_draft(*, step: str, month: str, year: int, source: Path) -> Path:
    """Copy PDF ไปโฟลเดอร์ draft ของ role นั้น (เก็บ snapshot ตอนสร้าง draft)."""
    stage = DRAFT_STEP_SNAPSHOT[step]
    dest = pdf_output_path(month=month, year=year, stage=stage)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, dest)
    return dest


def list_exportable_sheets(workbook_path: Path, months: tuple[str, ...] | None = None) -> list[str]:
    months = months or MONTH_SHEETS
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(workbook_path.resolve()))
        names = {s.name for s in wb.sheets}
        wb.close()
    finally:
        app.quit()
    return [name for name in months if name in names]


def export_sheet_to_pdf(
    *,
    source: Path,
    sheet_name: str,
    output: Path,
) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(source.resolve()))
        if sheet_name not in [s.name for s in wb.sheets]:
            raise ValueError(f"Sheet not found: {sheet_name} in {source.name}")
        wb.sheets[sheet_name].to_pdf(str(output.resolve()))
        wb.close()
    finally:
        app.quit()

    return output


def export_months(
    *,
    months: list[str],
    source: Path | None = None,
    year: int,
    stage: str = PDF_STAGE_EMPLOYEE,
    status: str | None = None,
    employee_tag: str = EMPLOYEE_FILE_TAG,
    dry_run: bool = False,
) -> list[PdfExportRecord]:
    ensure_dirs()
    if status is not None:
        legacy = {"pending": PDF_STAGE_EMPLOYEE, "sent": PDF_STAGE_ACCOUNTING_SENT}
        stage = legacy.get(status, status)

    source = Path(source or WORKING_FORMATTED)
    if not source.exists():
        raise FileNotFoundError(
            f"ไม่พบไฟล์ Excel: {source}\n"
            "รัน tools/excel/format_report.py หรือ sync template ก่อน"
        )

    available = set(list_exportable_sheets(source))
    records: list[PdfExportRecord] = []
    stamp = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

    for month in months:
        if month not in available:
            raise ValueError(f"ไม่พบชีต {month} ใน {source.name}")

        output = pdf_output_path(
            month=month,
            year=year,
            stage=stage,
            employee_tag=employee_tag,
        )
        if dry_run:
            print(f"[dry-run] {month} -> {output}")
            continue

        export_sheet_to_pdf(source=source, sheet_name=month, output=output)
        record = PdfExportRecord(
            month=month,
            year=year,
            sheet=month,
            source=str(source.resolve()),
            pdf=str(output.resolve()),
            stage=stage,
            exported_at=stamp,
        )
        records.append(record)
        print(f"Exported {month} -> {output}")

    if records and not dry_run:
        write_manifest(records, year=year, stage=stage)

    return records


def write_manifest(records: list[PdfExportRecord], *, year: int, stage: str) -> Path:
    ensure_dirs()
    payload = {
        "updated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "year": year,
        "stage": stage,
        "exports": [asdict(r) for r in records],
    }
    EXPORT_MANIFEST.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return EXPORT_MANIFEST
