"""
Central file layout for timesheet-automation.

PDF workflow (04-export/pdf/)
-----------------------------
01-employee/              export จาก Excel (ยังไม่ส่ง)
02-manager/draft|signed/  draft → คุณสายฝน นามกูล | รับกลับหลัง Manager เซ็น
03-approver/draft|signed/ draft → Achara | รับกลับหลัง Achara เซ็น
04-accounting/draft|sent/ draft → IT-D | เก็บหลังส่งแล้ว
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

ORIGINAL_DIR = ROOT / "01-original"
WORKING_DIR = ROOT / "02-working"
ARCHIVE_DIR = ROOT / "03-archive"
EXPORT_DIR = ROOT / "04-export"
PDF_DIR = EXPORT_DIR / "pdf"

# ลำดับชั้นตาม role (ชื่อโฟลเดอร์สื่อขั้นตอน)
PDF_EMPLOYEE_DIR = PDF_DIR / "01-employee"
PDF_MANAGER_DIR = PDF_DIR / "02-manager"
PDF_APPROVER_DIR = PDF_DIR / "03-approver"
PDF_ACCOUNTING_DIR = PDF_DIR / "04-accounting"

SUB_DRAFT = "draft"
SUB_SIGNED = "signed"
SUB_SENT = "sent"

# Legacy aliases (deprecated — ใช้ pdf_path stage แทน)
PDF_PENDING_DIR = PDF_EMPLOYEE_DIR
PDF_SIGNED_DIR = PDF_APPROVER_DIR / SUB_SIGNED
PDF_SENT_DIR = PDF_ACCOUNTING_DIR / SUB_SENT

ORIGINAL_TEMPLATE = ORIGINAL_DIR / "project-report-template.xlsx"
WORKING_DATA = WORKING_DIR / "report-data.xlsx"
WORKING_FORMATTED = WORKING_DIR / "report-formatted.xlsx"
FORMAT_REFERENCE_SHEET = "June"

MONTH_SHEETS = (
    "January",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)

DEFAULT_REPORT_YEAR = 2026
EXPORT_MANIFEST = EXPORT_DIR / "manifest.json"
DEFAULT_EXCEL_PATH = WORKING_DATA

# stage → (root_dir, subfolder or None for flat)
_PDF_STAGE_DIRS: dict[str, tuple[Path, str | None]] = {
    "employee": (PDF_EMPLOYEE_DIR, None),
    "manager_draft": (PDF_MANAGER_DIR, SUB_DRAFT),
    "manager_signed": (PDF_MANAGER_DIR, SUB_SIGNED),
    "approver_draft": (PDF_APPROVER_DIR, SUB_DRAFT),
    "approver_signed": (PDF_APPROVER_DIR, SUB_SIGNED),
    "accounting_draft": (PDF_ACCOUNTING_DIR, SUB_DRAFT),
    "accounting_sent": (PDF_ACCOUNTING_DIR, SUB_SENT),
}

# role ที่รับ reply จาก inbox → stage ปลายทาง
INBOUND_ROLE_STAGE = {
    "manager": "manager_signed",
    "approver": "approver_signed",
}

# draft step → stage ของ PDF ที่แนบ + stage draft snapshot
DRAFT_STEP_ATTACH = {
    "manager": "employee",
    "approver": "manager_signed",
    "accounting": "approver_signed",
}

DRAFT_STEP_SNAPSHOT = {
    "manager": "manager_draft",
    "approver": "approver_draft",
    "accounting": "accounting_draft",
}


def ensure_dirs() -> None:
    for root, sub in _PDF_STAGE_DIRS.values():
        if sub:
            (root / sub).mkdir(parents=True, exist_ok=True)
        else:
            root.mkdir(parents=True, exist_ok=True)
    for d in (ORIGINAL_DIR, WORKING_DIR, ARCHIVE_DIR, EXPORT_DIR, PDF_DIR):
        d.mkdir(parents=True, exist_ok=True)


def pdf_stage_dir(*, stage: str, year: int) -> Path:
    try:
        root, sub = _PDF_STAGE_DIRS[stage]
    except KeyError as exc:
        raise ValueError(f"Unknown PDF stage: {stage}") from exc
    base = root / sub if sub else root
    path = base / str(year)
    path.mkdir(parents=True, exist_ok=True)
    return path


def pdf_pending_dir(year: int) -> Path:
    """Legacy — ใช้ pdf_stage_dir(stage='employee')."""
    return pdf_stage_dir(stage="employee", year=year)


def pdf_signed_dir(year: int) -> Path:
    """Legacy — approver signed folder."""
    return pdf_stage_dir(stage="approver_signed", year=year)


def pdf_sent_dir(year: int) -> Path:
    """Legacy — accounting sent folder."""
    return pdf_stage_dir(stage="accounting_sent", year=year)


def assert_original_readonly() -> None:
    if not ORIGINAL_TEMPLATE.exists():
        raise FileNotFoundError(
            f"ไม่พบไฟล์ต้นฉบับ: {ORIGINAL_TEMPLATE}\n"
            "วาง template จากองค์กรใน 01-original/ (ห้ามแก้ไขไฟล์นี้)"
        )
