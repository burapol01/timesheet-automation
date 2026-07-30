"""Footer / signature block — layout ยืนยันจาก report-formatted.xlsx (Jul 2026)."""

from __future__ import annotations

from pathlib import Path

from email_config import APPROVER_NAME_TH, MANAGER_NAME_TH, MANAGER_TITLE

ROOT = Path(__file__).resolve().parent.parent

PLACE_APPROVER_SIGNATURE = False

REVIEW_LABEL = "Review by"
APPROVE_LABEL = "Approve by"
MANAGER_FOOTER_NAME = MANAGER_NAME_TH
MANAGER_FOOTER_TITLE = MANAGER_TITLE
APPROVER_FOOTER_NAME = APPROVER_NAME_TH
APPROVER_FOOTER_ROLE = "Project Manager"
EMPLOYEE_TITLE = "Programmer"
EMPLOYEE_FILE_TAG = "Burapol"

# January master cells (อ้างอิงทุกเดือน)
JANUARY_MASTER_MANAGER_CELL = "F43"
JANUARY_MASTER_APPROVER_NAME_CELL = "H43"
JANUARY_MASTER_TITLE_CELL = "A44"

# legacy aliases — format_report imports
APPROVER_TITLE = MANAGER_FOOTER_TITLE
APPROVER_NAME = MANAGER_FOOTER_NAME
APPROVER_ROLE = APPROVER_FOOTER_ROLE

# จัดตำแหน่ง label/signature ให้ตรง 3 คอลัมน์ (Jul 2026)
REVIEW_LABEL_CELL = f"{' ' * 65}{REVIEW_LABEL}"
EMPLOYEE_SIG_LINE = "\n             ลงชื่อ ..............................................."
REVIEW_SIG_LINE = "\n             ลงชื่อ ..............................................."
APPROVE_SIG_LINE = "\n             ลงชื่อ ..............................................."

FOOTER_FONT_NAME = "Angsana New"
FOOTER_FONT_SIZE = 14.0
FOOTER_DATE_FONT_SIZE = 14.0
FOOTER_DATE_ROW = 47

# ตำแหน่งลายเซ็นพนักงาน (ยืนยัน manual report-formatted.xlsx Jul 2026)
SIGNATURE_ANCHOR_ROW = 42
SIGNATURE_TOP_OFFSET = 25.44
SIGNATURE_LEFT = 55.30
SIGNATURE_WIDTH = 122.84
SIGNATURE_HEIGHT = 55.39
