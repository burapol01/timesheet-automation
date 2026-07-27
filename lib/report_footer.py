"""Footer / signature block — layout ยืนยันจาก report-formatted.xlsx (Jul 2026)."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

PLACE_APPROVER_SIGNATURE = False

REVIEW_LABEL = "Review by"
APPROVE_LABEL = "Approve by"
APPROVER_TITLE = "Manager - Fasttrack Development"
APPROVER_NAME = "คุณอัจฉรา ชัยภูมิ"
APPROVER_ROLE = "Project Manager"
EMPLOYEE_TITLE = "Programmer"
EMPLOYEE_FILE_TAG = "Burapol"

# ช่อง master บนชีต January (ห้ามทับด้วยเส้นลายเซ็น)
JANUARY_MASTER_APPROVER_CELL = "F43"
JANUARY_MASTER_TITLE_CELL = "A44"

# จัดตำแหน่งด้วยช่องว่าง (ยืนยัน manual Jul 2026)
REVIEW_LABEL_CELL = f"{' ' * 65}{REVIEW_LABEL}"
EMPLOYEE_SIG_LINE = "\n             ลงชื่อ ..............................................."
REVIEW_SIG_LINE = (
    "\n                                                                 "
    "ลงชื่อ ..............................................."
)
APPROVE_SIG_LINE = "\nลงชื่อ ..............................................."

FOOTER_DATE_FONT_SIZE = 16.0

# ตำแหน่งลายเซ็นพนักงาน (ยืนยัน manual report-formatted.xlsx Jul 2026)
SIGNATURE_ANCHOR_ROW = 42
SIGNATURE_TOP_OFFSET = 25.44
SIGNATURE_LEFT = 55.30
SIGNATURE_WIDTH = 122.84
SIGNATURE_HEIGHT = 55.39
