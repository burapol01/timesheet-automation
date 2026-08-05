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
# แถว footer ตาม report-layout-master.xlsx (July) — ห้าม offset เอง
FOOTER_BLOCK_FIRST_ROW = 42
FOOTER_BLOCK_LAST_ROW = 48
FOOTER_DATE_ROW = 48
FOOTER_DATE_COLUMNS = (1, 4, 7)  # A, D, G

# ขนาดรูปลายเซ็นพนักงาน (Picture 1) จาก report-layout-master.xlsx July
SIGNATURE_WIDTH = 124.39
SIGNATURE_HEIGHT = 53.65

# แถวว่างใต้ label ก่อน "ลงชื่อ" — วางรูปลายเซ็นพนักงาน
SIGNATURE_ANCHOR_ROW = 44
