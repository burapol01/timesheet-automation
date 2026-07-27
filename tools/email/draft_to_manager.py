"""
ขั้น 1: สร้าง Outlook draft ส่ง PDF ให้ Manager (Saifon) ลงนาม

- To: saifon.nam@tokiomarinelife.co.th เท่านั้น (ไม่มี CC)
- แนบ PDF จาก 04-export/pdf/01-employee/
- บันทึกใน Drafts — ไม่ส่งจริง

ตัวอย่าง:
  python tools/email/draft_to_manager.py --month July --dry-run
  python tools/email/draft_to_manager.py --month July
  python tools/email/draft_to_manager.py --month July --open
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from email_config import MANAGER_DISPLAY, MANAGER_EMAIL  # noqa: E402
from email_templates import body_manager_review, subject_manager_review  # noqa: E402
from outlook_draft import DraftMail, create_outlook_draft  # noqa: E402
from paths import DEFAULT_REPORT_YEAR, ensure_dirs  # noqa: E402
from pdf_export import pdf_output_path, PDF_STATUS_PENDING  # noqa: E402


def resolve_pdf(*, month: str, year: int, pdf: Path | None) -> Path:
    if pdf is not None:
        path = pdf.resolve()
        if not path.exists():
            raise FileNotFoundError(f"ไม่พบ PDF: {path}")
        return path

    path = pdf_output_path(month=month, year=year, status=PDF_STATUS_PENDING)
    if not path.exists():
        raise FileNotFoundError(
            f"ไม่พบ PDF: {path}\n"
            f"รัน: python tools/excel/export_pdf.py --month {month}"
        )
    return path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create Outlook draft: PDF report → Manager (Saifon), no CC"
    )
    parser.add_argument("--month", required=True, help="Sheet name e.g. July")
    parser.add_argument("--year", type=int, default=DEFAULT_REPORT_YEAR)
    parser.add_argument("--pdf", type=Path, help="Override PDF path")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="แสดง draft ที่จะสร้าง ไม่เปิด Outlook",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="เปิดหน้าต่าง draft ใน Outlook หลังบันทึก",
    )
    args = parser.parse_args()

    ensure_dirs()
    pdf_path = resolve_pdf(month=args.month, year=args.year, pdf=args.pdf)

    mail = DraftMail(
        to=MANAGER_EMAIL,
        cc="",
        subject=subject_manager_review(month_en=args.month, year=args.year),
        body=body_manager_review(month_en=args.month, year=args.year),
        attachments=[pdf_path],
        display_to=MANAGER_DISPLAY,
    )

    if args.dry_run:
        print("=== DRY RUN — ไม่สร้าง draft ใน Outlook ===")
        print(f"To:      {mail.display_to} <{mail.to}>")
        print(f"Cc:      (none)")
        print(f"Subject: {mail.subject}")
        print(f"Attach:  {pdf_path}")
        print("--- Body ---")
        print(mail.body.replace("\r\n", "\n"))
        return

    create_outlook_draft(mail, open_draft=args.open)
    print("Saved Outlook draft to Manager:", MANAGER_EMAIL)
    print("Subject:", mail.subject)
    print("Attach:", pdf_path.name)
    if not args.open:
        print("Open Outlook > Drafts to review before sending")


if __name__ == "__main__":
    main()
