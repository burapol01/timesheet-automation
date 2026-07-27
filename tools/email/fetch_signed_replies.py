"""
ดึง PDF ลงนามจากอีเมลตอบกลับ (Outlook Inbox)

บันทึกตาม role:
  Saifon  → 02-manager/signed/..._manager_signed.pdf
  Achara  → 03-approver/signed/..._approver_signed.pdf

  python tools/email/fetch_signed_replies.py --dry-run
  python tools/email/fetch_signed_replies.py --month June --year 2026
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from email_config import APPROVER_DISPLAY, MANAGER_DISPLAY, ROLE_BY_EMAIL  # noqa: E402
from outlook_inbox import InboundMail, debug_scan_inbox, save_attachment, scan_inbox  # noqa: E402
from paths import EXPORT_DIR, ensure_dirs  # noqa: E402
from pdf_export import PDF_STAGE_EMPLOYEE, pdf_inbound_path, pdf_output_path  # noqa: E402
from workflow_state import (  # noqa: E402
    RoleReceipt,
    WORKFLOW_STATE_PATH,
    get_month_state,
    is_message_processed,
    record_receipt,
)

FETCH_REPORT_PATH = EXPORT_DIR / "inbox_fetch_report.json"
ROLE_LABEL = {"manager": MANAGER_DISPLAY, "approver": APPROVER_DISPLAY}


def _role_for_mail(mail: InboundMail) -> str | None:
    return ROLE_BY_EMAIL.get(mail.sender_email)


def _employee_size(month: str, year: int) -> int | None:
    path = pdf_output_path(month=month, year=year, stage=PDF_STAGE_EMPLOYEE)
    if path.exists():
        return path.stat().st_size
    return None


def _backup_if_exists(path: Path) -> Path | None:
    if not path.exists():
        return None
    stamp = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.stem}_backup_{stamp}{path.suffix}")
    shutil.copy2(path, backup)
    return backup


def process_mail(
    mail: InboundMail,
    *,
    dry_run: bool,
) -> dict:
    role = _role_for_mail(mail)
    if not role:
        return {"status": "skip", "reason": "unknown_sender", "subject": mail.subject}

    month = mail.pdf.month_en
    year = mail.pdf.year
    dest = pdf_inbound_path(role=role, month=month, year=year)
    employee_bytes = _employee_size(month, year)
    backup: Path | None = None
    signed_bytes: int | None = None

    if dry_run:
        action = "would_save"
    else:
        backup = _backup_if_exists(dest)
        save_attachment(mail, dest)
        signed_bytes = dest.stat().st_size
        action = "saved"

    size_changed = None
    if employee_bytes is not None and signed_bytes is not None:
        size_changed = signed_bytes != employee_bytes

    receipt = RoleReceipt(
        role=role,
        from_email=mail.sender_email,
        from_name=mail.sender_name,
        subject=mail.subject,
        received_at=mail.received_at.isoformat(timespec="seconds"),
        message_id=f"{mail.store_id}:{mail.entry_id}",
        pdf_path=str(dest),
        pdf_bytes=signed_bytes or 0,
        pending_bytes=employee_bytes,
        size_changed=size_changed,
        verified=False,
        notes="รอ agent/ผู้ใช้ตรวจลายเซ็น",
    )

    if not dry_run:
        record_receipt(
            month=month,
            year=year,
            receipt=receipt,
            message_id=receipt.message_id,
        )

    wf = get_month_state(month, year)
    if wf:
        next_step = wf.next_step()
    elif role == "manager":
        next_step = "draft_approver"
    else:
        next_step = "draft_accounting"

    return {
        "status": action,
        "role": role,
        "role_label": ROLE_LABEL.get(role, role),
        "month": month,
        "year": year,
        "from": mail.sender_name,
        "email": mail.sender_email,
        "subject": mail.subject,
        "received_at": receipt.received_at,
        "saved_to": str(dest),
        "backup": str(backup) if backup else None,
        "employee_bytes": employee_bytes,
        "signed_bytes": signed_bytes,
        "size_changed": size_changed,
        "body_preview": mail.body_preview[:200],
        "next_step": next_step,
        "needs_review": True,
    }


def build_report(results: list[dict], *, dry_run: bool) -> dict:
    return {
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "dry_run": dry_run,
        "matched": len(results),
        "results": results,
        "workflow_state": str(WORKFLOW_STATE_PATH),
    }


def print_report(report: dict) -> None:
    print("=== Inbox fetch report ===")
    print(f"Matched: {report['matched']}  dry_run={report['dry_run']}")
    for r in report["results"]:
        print()
        print(f"[{r['status']}] {r['month']} {r['year']} — {r['role_label']} ({r['email']})")
        print(f"  Subject: {r['subject']}")
        print(f"  Saved:   {r['saved_to']}")
        if r.get("size_changed") is not None:
            flag = "OK (size changed)" if r["size_changed"] else "WARN same size as employee export"
            print(f"  Size:    employee={r['employee_bytes']} signed={r['signed_bytes']} {flag}")
        if r.get("body_preview"):
            print(f"  Body:    {r['body_preview'][:80]}...")
        print(f"  Next:    {r['next_step']}  needs_review={r['needs_review']}")
    print()
    print(f"Report JSON: {FETCH_REPORT_PATH}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch signed PDF attachments from Outlook inbox replies"
    )
    parser.add_argument("--month", help="Filter sheet name e.g. June")
    parser.add_argument("--year", type=int, help="Filter year e.g. 2026")
    parser.add_argument("--days", type=int, default=90, help="Scan inbox N days back")
    parser.add_argument("--dry-run", action="store_true", help="Scan only, do not save")
    parser.add_argument(
        "--include-processed",
        action="store_true",
        help="Re-process messages already in workflow_state",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Show near-miss inbox messages (diagnostics)",
    )
    args = parser.parse_args()

    ensure_dirs()

    if args.debug:
        hits = debug_scan_inbox(days_back=args.days)
        print("=== Inbox debug (near-miss) ===")
        for h in hits:
            print()
            print(f"from: {h['from']} <{h['email']}>  role={h['role']}")
            print(f"subj: {h['subject'][:100]}")
            print(f"recv: {h['received']}")
            print(f"pdfs: {h['pdfs']}")
            print(f"parse: {h['parsed']}")
        print(f"\nShown: {len(hits)}")
        return

    mails = scan_inbox(
        days_back=args.days,
        month_filter=args.month,
        year_filter=args.year,
        include_processed=args.include_processed,
        is_processed=is_message_processed,
    )

    results: list[dict] = []
    for mail in mails:
        results.append(process_mail(mail, dry_run=args.dry_run))

    report = build_report(results, dry_run=args.dry_run)
    FETCH_REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print_report(report)


if __name__ == "__main__":
    main()
