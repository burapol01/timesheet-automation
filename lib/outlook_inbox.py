"""Read Outlook inbox for signed timesheet reply emails."""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable
from datetime import datetime, timedelta, timezone
from pathlib import Path

from email_config import ROLE_BY_EMAIL
from email_match import (
    ParsedPdfAttachment,
    normalize_email,
    parse_pdf_filename,
    parse_subject_month_year,
)
from outlook_draft import _outlook_application

OL_FOLDER_INBOX = 6
OL_MAIL = 43

_COM_RETRIES = 5
_COM_RETRY_DELAY = 2.0


def _com_retry(fn, *, label: str = "Outlook"):
    last: Exception | None = None
    for attempt in range(1, _COM_RETRIES + 1):
        try:
            return fn()
        except Exception as exc:
            last = exc
            err = str(exc)
            if "rejected by callee" in err or "RPC_E" in err or "-2147418111" in err:
                if attempt < _COM_RETRIES:
                    time.sleep(_COM_RETRY_DELAY * attempt)
                    continue
            raise
    raise RuntimeError(f"{label} COM failed after {_COM_RETRIES} tries: {last}")


@dataclass
class InboundMail:
    entry_id: str
    store_id: str
    subject: str
    sender_email: str
    sender_name: str
    received_at: datetime
    body_preview: str
    pdf: ParsedPdfAttachment
    pdf_attachment_index: int


def _smtp_address(mail_item) -> str:
    """Best-effort SMTP address from Outlook MailItem."""
    for getter in (
        lambda: mail_item.Sender.GetExchangeUser().PrimarySmtpAddress,
        lambda: mail_item.SenderEmailAddress,
        lambda: mail_item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        ),
    ):
        try:
            value = getter()
            if value and "@" in str(value) and not str(value).startswith("/O="):
                return normalize_email(str(value))
        except Exception:
            continue
    try:
        return normalize_email(str(mail_item.SenderEmailAddress))
    except Exception:
        return ""


def _received_dt(mail_item) -> datetime:
    try:
        return mail_item.ReceivedTime.replace(tzinfo=timezone.utc).astimezone()
    except Exception:
        return datetime.now(timezone.utc).astimezone()


def _find_pdf_attachment(mail_item) -> tuple[int, ParsedPdfAttachment] | None:
    try:
        attachments = mail_item.Attachments
        count = attachments.Count
    except Exception:
        return None

    for i in range(1, count + 1):
        att = attachments.Item(i)
        name = str(att.FileName)
        parsed = parse_pdf_filename(name)
        if parsed:
            return i, parsed
    return None


def _match_role(sender_email: str) -> str | None:
    role = ROLE_BY_EMAIL.get(sender_email)
    if role:
        return role
    local = sender_email.split("@", 1)[0]
    for email, r in ROLE_BY_EMAIL.items():
        if email.split("@", 1)[0] == local:
            return r
    return None


def scan_inbox(
    *,
    days_back: int = 90,
    month_filter: str | None = None,
    year_filter: int | None = None,
    include_processed: bool = False,
    is_processed: Callable[[str], bool] | None = None,
) -> list[InboundMail]:
    """Scan inbox for signed PDF replies from manager / approver."""
    outlook = _com_retry(_outlook_application, label="Outlook.Application")
    namespace = _com_retry(lambda: outlook.GetNamespace("MAPI"), label="Outlook MAPI")
    inbox = _com_retry(lambda: namespace.GetDefaultFolder(OL_FOLDER_INBOX), label="Inbox")
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now(timezone.utc).astimezone() - timedelta(days=days_back)
    results: list[InboundMail] = []

    for item in items:
        try:
            if item.Class != OL_MAIL:
                continue
        except Exception:
            continue

        received = _received_dt(item)
        if received < cutoff:
            break

        sender = _smtp_address(item)
        role = _match_role(sender)
        if not role:
            continue

        pdf_hit = _find_pdf_attachment(item)
        if not pdf_hit:
            continue

        att_index, parsed = pdf_hit

        if month_filter and parsed.month_en != month_filter:
            continue
        if year_filter and parsed.year != year_filter:
            continue

        try:
            entry_id = str(item.EntryID)
            store_id = str(item.Parent.StoreID)
        except Exception:
            continue

        message_id = f"{store_id}:{entry_id}"
        if not include_processed and is_processed and is_processed(message_id):
            continue

        subject = str(item.Subject or "")
        subj_parse = parse_subject_month_year(subject)
        if subj_parse:
            subj_month, subj_year = subj_parse
            if subj_month != parsed.month_en or subj_year != parsed.year:
                # subject vs filename mismatch — still accept but agent should note
                pass

        body = ""
        try:
            body = str(item.Body or "")[:500]
        except Exception:
            pass

        results.append(
            InboundMail(
                entry_id=entry_id,
                store_id=store_id,
                subject=subject,
                sender_email=sender,
                sender_name=str(getattr(item, "SenderName", "") or ""),
                received_at=received,
                body_preview=body.strip(),
                pdf=parsed,
                pdf_attachment_index=att_index,
            )
        )

    return results


def debug_scan_inbox(*, days_back: int = 180, limit: int = 20) -> list[dict]:
    """List recent emails with PDF attachments or from known senders (diagnostics)."""
    outlook = _com_retry(_outlook_application, label="Outlook.Application")
    namespace = _com_retry(lambda: outlook.GetNamespace("MAPI"), label="Outlook MAPI")
    inbox = _com_retry(lambda: namespace.GetDefaultFolder(OL_FOLDER_INBOX), label="Inbox")
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    cutoff = datetime.now(timezone.utc).astimezone() - timedelta(days=days_back)
    hits: list[dict] = []

    for item in items:
        if len(hits) >= limit:
            break
        try:
            if item.Class != OL_MAIL:
                continue
        except Exception:
            continue
        received = _received_dt(item)
        if received < cutoff:
            break

        sender = _smtp_address(item)
        role = _match_role(sender)
        pdfs: list[str] = []
        parsed_pdfs: list[str] = []
        try:
            for i in range(1, item.Attachments.Count + 1):
                name = str(item.Attachments.Item(i).FileName)
                if name.lower().endswith(".pdf"):
                    pdfs.append(name)
                    p = parse_pdf_filename(name)
                    parsed_pdfs.append(str(p) if p else "no_match")
        except Exception:
            pass

        subject = str(item.Subject or "")
        if not (role or pdfs or "รายงาน" in subject or "Burapol" in subject):
            continue

        hits.append(
            {
                "from": str(getattr(item, "SenderName", "") or ""),
                "email": sender,
                "role": role,
                "subject": subject,
                "received": received.isoformat(timespec="seconds"),
                "pdfs": pdfs,
                "parsed": parsed_pdfs,
            }
        )
    return hits


def save_attachment(
    mail: InboundMail,
    dest: Path,
    *,
    outlook=None,
) -> Path:
    """Save PDF attachment from a matched inbox item."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()

    app = outlook or _com_retry(_outlook_application, label="Outlook.Application")
    namespace = _com_retry(lambda: app.GetNamespace("MAPI"), label="Outlook MAPI")
    item = _com_retry(
        lambda: namespace.GetItemFromID(mail.entry_id, mail.store_id),
        label="GetItemFromID",
    )
    att = item.Attachments.Item(mail.pdf_attachment_index)
    att.SaveAsFile(str(dest.resolve()))
    return dest
