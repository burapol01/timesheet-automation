"""Create Outlook draft emails (Windows + Outlook desktop)."""

from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path


@dataclass
class DraftMail:
    to: str
    cc: str
    subject: str
    body: str
    attachments: list[Path]
    display_to: str | None = None


def _outlook_application():
    try:
        import win32com.client  # type: ignore[import-untyped]
    except ImportError as exc:
        raise RuntimeError(
            "ต้องติดตั้ง pywin32 และใช้ Outlook บน Windows\n"
            "  pip install pywin32"
        ) from exc

    last_error: Exception | None = None
    for factory in (
        lambda: win32com.client.gencache.EnsureDispatch("Outlook.Application"),
        lambda: win32com.client.Dispatch("Outlook.Application"),
        lambda: win32com.client.DispatchEx("Outlook.Application"),
    ):
        try:
            return factory()
        except Exception as exc:
            last_error = exc
            continue

    detail = f" ({last_error})" if last_error else ""
    raise RuntimeError(
        "เปิด Outlook (Classic) ไม่ได้ — ตรวจว่าติดตั้ง Microsoft Outlook desktop"
        f"{detail}\n"
        "ถ้าใช้ Outlook ใหม่ (New Outlook) ให้สลับกลับ Classic Outlook แล้วลองใหม่\n"
        "หรือใช้ --dry-run ดูข้อความแล้ว copy เอง"
    )


def _com_retry(fn, *, retries: int = 5, delay: float = 2.0):
    last: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as exc:
            last = exc
            err = str(exc)
            if "rejected by callee" in err or "-2147418111" in err:
                if attempt < retries:
                    time.sleep(delay * attempt)
                    continue
            raise
    raise RuntimeError(f"Outlook COM failed: {last}")


def create_outlook_draft(
    mail: DraftMail,
    *,
    open_draft: bool = False,
) -> None:
    """Save mail to Outlook Drafts folder. Does not send."""
    outlook = _com_retry(_outlook_application)
    item = _com_retry(lambda: outlook.CreateItem(0))  # olMailItem

    if mail.display_to:
        item.To = f"{mail.display_to} <{mail.to}>"
    else:
        item.To = mail.to

    if mail.cc:
        item.CC = mail.cc

    item.Subject = mail.subject
    item.Body = mail.body

    for path in mail.attachments:
        if not path.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์แนบ: {path}")
        item.Attachments.Add(str(path.resolve()))

    item.Save()

    if open_draft:
        item.Display()
