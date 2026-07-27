"""Track timesheet email workflow per month (for agent + fetch tool)."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

from paths import EXPORT_DIR, ensure_dirs

WORKFLOW_STATE_PATH = EXPORT_DIR / "workflow_state.json"

ROLES = ("manager", "approver")


@dataclass
class RoleReceipt:
    role: str
    from_email: str
    from_name: str
    subject: str
    received_at: str
    message_id: str
    pdf_path: str
    pdf_bytes: int
    pending_bytes: int | None = None
    size_changed: bool | None = None
    verified: bool = False
    notes: str = ""


@dataclass
class MonthWorkflow:
    month: str
    year: int
    manager: RoleReceipt | None = None
    approver: RoleReceipt | None = None
    updated_at: str = ""

    def next_step(self) -> str:
        if self.approver is not None:
            return "draft_accounting"
        if self.manager is not None:
            return "draft_approver"
        return "await_manager_reply"


def _now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _key(month: str, year: int) -> str:
    return f"{year}/{month}"


def load_state() -> dict:
    ensure_dirs()
    if not WORKFLOW_STATE_PATH.exists():
        return {"updated_at": _now_iso(), "months": {}, "processed_message_ids": []}
    return json.loads(WORKFLOW_STATE_PATH.read_text(encoding="utf-8"))


def save_state(payload: dict) -> Path:
    ensure_dirs()
    payload["updated_at"] = _now_iso()
    WORKFLOW_STATE_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return WORKFLOW_STATE_PATH


def get_month_state(month: str, year: int) -> MonthWorkflow | None:
    data = load_state()
    raw = data.get("months", {}).get(_key(month, year))
    if not raw:
        return None
    def _role(key: str) -> RoleReceipt | None:
        r = raw.get(key)
        return RoleReceipt(**r) if r else None

    return MonthWorkflow(
        month=raw["month"],
        year=raw["year"],
        manager=_role("manager"),
        approver=_role("approver"),
        updated_at=raw.get("updated_at", ""),
    )


def record_receipt(
    *,
    month: str,
    year: int,
    receipt: RoleReceipt,
    message_id: str,
) -> MonthWorkflow:
    data = load_state()
    months = data.setdefault("months", {})
    key = _key(month, year)
    entry = months.get(key, {"month": month, "year": year})
    entry[receipt.role] = asdict(receipt)
    entry["updated_at"] = _now_iso()
    months[key] = entry

    processed: list[str] = data.setdefault("processed_message_ids", [])
    if message_id not in processed:
        processed.append(message_id)

    save_state(data)
    return MonthWorkflow(
        month=month,
        year=year,
        manager=RoleReceipt(**entry["manager"]) if entry.get("manager") else None,
        approver=RoleReceipt(**entry["approver"]) if entry.get("approver") else None,
        updated_at=entry["updated_at"],
    )


def is_message_processed(message_id: str) -> bool:
    data = load_state()
    return message_id in data.get("processed_message_ids", [])
