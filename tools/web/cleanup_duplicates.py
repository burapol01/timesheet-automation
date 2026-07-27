"""List and remove duplicate timesheet calendar entries."""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from timesheet_browser import TIMESHEET_URL, open_calendar_page  # noqa: E402


@dataclass
class CalendarEvent:
    event_id: int
    day: date
    title: str
    description: str

    @property
    def remark_score(self) -> int:
        text = (self.description or "").strip()
        if not text:
            return 0
        if text.lower() in {"fast track revamp development work", "fast track revamp"}:
            return 1
        return len(text)


def fetch_events(page) -> list[CalendarEvent]:
    with page.expect_response(
        lambda r: "GetEvents" in r.url and r.ok,
        timeout=30_000,
    ) as resp_info:
        page.reload(wait_until="domcontentloaded")
    data = resp_info.value.json()
    events: list[CalendarEvent] = []
    for row in data:
        raw = row.get("dateStart") or row.get("DateStart") or ""
        if isinstance(raw, str) and "T" in raw:
            day = datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
        else:
            day = datetime.fromisoformat(str(raw)[:10]).date()
        events.append(
            CalendarEvent(
                event_id=int(row.get("eventId") or row.get("EventId")),
                day=day,
                title=str(row.get("projectName") or row.get("ProjectName") or ""),
                description=str(row.get("Description") or row.get("description") or ""),
            )
        )
    return events


def choose_keepers(events: list[CalendarEvent]) -> dict[date, CalendarEvent]:
    by_day: dict[date, list[CalendarEvent]] = defaultdict(list)
    for ev in events:
        by_day[ev.day].append(ev)

    keepers: dict[date, CalendarEvent] = {}
    for day, day_events in by_day.items():
        if len(day_events) == 1:
            keepers[day] = day_events[0]
            continue
        ranked = sorted(
            day_events,
            key=lambda e: (e.remark_score, e.event_id),
            reverse=True,
        )
        keepers[day] = ranked[0]
    return keepers


def delete_event(page, event_id: int) -> None:
    ok = page.evaluate(
        """async (eventId) => {
            const res = await fetch('/timesheet/timesheets/DeleteEvent', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json; charset=utf-8' },
                body: JSON.stringify({ id: eventId }),
                credentials: 'same-origin',
            });
            return res.ok;
        }""",
        event_id,
    )
    if not ok:
        raise RuntimeError(f"DeleteEvent {event_id} failed")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Remove duplicate intranet timesheet entries.")
    parser.add_argument("--from-date", default="2026-07-01")
    parser.add_argument("--to-date", default="2026-07-31")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--headless", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    from_date = date.fromisoformat(args.from_date)
    to_date = date.fromisoformat(args.to_date)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=args.headless, slow_mo=60)
        page = browser.new_page(ignore_https_errors=True)
        session = open_calendar_page(page, TIMESHEET_URL)

        all_events = fetch_events(page)
        scoped = [e for e in all_events if from_date <= e.day <= to_date]
        keepers = choose_keepers(scoped)

        to_delete: list[CalendarEvent] = []
        for ev in scoped:
            keeper = keepers[ev.day]
            if ev.event_id != keeper.event_id:
                to_delete.append(ev)

        print(f"Events in range: {len(scoped)}")
        dup_days = sorted({e.day for e in to_delete})
        print(f"Duplicate days: {len(dup_days)}")
        for day in dup_days:
            day_events = [e for e in scoped if e.day == day]
            keeper = keepers[day]
            print(f"\n{day} ({len(day_events)} entries) -> keep id={keeper.event_id}")
            for ev in day_events:
                mark = "KEEP" if ev.event_id == keeper.event_id else "DELETE"
                desc = (ev.description or "(empty)")[:80]
                print(f"  [{mark}] id={ev.event_id}  {desc}")

        if args.dry_run:
            print(f"\nDry run: would delete {len(to_delete)} entries")
            browser.close()
            return 0

        deleted = 0
        for ev in to_delete:
            delete_event(page, ev.event_id)
            deleted += 1
            print(f"Deleted id={ev.event_id} on {ev.day}")

        page.reload(wait_until="domcontentloaded")
        session.wait_calendar_idle()
        remaining = fetch_events(page)
        remaining_scoped = [e for e in remaining if from_date <= e.day <= to_date]
        by_day: dict[date, list[CalendarEvent]] = defaultdict(list)
        for e in remaining_scoped:
            by_day[e.day].append(e)
        still_dup = [d for d, evs in by_day.items() if len(evs) > 1]

        print(f"\nDone. deleted={deleted}")
        if still_dup:
            print(f"WARNING: still duplicated days: {still_dup}", file=sys.stderr)
            browser.close()
            return 1

        browser.close()
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
