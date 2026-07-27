"""
ส่ง timesheet จาก 02-working/report-data.xlsx ไป intranet

อ่านคอลัมน์ รายละเอียดของงานที่ทำ -> Remark บนเว็บ
ข้ามวันที่มี entry อยู่แล้ว (วันละ 1 งาน)
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout, sync_playwright

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

from excel_report import ReportEntry, read_work_entries  # noqa: E402
from paths import WORKING_DATA  # noqa: E402
from timesheet_browser import (  # noqa: E402
    TIMESHEET_URL,
    TimesheetEntry,
    close_modal_if_open,
    ensure_edit_event_id,
    fill_timesheet_form,
    open_calendar_page,
    read_form_snapshot,
    save_timesheet,
)


def parse_args() -> argparse.Namespace:
    today = date.today()
    p = argparse.ArgumentParser(description="Submit intranet timesheets from report-data.xlsx")
    p.add_argument("--data", default=str(WORKING_DATA), help="Data file (default: 02-working/report-data.xlsx)")
    p.add_argument("--sheet", default="July")
    p.add_argument("--from-date", default="2026-07-01")
    p.add_argument("--to-date", default=today.isoformat())
    p.add_argument("--url", default=TIMESHEET_URL)
    p.add_argument("--job-type", default="Project")
    p.add_argument("--project", default="Fast Track Revamp")
    p.add_argument("--activity", default="Development")
    p.add_argument("--effort", default="8")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--no-save", action="store_true")
    p.add_argument(
        "--force",
        action="store_true",
        help="Update existing entry if the day already has one (never create duplicate)",
    )
    p.add_argument("--headless", action="store_true")
    return p.parse_args()


def to_timesheet(entry: ReportEntry, args: argparse.Namespace) -> TimesheetEntry:
    return TimesheetEntry(
        job_type=args.job_type,
        project=args.project,
        activity=args.activity,
        effort_hours=args.effort,
        event_date=entry.event_date,
        remark=entry.remark,
    )


def submit(entries: list[ReportEntry], args: argparse.Namespace) -> int:
    if not entries:
        print("No work entries in range.")
        return 0

    skip = not args.force
    ok = fail = skip_n = 0

    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel="msedge", headless=args.headless, slow_mo=80)
        page = browser.new_page(ignore_https_errors=True)
        try:
            session = open_calendar_page(page, args.url)
            for i, entry in enumerate(entries, 1):
                label = entry.event_date.isoformat()
                print(f"\n[{i}/{len(entries)}] {label}")
                if skip and session.event_count(entry.event_date) >= 1:
                    print("  skip: already has entry")
                    skip_n += 1
                    continue
                ts = to_timesheet(entry, args)
                try:
                    close_modal_if_open(page)
                    before_count = session.event_count(entry.event_date)
                    editing_event_id = session.open_entry_modal(
                        entry.event_date,
                        edit_if_exists=not skip,
                    )
                    fill_timesheet_form(page, ts)
                    snap = read_form_snapshot(page)
                    if snap["remark"] != ts.remark:
                        raise RuntimeError("Remark mismatch after fill")
                    if editing_event_id is not None:
                        ensure_edit_event_id(page, editing_event_id)
                    if args.no_save:
                        close_modal_if_open(page)
                        print("  filled (not saved)")
                    else:
                        save_timesheet(
                            page,
                            session,
                            editing_event_id=editing_event_id,
                        )
                        after_count = session.event_count(entry.event_date)
                        if editing_event_id is not None and after_count > before_count:
                            raise RuntimeError(
                                f"Duplicate created: events {before_count} -> {after_count}"
                            )
                        ok += 1
                        action = "updated" if editing_event_id is not None else "saved"
                        print(f"  {action} (event id={editing_event_id or 'new'})")
                except Exception as exc:
                    fail += 1
                    print(f"  ERROR: {exc}", file=sys.stderr)
                    close_modal_if_open(page)
        except PlaywrightTimeout as exc:
            print(f"Timeout: {exc}", file=sys.stderr)
            return 1
        finally:
            browser.close()

    print(f"\nDone. saved={ok} skipped={skip_n} failed={fail}")
    return 1 if fail else 0


def main() -> int:
    args = parse_args()
    fd = date.fromisoformat(args.from_date)
    td = date.fromisoformat(args.to_date)
    entries = read_work_entries(
        Path(args.data),
        sheet_name=args.sheet,
        from_date=fd,
        to_date=td,
    )
    print(f"Source: {args.data}")
    print(f"Found {len(entries)} work entries ({args.sheet}, {fd} .. {td})")
    for e in entries:
        print(f"  {e.event_date}: {e.remark[:80]}...")
    if args.dry_run:
        return 0
    return submit(entries, args)


if __name__ == "__main__":
    raise SystemExit(main())
