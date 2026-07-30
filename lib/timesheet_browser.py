"""Playwright helpers for TMLTH intranet timesheet."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

from playwright.sync_api import Page

TIMESHEET_URL = "http://intranet02/timesheet/Timesheets"
CALENDAR_LINK = 'a[href="/timesheet/Timesheets"]'
MODAL_TITLE = "Timesheet logging"


@dataclass
class TimesheetEntry:
    job_type: str = "Project"
    project: str = "Fast Track Revamp"
    activity: str = "Development"
    effort_hours: str = "8"
    event_date: date | None = None
    remark: str = ""


@dataclass
class CalendarSession:
    page: Page
    visible_month: date | None = field(default=None, init=False)
    editing_event_id: int | None = field(default=None, init=False)

    def _read_visible_month(self) -> date:
        title = self.page.locator(".fc-center h2").inner_text().strip()
        return datetime.strptime(title, "%B %Y").date().replace(day=1)

    def ensure_month(self, target: date) -> None:
        target_month = target.replace(day=1)
        if self.visible_month == target_month:
            return
        guard = 0
        visible = self._read_visible_month()
        while visible != target_month and guard < 24:
            if visible < target_month:
                self.page.locator(".fc-next-button").click()
            else:
                self.page.locator(".fc-prev-button").click()
            self.page.wait_for_timeout(300)
            visible = self._read_visible_month()
            guard += 1
        if visible != target_month:
            raise RuntimeError(f"Could not navigate calendar to {target.strftime('%B %Y')}")
        self.visible_month = target_month

    def event_count(self, target: date) -> int:
        self.ensure_month(target)
        iso = target.isoformat()
        return int(
            self.page.evaluate(
                """(iso) => {
                    const cal = window.jQuery ? jQuery('#calendar') : null;
                    if (!cal || !cal.length || typeof cal.fullCalendar !== 'function') {
                        return document.querySelectorAll(`td.fc-day[data-date="${iso}"] .fc-event`).length;
                    }
                    return cal.fullCalendar('clientEvents', (ev) => {
                        return moment(ev.start).format('YYYY-MM-DD') === iso;
                    }).length;
                }""",
                iso,
            )
        )

    def open_logging_modal(self, target: date) -> None:
        self.ensure_month(target)
        iso = target.isoformat()
        cell = self.page.locator(f'[data-date="{iso}"]').first
        if cell.count():
            cell.click(force=True)
        else:
            day = self.page.locator(
                ".fc-daygrid-day:not(.fc-day-other) .fc-daygrid-day-number"
            ).filter(has_text=str(target.day)).first
            if not day.count():
                raise RuntimeError(f"Calendar cell not found for {iso}")
            day.click()
        self.page.wait_for_selector(".modal.in, .modal.show", state="visible", timeout=10_000)
        self.page.wait_for_selector(f".modal-title:has-text('{MODAL_TITLE}')", timeout=10_000)

    def open_existing_event_modal(self, target: date) -> int:
        """Open edit modal via ShowEditData; returns event id (update path, not create)."""
        self.ensure_month(target)
        iso = target.isoformat()
        result = self.page.evaluate(
            """(iso) => {
                return new Promise((resolve, reject) => {
                    const cal = window.jQuery ? jQuery('#calendar') : null;
                    if (!cal || !cal.length || typeof cal.fullCalendar !== 'function') {
                        resolve({ ok: false, reason: 'calendar unavailable' });
                        return;
                    }
                    const events = cal.fullCalendar('clientEvents', (ev) =>
                        moment(ev.start).format('YYYY-MM-DD') === iso);
                    if (!events.length) {
                        resolve({ ok: false, reason: 'no events' });
                        return;
                    }
                    events.sort((a, b) => {
                        const score = (ev) => ((ev.description || ev.title || '') + '').length;
                        return score(b) - score(a);
                    });
                    const picked = events[0];
                    const id = parseInt(picked.eventID || picked.id, 10);
                    sessionStorage.setItem('Id', String(id));
                    jQuery('#myModal5').modal('show');
                    jQuery.ajax({
                        type: 'POST',
                        contentType: 'application/json; charset=utf-8',
                        url: '/timesheet/timesheets/ShowEditData',
                        data: JSON.stringify({ id }),
                        success: () => resolve({ ok: true, id }),
                        error: (_xhr, _status, err) => reject(err || 'ShowEditData failed'),
                    });
                });
            }""",
            iso,
        )
        if not result.get("ok"):
            raise RuntimeError(f"No calendar event to edit for {iso} ({result.get('reason', 'unknown')})")
        event_id = int(result["id"])
        self.editing_event_id = event_id
        self.page.wait_for_selector(".modal.in, .modal.show", state="visible", timeout=10_000)
        self.page.wait_for_selector(f".modal-title:has-text('{MODAL_TITLE}')", timeout=10_000)
        self.page.wait_for_selector("#comment", state="visible", timeout=10_000)
        self.page.wait_for_timeout(800)
        ensure_edit_event_id(self.page, event_id)
        return event_id

    def open_entry_modal(self, target: date, *, edit_if_exists: bool = False) -> int | None:
        """Open modal. Returns event id when updating, None when creating."""
        self.editing_event_id = None
        if edit_if_exists and self.event_count(target) >= 1:
            return self.open_existing_event_modal(target)
        ensure_new_event(self.page)
        self.open_logging_modal(target)
        return None

    def wait_calendar_idle(self) -> None:
        try:
            self.page.wait_for_response(
                lambda r: "GetEvents" in r.url and r.ok,
                timeout=10_000,
            )
        except Exception:
            self.page.wait_for_timeout(500)


def ensure_edit_event_id(page: Page, event_id: int) -> None:
    stored = page.evaluate(
        """(expected) => {
            sessionStorage.setItem('Id', String(expected));
            return parseInt(sessionStorage.getItem('Id') || '0', 10);
        }""",
        event_id,
    )
    if stored != event_id:
        raise RuntimeError(f"sessionStorage Id mismatch: expected {event_id}, got {stored}")


def ensure_new_event(page: Page) -> None:
    page.evaluate("""() => { sessionStorage.setItem('Id', '0'); }""")


def open_calendar_page(page: Page, url: str = TIMESHEET_URL) -> CalendarSession:
    page.goto(url, wait_until="networkidle", timeout=120_000)
    if page.locator("#calendar").count() == 0:
        page.locator(CALENDAR_LINK).click()
        page.wait_for_load_state("networkidle")
    page.wait_for_selector("#calendar", state="visible", timeout=60_000)
    page.wait_for_selector(".fc-view-container", state="visible", timeout=30_000)
    session = CalendarSession(page)
    session.visible_month = session._read_visible_month()
    return session


def close_modal_if_open(page: Page) -> None:
    if page.locator("#myModal5.modal.in, #myModal5.modal.show").count() == 0:
        return
    page.evaluate(
        """() => {
            const m = window.jQuery ? jQuery('#myModal5') : null;
            if (m && m.length) m.modal('hide');
        }"""
    )
    page.wait_for_selector("#myModal5", state="hidden", timeout=5_000)


def trigger_select_change(page: Page, selector: str) -> None:
    page.evaluate(
        """(sel) => {
            const el = document.querySelector(sel);
            if (!el) throw new Error('Missing select: ' + sel);
            if (window.jQuery) jQuery(el).trigger('change');
            else el.dispatchEvent(new Event('change', { bubbles: true }));
        }""",
        selector,
    )


def select_option(page: Page, selector: str, *, label: str) -> None:
    page.select_option(selector, label=label)
    trigger_select_change(page, selector)


def select_option_by_label_jquery(page: Page, selector: str, label: str) -> None:
    page.evaluate(
        """([sel, name]) => {
            const el = document.querySelector(sel);
            const opt = [...el.options].find(o => o.text.trim() === name);
            if (window.jQuery) jQuery(el).val(opt.value).trigger('change');
            else { el.value = opt.value; el.dispatchEvent(new Event('change', { bubbles: true })); }
        }""",
        [selector, label],
    )


def wait_for_project_option(page: Page, project_name: str, timeout_ms: int = 20_000) -> None:
    page.wait_for_function(
        """([sel, name]) => {
            const el = document.querySelector(sel);
            return el && Array.from(el.options).some(o => o.text.trim() === name);
        }""",
        arg=["#project", project_name],
        timeout=timeout_ms,
    )


def select_job_type(page: Page, label: str) -> None:
    with page.expect_response(
        lambda response: "getProjectDetail" in response.url and response.ok,
        timeout=20_000,
    ):
        select_option_by_label_jquery(page, "#jobtype", label)


def set_remark(page: Page, remark: str) -> None:
    page.evaluate(
        """(text) => {
            const el = document.querySelector('#comment');
            el.value = text;
            if (window.jQuery) jQuery(el).trigger('input').trigger('change');
        }""",
        remark,
    )


def fill_timesheet_form(page: Page, entry: TimesheetEntry) -> None:
    page.wait_for_selector("#jobtype", state="visible")
    select_job_type(page, entry.job_type)
    wait_for_project_option(page, entry.project)
    select_option_by_label_jquery(page, "#project", entry.project)
    select_option(page, "#activities", label=entry.activity)
    page.locator("#effort").fill("")
    page.locator("#effort").fill(entry.effort_hours)
    if entry.event_date:
        page.locator("#start").fill(entry.event_date.strftime("%d/%m/%Y"))
    set_remark(page, entry.remark)


def read_form_snapshot(page: Page) -> dict[str, str]:
    return page.evaluate(
        """() => ({
            project: document.querySelector('#project')?.selectedOptions[0]?.text?.trim() || '',
            remark: document.querySelector('#comment')?.value || '',
        })"""
    )


def save_timesheet(
    page: Page,
    session: CalendarSession | None = None,
    *,
    editing_event_id: int | None = None,
) -> None:
    if editing_event_id is not None:
        ensure_edit_event_id(page, editing_event_id)
    else:
        stored = page.evaluate("""() => parseInt(sessionStorage.getItem('Id') || '0', 10)""")
        if stored != 0:
            raise RuntimeError(
                f"Refusing to save new entry while sessionStorage Id={stored} (use edit path)"
            )

    with page.expect_response(
        lambda response: "SaveTimesheet" in response.url,
        timeout=30_000,
    ) as response_info:
        page.locator("#btnSave").click()
    if not response_info.value.ok:
        raise RuntimeError(f"SaveTimesheet failed: HTTP {response_info.value.status}")
    page.wait_for_selector("#myModal5", state="hidden", timeout=15_000)
    if session:
        session.wait_calendar_idle()
