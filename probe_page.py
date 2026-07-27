"""Probe the calendar timesheet page and modal form."""

from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

URL = "http://intranet02/timesheet/Timesheets"
OUT = Path(__file__).parent / "probe_output"


def main() -> None:
    OUT.mkdir(exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=False, slow_mo=150)
        page = browser.new_page(ignore_https_errors=True)
        page.goto(URL, wait_until="networkidle", timeout=120_000)
        page.wait_for_timeout(2000)
        print(f"After goto URL: {page.url}")

        # Ensure we are on calendar page
        if "Timesheet" not in page.locator("h2").first.inner_text():
            page.locator('a[href="/timesheet/Timesheets"]').click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)
            print(f"After sidebar click URL: {page.url}")

        (OUT / "calendar_page.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(OUT / "calendar_page.png"), full_page=True)

        # FullCalendar selectors
        for sel in [".fc-daygrid-day", ".fc-day", "[data-date]", "#calendar", ".fc"]:
            print(f"{sel}: {page.locator(sel).count()}")

        # Click today (2026-07-27)
        today = page.locator('[data-date="2026-07-27"]').first
        if today.count():
            today.click(force=True)
            page.wait_for_timeout(2000)
        else:
            # fallback: click a day number in current month
            day = page.locator(".fc-daygrid-day:not(.fc-day-other) .fc-daygrid-day-number").filter(has_text="27").first
            if day.count():
                day.click()
                page.wait_for_timeout(2000)

        modal = page.locator(".modal.in, .modal.show, .modal[style*='display: block']")
        print(f"Visible modal: {modal.count()}")
        if modal.count():
            (OUT / "modal.html").write_text(modal.first.inner_html(), encoding="utf-8")
            page.screenshot(path=str(OUT / "modal.png"))

        selects = page.eval_on_selector_all(
            "select",
            """els => els.map((el, i) => ({
                index: i,
                id: el.id,
                name: el.name,
                visible: el.offsetParent !== null,
                options: [...el.options].slice(0, 30).map(o => ({value: o.value, text: o.text.trim()}))
            }))""",
        )
        inputs = page.eval_on_selector_all(
            "input, textarea",
            """els => els.map((el, i) => ({
                index: i,
                tag: el.tagName.toLowerCase(),
                id: el.id,
                name: el.name,
                type: el.type || '',
                placeholder: el.placeholder || '',
                visible: el.offsetParent !== null,
                value: el.value || ''
            }))""",
        )
        (OUT / "calendar_selects.json").write_text(
            json.dumps(selects, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        (OUT / "calendar_inputs.json").write_text(
            json.dumps([x for x in inputs if x["visible"]], indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print("Done. Keeping browser open 8s...")
        page.wait_for_timeout(8000)
        browser.close()


if __name__ == "__main__":
    main()
