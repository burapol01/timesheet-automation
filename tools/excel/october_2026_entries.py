"""October 2026 seed data for write_month_data.py."""

from __future__ import annotations

from datetime import date

from excel_report import ATTENDANCE_WORK, ReportEntry

OCTOBER_2026_ENTRIES: list[ReportEntry] = [
    ReportEntry(
        date(2026, 10, 1),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Proof + Development: Oct-1 Cursor watcher lock/smoke + Tester Excel pipeline + plugin publish + testcase closures",
        "Oct-1 — Windows queue: live watcher lock mutex + reinstall task; install-smoke MSI review + lock proof r2 handoffs; tester-excel-reconcile MSI review (git blob SHA-256); Sync-TesterExcel pipeline + column audit; registry sync cases 9/10/13 Failed + Owner closes 26/31/33/48 (TA2610015001 EKA/EKP); case29 IDP recheck S5011 + case08/31/43 evidence PNGs; ftr-work-tracker v0.10.10–0.10.13 Tester filters + closed-15 UI; publish ftr-error-community (FTR-KI-021), tools/office plugins; F003 Agent Memo / None Memo docs — VERIFY_ONLY, Activation NO-GO",
    ),
]
