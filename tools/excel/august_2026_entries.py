"""August 2026 seed data for write_month_data.py."""

from __future__ import annotations

from datetime import date

from excel_report import ATTENDANCE_WORK, ReportEntry

AUGUST_2026_ENTRIES: list[ReportEntry] = [
    ReportEntry(
        date(2026, 8, 3),
        ATTENDANCE_WORK,
        "FTRV033",
        "Planning: Gate 3A MSI no-DB feasibility, canonical repo sync, FTR workflow tooling repair และ remote commit audit",
        "Gate 3A MSI no-DB feasibility, canonical repository sync, ftr-session-state tooling repair และ new remote commit audit — เตรียม Windows handoff Lane A/B (runtime OFF)",
    ),
    ReportEntry(
        date(2026, 8, 4),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing + Analysis: Gate 3A Lane B partial proof/closure, testdb2 migration, AS400 success TA2608045001 และ Azure Sprint 32 parity checkpoints",
        "Gate 3A Lane B partial proof/closure, testdb2 migration+config sync, worker AS400 success TA2608045001, Azure Sprint 32 scope/rationale/Kanban checkpoints และ legacy/new-system parity assessment (~57%/65%)",
    ),
    ReportEntry(
        date(2026, 8, 5),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Analysis: Validator release Windows DB proof, Remain SP REVAMP N/W root-cause, _TEST SP tooling, closure docs และ PR #8318",
        "Validator release state V1 Windows DB discovery+proof (54cc404a one-shot), Remain SP REVAMP vs legacy Y/W root-cause, SP_OA_UPDATE_APPLICATION_REMAIN_REVAMP_TEST tooling+config override, red-team closure decision+Knowledge docs, Backend PR #8318 merged dev (Activation NO-GO)",
    ),
]
