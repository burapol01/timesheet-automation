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
    ReportEntry(
        date(2026, 8, 6),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof: Task 39950 production Remain Y/W, worker pickup, SOAP guard local branch และ PARTIAL Owner sign-off",
        "Task 39950: DBA production Remain SP → Y/W+Gate3A PASS, worker pickup GetApplication on OLAE_RV_2024 (shared testdb2), local feature branch BlockAs400Send+life_*→BaseService guard+FetchOnlyForceRefNo, fetch-only log deferred (shared worker by design), OWNER-SIGNOFF-PARTIAL+Knowledge commit/push 519cf3f",
    ),
    ReportEntry(
        date(2026, 8, 7),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Development: Obsidian System Book V0.1–V0.1.2, Azure live recheck และ Backend/Knowledge repo sync",
        "Obsidian System Book V0.1 red-team+prototype, V0.1.1 visual polish, V0.1.2 TOC readability fix; Azure live recheck read-only (39955 Develop/Burapol); Backend dev+Knowledge main repository sync confirmed current — docs/plugin only, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 10),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Testing + Documentation: Task 39955 NonAtomic recovery contract, Windows proof 31/31, PR #8364 merge และ Azure closure",
        "Task 39955 Pure C# NonAtomic recovery/reconciliation contract (27ccfdca): targeted 31/31 PASS on MSI+Windows, Release 0/0; Windows contract proof+Knowledge publish; PR #8364 merged dev, AB#39955 Closed; Dashboard/Work Tracker V0.4.1 closeout — R061 durable DB ownership/runtime still NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 11),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Testing + Proof: Task 40322 frontend status contract + Live DB rollback, Task 39962 Azure close, queue finalize hardening",
        "Task 40322: Frontend status guide publish, real-app rollback Live DB 5/5 PASS (aa7eade1), MSI Pure C# 23/23, daily close sync; Task 39962 Azure Closed (Phase 6C Worker pickup f0d494da); Task 39957 defer/metadata decisions; Windows queue exit-status fix + timesheet daily bootstrap — Activation NO-GO",
    ),
]
