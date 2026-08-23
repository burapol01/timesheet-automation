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
    ReportEntry(
        date(2026, 8, 13),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Analysis + Testing: WS activity root-cause, Tasks 40411/40413 proof, Worker route/testdb2 audits",
        "WS activity ROOTCAUSE_PACKET_PASS + Azure backlog; Tasks 40411/40413 PROOF_READY 14/14; Worker testdb2 SSQ HOLD_ROUTE_MISMATCH; effective route EFFECTIVE_TESTDB2_PROVEN; Testdb2 REVAMP Worker direct investigation; AS400 historical OLAE PROBE_PASS; WS log deploy cancelled zero side-effect — Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 14),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Testing + Proof: Task 40413 post-blackout AS400 terminal, HSGB master, payment MPOS prep",
        "Task 40413: same-TA post-blackout SAME_TA_AS400_TERMINAL_PROVEN (TA2608149017); HSGB master SendOrder 17 on TA2608145001 APP_ID 364; FTR-PAY-MODAL-TOTAL-001 payment modal mismatch; payment/record-mpos v5 TA2608145003 PA Payment_Tr 1/9493 stopped before OLAE; queue agent-first/monitor fixes — Activation NO-GO",
    ),
    # advance — prepared 2026-08-16 (Sunday); work performed on non-work day, row for next work day 2026-08-17
    ReportEntry(
        date(2026, 8, 17),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Development + Testing: Cursor UI Bridge V3-V10 paste proofs, queue poll hardening, Local LLM watcher",
        "Cursor UI Bridge V3-V10 Owner-visible paste/hotkey/composer proofs (V5-V10), superseded READY filter, fetch-blocked poll hardening, persistent paste watcher V8B, Local LLM asset upload watcher dispatch — Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 18),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Analysis: MSI workspace cleanup, evidence recovery, Cursor UI Bridge intake",
        "MSI workspace cleanup — recover stale worktree evidence (daily report, Task 40411/40413 intake, Worker testdb2 plan, Cursor UI Bridge V4 POC) — Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 19),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing + Analysis: Revamp UAT Worker session, 5010 F986 despatch, ID dash parity, OPD ST733",
        "Revamp2024 UAT PM — TA2608185008 SM (ID dash fix), TA2608196006 OPD1 ST733 F, 5010 despatch F986 evidence, Frontend PayMode verify closure, Worker investigation 17-19 Aug — Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 20),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Testing + Development: UAT EOD OPD/TR3N/K10C gaps, Local LLM Windows POC",
        "Revamp2024 UAT 20 Aug EOD — OA_M_WIN_DESCRIPTION wslog parity, OPD three-fields, TR3N pair, K10C Menu not found; Local LLM supervised merge Cursor POC + Git triage on Windows — Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 8, 21),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Testing + Proof: UAT SM cases, LTEWI5 closure, post-send/OPD closures, CRS Azure handoff",
        "UAT 21 Aug — SM 5001/5002/5004/5006/5007, LTEWI5 unlock CLOSED, OA_M_WIN_DESCRIPTION post-send CLOSED, OPD1 6006/5004 CLOSED, CRS gap D153/D154 Azure handoff, Backend dev 5dc98535 v1.20.1, Obsidian overnight relay + Cursor live watcher — Activation NO-GO",
    ),
    # advance — prepared 2026-08-23 (Sunday); weekend work 2026-08-22–23, row for next work day 2026-08-24
    ReportEntry(
        date(2026, 8, 24),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Development + Testing: Queue auto-finalize publish fix, automation smoke proof retry 2, Task 41135 F986 analysis",
        "Windows queue auto-finalize publish fix (remote readback PASS), FTR final automation smoke proof retry 2 PASS, Task 41135 F986/5010 despatch analysis-design (retest 8036001 S) — Activation NO-GO",
    ),
]
