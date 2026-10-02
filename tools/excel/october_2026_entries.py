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
    ReportEntry(
        date(2026, 10, 2),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Proof + Development: Oct-2 queue phase1 identity + live rollout proofs; Windows runtime atlas; legacy remain SP forensic; BUS-21/BUS-44 read-only worker proofs; Tester Excel plugin reconcile R1/R2; memo legacy parity docs",
        "Oct-2 — Phase1 payload identity hardening + identity-contract live proof r3; first-live-readonly rollout; agent-queue-plugin main reconcile; windows-runtime-status post-continuity + system-atlas runtime inventory; legacy-remain SP forensic + provenance reconcile; BUS-21 Enhance worker monitoring reconciliation (read-only); BUS-44 deployed worker selector fetch proof; tester-excel-plugin reconcile r1/r2 + R2 evidence contract; tester-excel-sync architecture review; memo workflow legacy parity + SQL evidence; office governor nightshift event log — VERIFY_ONLY, Activation NO-GO",
    ),
]
