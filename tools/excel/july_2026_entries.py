"""July 2026 seed data for write_month_data.py."""

from __future__ import annotations

from datetime import date

from excel_report import ATTENDANCE_WORK, ReportEntry

JULY_2026_ENTRIES: list[ReportEntry] = [
    ReportEntry(
        date(2026, 7, 1),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation: สรุป Phase 8C UAT planning และ owner decisions สำหรับ AppInfoRevamp.data same-route compatibility mode (default OFF)",
        "สรุป Phase 8C UAT planning และ owner decisions สำหรับ AppInfoRevamp.data same-route compatibility mode (default OFF, no queue/AS400/tmlth_app_id)",
    ),
    ReportEntry(
        date(2026, 7, 2),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: วิเคราะห์ Owner SQL tracking สำหรับ FTRV033-8G01 บน OLAE_RV_2024 ยืนยัน REF_NO/APP_ID/MASTER_ID linkage จาก Swagger write",
        "วิเคราะห์ Owner SQL tracking สำหรับ FTRV033-8G01 บน OLAE_RV_2024 ยืนยัน REF_NO/APP_ID/MASTER_ID linkage จาก Swagger write",
    ),
    ReportEntry(
        date(2026, 7, 3),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof: Phase 8H rider controlled expansion — owner authorization, committed write proof และ coverage-only completion FTRV033-8H01 (11 coverage rows)",
        "Phase 8H rider controlled expansion: owner authorization, committed write proof และ coverage-only completion FTRV033-8H01 (11 coverage rows, runtime/worker OFF)",
    ),
    ReportEntry(
        date(2026, 7, 6),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing: Phase 8H-2H final read-only verification และ 8H-3B legacy-to-C# OLAE save path parity analysis สำหรับ FTRV033-8H01",
        "Phase 8H-2H final read-only verification และ 8H-3B legacy-to-C# OLAE save path parity analysis สำหรับ FTRV033-8H01 (zero drift, worker/AS400 OFF)",
    ),
    ReportEntry(
        date(2026, 7, 7),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development: Phase 8H-4B/4C reducer2024 pre-transform logging seam — fail-open config-gated capture ก่อน ScriptJS NodeJS exports",
        "Phase 8H-4B/4C reducer2024 pre-transform logging seam: fail-open config-gated capture ก่อน ScriptJS NodeJS exports (default disabled)",
    ),
    ReportEntry(
        date(2026, 7, 8),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing: Phase 8H-4G-2/4G-3 Windows UI/API reducer payload logging validation และ SOURCE_APPINFOREVAMP_DATA real-request validation",
        "Phase 8H-4G-2/4G-3 Windows UI/API reducer payload logging validation และ SOURCE_APPINFOREVAMP_DATA real-request validation (WriteExecuted=false)",
    ),
    ReportEntry(
        date(2026, 7, 9),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: Phase 8H-3B OLAE full table coverage matrix — จัดทำ parity/gap classification ระหว่าง legacy 8-SP chain กับ C# command plan",
        "Phase 8H-3B OLAE full table coverage matrix: parity/gap classification ระหว่าง legacy 8-SP chain กับ C# command plan (QU detail/payment/send held)",
    ),
    ReportEntry(
        date(2026, 7, 10),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: Owner review gap — OLAE UAT synthetic hardcode vs AppInfoRevamp.data partial mapping และ FTRV033_OWNER_UAT_PLAY gate",
        "Owner review gap: OLAE UAT synthetic hardcode vs AppInfoRevamp.data partial mapping และ FTRV033_OWNER_UAT_PLAY gate (real tester-data parity pending)",
    ),
    ReportEntry(
        date(2026, 7, 13),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: Submit parameter SP/table lineage audit — trace caller/auth/parameter contract สำหรับ eapp-revamp-submission path",
        "Submit parameter SP/table lineage audit: trace caller/auth/parameter contract สำหรับ eapp-revamp-submission path (read-only evidence)",
    ),
    ReportEntry(
        date(2026, 7, 14),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development: Application P0 real tester-data mapping — map Weight/Height/Name/ID/DOB/age/sex จาก AppInfoRevamp.data ใน C# OLAE write executor",
        "Application P0 real tester-data mapping: map Weight/Height/Name/ID/DOB/age/sex จาก AppInfoRevamp.data ใน C# OLAE write executor (fail-closed before write)",
    ),
    ReportEntry(
        date(2026, 7, 15),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing: Application P0 projection compare และเตรียม Windows x64 DB read-back verification สำหรับ real tester mapping",
        "Application P0 projection compare และเตรียม Windows x64 DB read-back verification สำหรับ real tester mapping (runtime OFF / activation NO-GO)",
    ),
    ReportEntry(
        date(2026, 7, 16),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: Payment handoff / canonical Unit-of-Work boundary review — reconcile waiver และ runtime repair evidence",
        "Payment handoff / canonical Unit-of-Work boundary review: reconcile waiver และ runtime repair evidence (verify-only, no activation)",
    ),
    ReportEntry(
        date(2026, 7, 17),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing: Payment direct evidence และ juvenile payer reconcile verification บน Windows candidate branches",
        "Payment direct evidence และ juvenile payer reconcile verification บน Windows candidate branches (verify-only, runtime OFF)",
    ),
    ReportEntry(
        date(2026, 7, 20),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development: Warning checkpoints branch — harness repair, canonical runtime fixes และ build/test gate maintenance",
        "Warning checkpoints branch: harness repair, canonical runtime fixes และ build/test gate maintenance (no runtime activation)",
    ),
    ReportEntry(
        date(2026, 7, 21),
        ATTENDANCE_WORK,
        "FTRV033",
        "Testing: Agent guard verify และ AppInfo lookup verification — negative/parity cases สำหรับ submission enrichment path",
        "Agent guard verify และ AppInfo lookup verification: negative/parity cases สำหรับ submission enrichment path",
    ),
    ReportEntry(
        date(2026, 7, 22),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation: อัปเดต FTR roadmap/AI-Brain, harness control docs และ session evidence ตาม sprint FTRV033 status ล่าสุด",
        "อัปเดต FTR roadmap/AI-Brain, harness control docs และ session evidence ตาม sprint FTRV033 status ล่าสุด",
    ),
    ReportEntry(
        date(2026, 7, 23),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development: EApp Step 03b agent validation — vw-useragent verification, field setup both systems และ test-data SQL scripts",
        "EApp Step 03b agent validation: vw-useragent verification, field setup both systems และ test-data SQL scripts (00020001)",
    ),
    ReportEntry(
        date(2026, 7, 24),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development: Agent submit error UX parity — AgentInfoEquivalentGuard pipeline, Swagger/API test evidence และ AMLO/SD gap decision",
        "Agent submit error UX parity: AgentInfoEquivalentGuard pipeline, Swagger/API test evidence และ AMLO/SD gap decision (FE handoff pending)",
    ),
    ReportEntry(
        date(2026, 7, 27),
        ATTENDANCE_WORK,
        "FTRV033",
        "Analysis: AutoSending TA_FASTTRACK — review, fetch-only probe, local install runbook",
        "Review AutoSending (TA_FASTTRACK): BlockAs400Send + probe testtmldb3, ติดตั้ง dev, จำลองคิว/revert, runbook FTR knowledge, timesheet automation",
    ),
]
