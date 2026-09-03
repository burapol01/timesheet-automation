"""September 2026 seed data for write_month_data.py."""

from __future__ import annotations

from datetime import date

from excel_report import ATTENDANCE_WORK, ReportEntry

SEPTEMBER_2026_ENTRIES: list[ReportEntry] = [
    ReportEntry(
        date(2026, 9, 1),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Analysis + Proof: Sep-1 UAT wave, PA alias V1, combined predev live gates, Windows VPN/queue recovery tooling",
        "Sep-1 queue missions — overnight baseline + morning report; PA alias live TA + combined predev live R1/R2; TA history 23-31; business terminal R2/R3; PA rider required V1 live; UAT PA rider config override; PA52 SSQ ActiveLog + enhance/revamp paired terminal; TA2609015012 independent forensics; UI endpoint patch + touchpad batch closeout; Ivanti VPN open + MFA retry (FLOW_COMPLETE route OK); Windows queue recovery nudges + new Start-IvantiVpnFullFlow tooling and agent rule — read-only, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 2),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Analysis + Proof: Sep-2 PA52 lineage + rider catalog reconcile + SSQ terminal-pass R2 + Evidence Control Tower V1.1 + PA Rider V1 preflight",
        "Sep-2 queue missions — PA52 request-field lineage (FIRST_DIVERGENCE_PROVEN: 5009 10-call vs 5011 single-session; PA43 5058 counterevidence); rider catalog three-way reconcile (PA43/44/52 CMK vs DB vs OLAE); terminal-pass SSQ R2 (63/63 cohort, 22 SEMANTICS_PROVED); Evidence Control Tower V1.1 CORE_COMPLETE; PA52 testdb2 appsnapshot/product/historical oracle + enhance2002 clone; PA Rider V1 isolated preflight PASS (3ea3d697, 18/18); FTR skill pack install; TA case inventory origin; control tower safe shutdown — read-only, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 3),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-3 PA43 OLAE guard submit fix PR #8682 merge dev + UAT TA2609035028 createnew pass + SOAP/rowshape forensics + contact-address A/B",
        "Sep-3 — PA43 paAttach submit-layer fix: wire IPAAttachedRiderRepository fail-closed + null OLAE shape for attached riders; PR #8682 merged dev; UAT TA2609035028 createnew+send SEND_STATUS=S (10 cov S, paAttach COVERAGE/PAYMENT null); morning forensics — SOAP semantic R2, row-shape A/B reject, extra-TA containment, effective-config ON, wslog TA2608245058 vs TA2609035012 compare, createnew XML replay; afternoon TA2609035052 client-create address mapping fail proof + AppInfo JSON export; T6776 DrugFlag ever gap still open; Worker VB ON HOLD — read-only/Owner-authorized UAT write, Activation NO-GO",
    ),
]
