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
]
