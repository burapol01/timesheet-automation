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
    ReportEntry(
        date(2026, 9, 4),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-4 OA_T_CLIENT always-2 self-addr PR #8722 merge + UAT SM/S + SSQ memo/SEND forensics + FAT/MED gap",
        "Sep-4 — OA_T_CLIENT always-2 self address: PR #8722 merged dev; UAT TA2609045033/5038/5040 SM/S (clients=2); TA2609045041 QU/W memo C20+PRE only; morning DrugFlag ever→T6776 inventory; evening SSQ/memo SEND forensics; FAT/MED + 90%≠SEND_LIFE_ASIA=N gap noted — Owner-authorized UAT write, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 7),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-7 salary gate blank Payor income→0 PR #8747 merge + Owner FE TA2609075010 SM/S",
        "Sep-7 — salaryhistory Enhance parity: Backend blank/null Payor ANNUAL_INCOME→0 (PayerClientCommandMaterializer + mapPayer); PR #8747 merged dev 41727cea; prove TA2609075007 salary PASS; Owner FE TA2609075010 / 18035645 SM/S cov 10/10 receipt 109+I2; LOG 9675 = probe pay 106 BANK null not Coverage; residual FE clear TA2609075002 + distinct-payor force-address — read-only/Owner-authorized UAT write, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 8),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-8 FAT/MED DocApp identity compatibility + TA2609085014 C20+MED + Azure 42547 Closed",
        "Sep-8 — memo FAT/MED root cause DocApp.app_number Revamp ID vs legacy TA query mismatch; Backend fix supports both identities; prove TA2609085014 C20+MED; ADP 90% Owner accepted; DrugFlag separated to 42644 Frontend/UAT; Backend 831572c9 release 32cde1d78 1.43.4 — read-only/Owner-authorized UAT write, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 9),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Analysis: Sep-9 PA Rider catalog reconcile (42548) + website/TESTTMLDB3/Tester TA authority + plugin baseline prep",
        "Sep-9 — review Azure 42548 PA Rider terminal failure vs prior PA43 PR #8682 merge; build product catalog from website + TESTTMLDB3 + historical Tester TA; PA43–PA48 web-active, PA53 TEST-only, PA52 production closed per Owner; no new terminal PASS claim without SSQ revalidate; reusable query pack + FTR Work Tracker baseline prepared for Windows DB publish — read-only, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 10),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-10 PA45–48/PA53/MED terminal PASS + SMS 42685 P0 + Windows DB workspace inventory R1/R2",
        "Sep-10 — PA45–PA48 TA2609095003/5005/5006/5007 terminal SM/S 10S/0F cards 42679–42682 closed; PA53 TA2609095025 PASS; MED DocApp 18035799 42671 closed; Health-Q memo dev tests 30/30; P0 42685 SMS null-id investigation; Windows DB workspace inventory (Knowledge dirty paths, worktrees/staging/backup sizing) + R1 Recycle Bin ledger ~10GB + R2 HOT/WARM/COLD cleanup plan — host maintenance separate from business delivery, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 11),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-11 card closures + Qwen Route C R5 + Worker health monitor V1 read-only",
        "Sep-11 — close 42671 MED, 42679–42682 PA45–48, PA53/42548, 42685 SMS after schema owner auto-id restore, 42549 FE submit contract, 39953 writeback scope; Qwen Route C R5 receipt-before-Cursor proof + regression 82/82 bridge; fix Qwen hang (planner path, receipt lock, context limit); Worker health shadow monitor V1 read-only (39963); finalization gate for no-auto-finalize jobs — read-only/Owner-authorized, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 14),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-14 close 42755/42644/PA54/WLP5 + Health-Q memo gap 42756 + Insured P0 tablet FTR-KI-011",
        "Sep-14 — LT3N/HS7N TA2609145025 SM/S close Azure 42755; DrugFlag TA2609145026 close 42644; PA54 terminal proof; WLP5 TA2609145032; discover 42756 natural submit memo gap (0/4 DMQ/HPQ/HTQ/TUQ on clone TA2609145034); tablet FE v1.00.144 Insured P0 fail-closed FTR-KI-011; bug registry B01–B10 — read-only/Owner-authorized UAT, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 15),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-15 G2 probe + Worker deploy WaitStart/file log + send proofs + Phase A health + plugin ICT time",
        "Sep-15 — G2 availability 503 root cause TA2609145003; skip-pay SEND_LIFE_ASIA Enhance Y vs Revamp N parity FTR-KI-010; infra unblock send 18036131; deploy Worker de2a255 WaitStart=150000 + NetworkService ACL file log; TA2609155001 Health-Q 4/4 natural; TA2609155019 post-deploy SM/S 18036169; Phase A worker health probes G2/FILE_LOG/wslog/lane drift; ftr-work-tracker health UI Thai timezone; local worker guard — Owner-authorized UAT deploy read-only monitor, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 16),
        ATTENDANCE_WORK,
        "FTRV033",
        "Development + Proof + Documentation: Sep-16 EKYC bypass API merge dev + skip-pay parity close + Work Tracker 0.8.0 Azure reconcile",
        "Sep-16 — EKYC bypass API merged Backend dev c251612a; TA2609165012/5016/5017/5025/5027 bypass PASS; skip-pay FTR-KI-010 closed TA2609145002 vs Enhance TA2609152003; Work Tracker 0.8.0 sprint validator fix + Azure 18-card reconcile; TA2609165006 address audit 42787; plugin ICT time baseline — Owner-authorized UAT write, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 17),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-17 address submit ADP matrix PA FATCA Recurring/SD/Memo UAT + PA43 Worker dirty-deploy fix FTR-KI-013",
        "Sep-17 — TA2609165006 address submit PASS; ADP matrix FTR-KI-012 to MSI; PA FATCA TA2609175002 probe; Owner UAT PASS Recurring TA2609175015 GB Prime TA2609175023 Solution Design TA2609175019/5020 Memo TA2609175013; PA43 Worker dirty-tree deploy incident TA2609175038 fail to TA2609175043 PASS; deploy preflight rule FTR-KI-013 — Owner-authorized UAT deploy, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 18),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-18 FATCA S5011 Revamp vs Enhance parity probe + datadic design + plugin/UI missions",
        "Sep-18 — FATCA Revamp LFAST S5011 matrix TA2609182005/5010/5036/5039/5045; W8B+DOC gap escalate LFAST; tester playbook; datadic design 20260918; claude evidence badges gantt; csharp quality gap cards; 42546 distinct payer; memo decode reference update — read-only survey/probe, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 21),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-21 Revamp Sprint UAT BOM day-1 handoff + Case 26 Enhance clone R2 + PA 12/21 HOLD",
        "Sep-21 — Sprint Excel 49-case handoff to Tester agent 00020018; PA Case 12/21 BOM clone HOLD QU/W; Case 26 Enhance oracle TA2609212008 PASS; Revamp R1 TA2609215014 SendLifeAsia FTR-KI-008; R2 TA2609215018 SM/S TRP3 AS400 575 vs 604 issue-date drift; local API 95 commits behind dev; datadic DB diagram prep — BOM via local API, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 22),
        ATTENDANCE_WORK,
        "FTRV033",
        "Proof + Documentation: Sep-22 Happy-path BOM closeout + TA2609145009 partial submit pipeline + repro TA2609225048 + log checker proposal",
        "Sep-22 — CL5FAC agent registry 00017506; Happy-path BOM agent 00020022 batch/ledger/closeout PA53 closed sale; post-migrate prod parity buckets 0; GIO 18/36 submit block questionnaire routing; anchor TA2609145009 PA+OLAE QU/N no SEND_LIFE_ASIA row vs TA2609145007 SM/Y; C# submit date AppInfoRevampSubmitStatusPersister; clone repro TA2609225048 writeExecuted Y; Backend dev 1e2fb6e1 Release build; proposal ftr-revamp-submit-pipeline-integrity-probe — read-only/Owner-authorized UAT API, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 23),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Proof + Analysis: Sep-23 Revamp sprint Excel/Tester sync + legacy BUS-49–55 atlas + pilot 2/4/22 readback + Case 1 KYP BOM",
        "Sep-23 — sync Test_Revamp_sprint (3).xlsm snapshot/manifest; fill Actual result 13 cases from BOM readback (0/49 CLEAN PASS); legacy active-path atlas + Azure BUS-49–55 #43757–#43763 inventory; AgentMate pre-WCF reducer vs WCF mapper census; pilot cases 2/4/22 six-gate 0/3 CLEAN PASS EVIDENCE_PARTIAL; Case 1 KYP BOM TA2609235008 KYP_MISSING; Case 20 EK BOM; queue terminal coherence; webservice parity packet readback — read-only/Owner-authorized UAT, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 24),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Proof + Analysis: Sep-24 Revamp sprint Excel (4) snapshot + AS400 ref cases 2/4/20/21/22 + WL21 rider probe",
        "Sep-24 — sync Test_Revamp_sprint (4).xlsm snapshot/manifest testcase-sheet-20260924; AS400 reference pack (TA2608185008 CRS/MEJ, TA2609155013 pass, TA2609235002 APD/EKA, TA2608315066 pass, TA2609235014 WL21 missing / WP-F ST5D3 rider fail readback); sprint case registry cleanPass 2/49; queue fix timesheet daily finalize without MSI review gate — read-only/Owner-authorized UAT, Activation NO-GO",
    ),
    ReportEntry(
        date(2026, 9, 25),
        ATTENDANCE_WORK,
        "FTRV033",
        "Documentation + Proof + Analysis: Sep-25 memo BOM draft batch + Case 9 F986 SSQ RCA + Tester Excel sync",
        "Sep-25 — Invoke-SprintMemoCasesBomDraft cases 9/10/15/19/24 save-draft agent 00020022; TA2609255013 Case 9 ER root cause AS400 F986 Automatic Nums Exhausted (SSQ S2480 Action A) runbook fb342dba; SSQ log fetch 2026-09-25; daily report + queue closeout — read-only/Owner-authorized UAT, Activation NO-GO",
    ),
]
