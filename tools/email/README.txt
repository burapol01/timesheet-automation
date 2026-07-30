Email workflow (Outlook draft — ไม่ส่งอัตโนมัติ)

=== ลำดับชั้น PDF ===
01-employee/              export จาก Excel
02-manager/draft|signed/  draft → คุณสายฝน นามกูล | รับกลับจาก Manager
03-approver/draft|signed/ draft → Achara | รับกลับจาก Achara
04-accounting/draft|sent/ draft → IT-D | หลังส่งแล้ว

ชื่อไฟล์:
  ..._FTR_Timesheet.pdf              (employee)
  ..._FTR_Timesheet_to_manager.pdf    (snapshot draft)
  ..._FTR_Timesheet_manager_signed.pdf
  ..._FTR_Timesheet_to_approver.pdf
  ..._FTR_Timesheet_approver_signed.pdf
  ..._FTR_Timesheet_to_accounting.pdf
  ..._FTR_Timesheet_accounting_sent.pdf

=== คำสั่ง ===
export:
  python tools/excel/export_pdf.py --month July

draft ครบ 3 role:
  python tools/email/draft_workflow.py --month July --all

ดึง PDF จาก inbox:
  python tools/email/fetch_signed_replies.py --month June --year 2026

| ขั้น | To | แนบจาก |
|------|-----|--------|
| manager | สายฝน นามกูล | 01-employee |
| approver | Achara | 02-manager/signed |
| accounting | IT-D | 03-approver/signed |

รายงาน: 04-export/inbox_fetch_report.json
สถานะ:   04-export/workflow_state.json
