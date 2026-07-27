โฟลเดอร์ส่งมอบ — แยกจาก 02-working (Excel)

ลำดับชั้น PDF (email workflow 3 role):
  pdf/01-employee/{year}/{Month}/           export จาก Excel
  pdf/02-manager/draft|signed/{year}/{Month}/
  pdf/03-approver/draft|signed/{year}/{Month}/
  pdf/04-accounting/draft|sent/{year}/{Month}/

manifest.json                 รายการ export ล่าสุด
workflow_state.json           สถานะรับ PDF จาก inbox
inbox_fetch_report.json       รายงาน fetch ล่าสุด

คำสั่ง:
  python tools/excel/export_pdf.py --month July
  python tools/email/fetch_signed_replies.py --dry-run
  python tools/email/draft_workflow.py --month July --step accounting
