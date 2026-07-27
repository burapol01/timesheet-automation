ไฟล์ทำงาน — agent และเครื่องมือแก้ได้

- report-data.xlsx       ข้อมูลรายเดือน (openpyxl)
- report-formatted.xlsx  จัด format แล้ว (xlwings, มีโลโก้/ลายเซ็น) → export PDF

ขั้นถัดไป: python tools/excel/export_pdf.py --month July
Output ไป 04-export/pdf/01-employee/
