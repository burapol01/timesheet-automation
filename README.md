# Timesheet Automation

เครื่องมือแยกชั้นชัดเจน — **ไฟล์ต้นฉบับไม่ถูกแตะต้อง**

## โครงสร้าง

```
timesheet-automation/
├── 01-original/                    # ต้นฉบับ (read-only)
│   └── project-report-template.xlsx
├── 02-working/                     # ไฟล์ทำงาน
│   ├── report-data.xlsx            # ข้อมูลรายเดือน → เครื่องมือเว็บอ่าน
│   └── report-formatted.xlsx       # จัด format แล้ว → Save PDF
├── 03-archive/                     # สำรอง report-formatted ก่อนเขียนทับ
├── 04-export/                      # PDF workflow 3 role
│   ├── pdf/01-employee/{year}/{Month}/
│   ├── pdf/02-manager/draft|signed/
│   ├── pdf/03-approver/draft|signed/
│   ├── pdf/04-accounting/draft|sent/
│   └── manifest.json
├── lib/                            # โมดูลร่วม
│   ├── paths.py
│   ├── excel_report.py
│   └── timesheet_browser.py
└── tools/
    ├── web/                        # ส่ง intranet
    │   ├── submit_timesheet.py
    │   └── cleanup_duplicates.py
    └── excel/                      # จัด Excel
        ├── write_month_data.py
        ├── format_report.py
        ├── sync_template_layout.py
        ├── export_pdf.py
        ├── verify_holidays.py
        └── july_2026_entries.py
    └── email/                      # draft / ส่ง PDF
        ├── draft_to_manager.py
        └── README.txt
```

## ขั้นตอนใช้งาน

### 1) เขียนข้อมูล (openpyxl → report-data.xlsx)

```powershell
cd D:\Projects\timesheet-automation
python tools/excel/write_month_data.py --month July
```

### 2) ตรวจวันหยุด / วันทำงาน

```powershell
python tools/excel/verify_holidays.py --month July --year 2026
```

### 3) จัด format ให้เหมือน template + คืนโลโก้/ลายเซ็น (xlwings → report-formatted.xlsx)

```powershell
python tools/excel/format_report.py --month July
```

### 3b) Sync layout ยืนยันเข้า template / ไฟล์ทำงานทุกเดือน

```powershell
python tools/excel/sync_template_layout.py --target all
```

ใช้หลังยืนยัน footer/summary ใน `report-formatted.xlsx` แล้ว — อัปเดต `01-original/` และชีต Jan/Jun–Dec ที่เหลือ

### 4) Export PDF (xlwings → 04-export/)

```powershell
python tools/excel/export_pdf.py --month July
python tools/excel/export_pdf.py --all
```

### 5) Draft email → Manager ลงนาม (Outlook, ไม่ส่งอัตโนมัติ)

```powershell
python tools/email/draft_to_manager.py --month July --dry-run
python tools/email/draft_to_manager.py --month July --open
```

To: `saifon.nam@tokiomarinelife.co.th` (ไม่มี CC) · แนบ PDF จาก `04-export/pdf/01-employee/` · เปิด Outlook > Drafts ตรวจก่อนส่ง

### 6) ส่ง timesheet บน intranet (อ่าน report-data.xlsx เท่านั้น)

```powershell
python tools/web/submit_timesheet.py --sheet July --from-date 2026-07-01 --to-date 2026-07-27
```

- คอลัมน์ **รายละเอียดของงานที่ทำ** → Remark บนเว็บ
- ข้ามวันที่มี entry อยู่แล้ว (วันละ 1 งาน)
- `--dry-run` ดูรายการก่อนส่ง, `--no-save` กรอกแต่ไม่ save

### 7) ลบ entry ซ้ำ (ถ้ามี)

```powershell
python tools/web/cleanup_duplicates.py --dry-run
python tools/web/cleanup_duplicates.py
```

## กฎสำคัญ

| ไฟล์ | เครื่องมือ | หมายเหตุ |
|------|-----------|----------|
| `01-original/*` | **sync_template_layout เท่านั้น** | layout ยืนยันแล้ว; format copy ออกเท่านั้น |
| `02-working/report-data.xlsx` | openpyxl + sync | ข้อมูล + ส่งเว็บ |
| `02-working/report-formatted.xlsx` | xlwings + sync | export PDF |
| `04-export/pdf/01-employee/` | export_pdf | export เริ่มต้น |
| `04-export/pdf/02-manager/signed/` | fetch_signed_replies | รับจาก คุณสายฝน นามกูล |
| `04-export/pdf/03-approver/signed/` | fetch_signed_replies | รับจาก Achara |
| `04-export/pdf/04-accounting/sent/` | (หลังส่ง IT-D) | เก็บหลังส่งแล้ว |

## ติดตั้ง

```powershell
pip install -r requirements.txt
playwright install msedge
```

## โฟลเดอร์เก่า

`Documents/` เป็นไฟล์ช่วงทดลอง — ใช้ `01-original/` และ `02-working/` แทน
