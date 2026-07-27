# Daily Work Rhythm

**หลักการ:** เอกสาร = source of truth → **คุณบอกเมื่อไหร่ ค่อยอัปเดต** (ไม่มี auto / ไม่เปิดค้าง / ไม่ cron)  
**Email workflow:** ใช้เฉพาะ**สิ้นเดือน** — ไม่ใช่ทุกวัน

> **ไม่กินทรัพยากรเครื่อง** — ไม่มี agent รันค้าง, ไม่มี Outlook poll อัตโนมัติ, ไม่มี scheduled task. ทุกอย่างรอคำสั่งจากคุณเท่านั้น.

---

## แบ่ง 2 โปรเจกต์ (manual trigger)

| เมื่อไหร่ | โปรเจกต์ | ทำอะไร |
|-----------|----------|--------|
| **คุณบอก "เริ่มแก้เอกสาร"** | `Fast-track-revamp-engineering-knowledge` | เขียน/แก้ evidence, analysis ตามงานวันนั้น |
| **คุณบอก "อัปเดท timesheet / ส่งเว็บ"** | `timesheet-automation` | สรุปจากเอกสาร → Excel → intranet |

ไม่มีขั้นตอนไหนรันเองตามเวลา 18:00 — 18:00 เป็นแค่ rhythm ที่คุณเลือกจะสั่งเอง

---

## กลางวัน — งาน + เอกสาร (FTR)

1. ทำงานตาม task (เช่น review code, test, analysis)
2. **บันทึกใน knowledge base ก่อน** — อย่ารอถึง 18:00  
   - Analysis → `11-Data-Analysis/.../analyses/`  
   - Evidence → `00-AI-Brain/evidence/ftrv033/` หรือ `50-Evidence/FTRV033/`  
   - Handoff → `30-Developer-Handoff/FTRV033/`
3. อ้างจากเอกสารเมื่อวาน/วันนี้เพื่อต่อเนื่องงาน

**ตัวอย่างวันนี้ (2026-07-27):**  
Review `D:\Projects\tmlth_windowsservice` →  
[`../Fast-track-revamp-engineering-knowledge/11-Data-Analysis/OneAndDone/legacy-analysis/analyses/autosending-windows-service/source-review-20260727.md`](../Fast-track-revamp-engineering-knowledge/11-Data-Analysis/OneAndDone/legacy-analysis/analyses/autosending-windows-service/source-review-20260727.md)

---

## เมื่อพร้อม — Timesheet + หน้าเว็บ (รอคุณสั่ง)

พูดสั้นๆ ใน chat เช่น:
- *"เริ่มแก้เอกสารวันนี้"* → อัปเดต FTR knowledge จากงานที่ทำ
- *"อัปเดท timesheet / ส่งเว็บ"* → รันเครื่องมือด้านล่าง

### ขั้นตอน (รันครั้งเดียวตอนคุณสั่ง แล้วจบ)

```powershell
cd D:\Projects\timesheet-automation
```

**1. เปิด seed ของเดือน** — แก้ `tools/excel/july_2026_entries.py` (หรือ seed เดือนนั้น)

```python
ReportEntry(
    date(2026, 7, 28),           # วันทำงาน
    ATTENDANCE_WORK,             # "เข้าปฏิบัติงาน"
    "FTRV033",
    "Analysis: short English label",      # col D
    "รายละเอียดจากเอกสารที่เขียนวันนี้",  # col E → intranet Remark
)
```

- ข้อความ col E **ยึดจากเอกสาร FTR** ที่เขียนวันนั้น (ไม่ invent)
- prefix ที่ใช้: `Documentation:` · `Analysis:` · `Development:` · `Testing:` · `Proof:`

**2. เขียนลง Excel**

```powershell
python tools/excel/write_month_data.py --month July
python tools/excel/format_report.py --month July
```

> `write_month_data` → `report-data.xlsx` · **`format_report` → `report-formatted.xlsx`** (ไฟล์ที่เปิดดู/PDF ต้องรันทั้งคู่)

**3. ตรวจวันหยุด (ถ้าต้องการ)**

```powershell
python tools/excel/verify_holidays.py --month July --year 2026
```

**4. Dry-run ส่ง intranet**

```powershell
python tools/web/submit_timesheet.py --sheet July --from-date 2026-07-28 --to-date 2026-07-28 --dry-run
```

**5. ส่งจริง (คุณกด save บนเว็บ — หรือ script save ให้)**

```powershell
# ครั้งแรก (วันที่ยังไม่มี entry)
python tools/web/submit_timesheet.py --sheet July --from-date 2026-07-28 --to-date 2026-07-28

# แก้ remark วันเดิม — ต้องใช้ --force (อัปเดตแถบเดิม ไม่สร้างแถบใหม่)
python tools/web/submit_timesheet.py --sheet July --from-date 2026-07-28 --to-date 2026-07-28 --force
```

> `--force` เปิด modal แก้ไข + ตั้ง `sessionStorage.Id` ก่อน Save — ป้องกัน duplicate bar บนปฏิทิน

---

## สิ้นเดือนเท่านั้น — Email + PDF

```powershell
python tools/excel/format_report.py --month July
python tools/excel/export_pdf.py --month July
python tools/email/draft_workflow.py --month July --step manager --open
# ... รอ reply → fetch_signed_replies → draft approver → accounting
```

ดูรายละเอียด: [`tools/email/README.txt`](tools/email/README.txt)

---

## Flow รวม (on-demand)

```
[คุณทำงาน]     งานจริง
[คุณสั่ง]      "เริ่มแก้เอกสาร" → agent อัปเดต FTR docs
[คุณสั่ง]      "อัปเดท timesheet" → write_month_data → submit intranet
[สิ้นเดือน]    export PDF + email (สั่งเมื่อถึงเวลา)
```

---

## วันนี้ / เมื่อวาน (July 2026)

| วัน | เอกสาร / งานหลัก |
|-----|------------------|
| **2026-07-24** | Agent submit error UX parity (FTR handoff + evidence) |
| **2026-07-27** | AutoSending TA_FASTTRACK review + local install runbook ([doc](../Fast-track-revamp-engineering-knowledge/11-Data-Analysis/OneAndDone/legacy-analysis/analyses/autosending-windows-service/local-install-queue-simulation-runbook-20260727.md)) · timesheet automation |

> เมื่อพร้อม: บอกให้อัปเดต entry 2026-07-27 จากเอกสาร AutoSending review แล้วรัน submit
