"""Email addresses and display names — แก้ตรงนี้เมื่อองค์กรเปลี่ยน."""

from __future__ import annotations

# ขั้น 1: ส่ง PDF ให้ Manager ลงนาม (draft → คุณสายฝน)
MANAGER_EMAIL = "saifon.nam@tokiomarinelife.co.th"
MANAGER_NAME_TH = "คุณสายฝน นามกูล"
MANAGER_TITLE = "Manager - Fasttrack Development"
MANAGER_DISPLAY = "สายฝน นามกูล"
MANAGER_GREETING = "คุณสายฝน"

# ขั้น 2–3 (ยังไม่ใช้ใน draft แรก)
APPROVER_EMAIL = "achara.cha@tokiomarinelife.co.th"
APPROVER_NAME_TH = "คุณอัจฉรา ชัยภูมิ"
APPROVER_DISPLAY = "อัจฉรา ชัยภูมิ"
ACCOUNTING_TO = "accounting@it-d.biz"
ACCOUNTING_CC = "ar@it-d.biz"

SENDER_DISPLAY = "Burapol Ussawawirulrit"
SENDER_NAME_TH = "บุรพล อัศววิรุฬห์ฤทธิ์"

# โดเมนอีเมลองค์กร (Exchange อาจใช้ tokiomarine หรือ tokiomarinelife)
ORG_EMAIL_DOMAINS = ("tokiomarinelife.co.th", "tokiomarine.co.th")

# alias สำหรับจับ reply — local-part เดียวกัน = คนเดียวกัน
MANAGER_EMAIL_ALIASES = (
    MANAGER_EMAIL,
    "saifon.nam@tokiomarine.co.th",
)
APPROVER_EMAIL_ALIASES = (
    APPROVER_EMAIL,
    "achara.cha@tokiomarine.co.th",
)

ROLE_BY_EMAIL: dict[str, str] = {}
for _addr in MANAGER_EMAIL_ALIASES:
    ROLE_BY_EMAIL[_addr.lower()] = "manager"
for _addr in APPROVER_EMAIL_ALIASES:
    ROLE_BY_EMAIL[_addr.lower()] = "approver"
