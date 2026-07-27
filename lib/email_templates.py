"""Subject / body templates for timesheet report emails."""

from __future__ import annotations

from email_config import MANAGER_GREETING, SENDER_NAME_TH

THAI_MONTHS = {
    "January": "มกราคม",
    "February": "กุมภาพันธ์",
    "March": "มีนาคม",
    "April": "เมษายน",
    "May": "พฤษภาคม",
    "June": "มิถุนายน",
    "July": "กรกฎาคม",
    "August": "สิงหาคม",
    "September": "กันยายน",
    "October": "ตุลาคม",
    "November": "พฤศจิกายน",
    "December": "ธันวาคม",
}


def thai_month(month_en: str) -> str:
    try:
        return THAI_MONTHS[month_en]
    except KeyError as exc:
        raise ValueError(f"Unknown month sheet: {month_en}") from exc


def buddhist_year(year: int) -> int:
    return year + 543


def subject_manager_review(*, month_en: str, year: int) -> str:
    return (
        f"ส่งรายงานการปฏิบัติงานประจำเดือน{thai_month(month_en)} {buddhist_year(year)}"
    )


def body_manager_review(*, month_en: str, year: int) -> str:
    month_th = thai_month(month_en)
    year_be = buddhist_year(year)
    return (
        f"เรียน {MANAGER_GREETING}\r\n"
        f"ผมขอนำส่งรายงานการปฏิบัติงานประจำเดือน{month_th} {year_be} "
        f"ตามไฟล์แนบครับ\r\n"
        f"รบกวนตรวจสอบและลงนาม หากต้องแก้ไขหรือเพิ่มเติมส่วนใด แจ้งผมได้เลยครับ\r\n"
        f"ขอบคุณครับ\r\n"
        f"{SENDER_NAME_TH}"
    )


# ขั้น 2 (รอ implement): ส่งให้คุณอัจฉรา
def subject_to_approver(*, month_en: str, year: int) -> str:
    return subject_manager_review(month_en=month_en, year=year)


def body_to_approver(*, month_en: str, year: int) -> str:
    month_th = thai_month(month_en)
    year_be = buddhist_year(year)
    return (
        f"เรียน คุณอัจฉรา\r\n"
        f"ผมขอนำส่งรายงานการปฏิบัติงานประจำเดือน{month_th} {year_be} ตามไฟล์แนบครับ\r\n"
        f"รบกวนคุณอัจฉราตรวจสอบเอกสาร หากต้องแก้ไขหรือเพิ่มเติมส่วนใด แจ้งผมได้เลยครับ\r\n"
        f"ขอบคุณครับ\r\n"
        f"{SENDER_NAME_TH}"
    )


# ขั้น 3: ส่งทีมบัญชี IT-D (ไฟล์ลงนามครบ)
def subject_to_accounting(*, month_en: str, year: int) -> str:
    return subject_manager_review(month_en=month_en, year=year)


def body_to_accounting(*, month_en: str, year: int) -> str:
    month_th = thai_month(month_en)
    year_be = buddhist_year(year)
    return (
        f"เรียน ทีมบัญชี IT-D\r\n"
        f"ผมขอนำส่งรายงานการปฏิบัติงานประจำเดือน{month_th} {year_be} "
        f"ที่ลงนามเรียบร้อยแล้ว ตามไฟล์แนบครับ\r\n"
        f"รบกวนตรวจสอบเอกสาร หากต้องแก้ไขหรือเพิ่มเติมข้อมูลส่วนใด แจ้งผมได้เลยครับ\r\n"
        f"ขอบคุณครับ\r\n"
        f"{SENDER_NAME_TH}"
    )
