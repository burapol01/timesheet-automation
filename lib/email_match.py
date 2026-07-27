"""Match inbound Outlook replies to timesheet PDF workflow."""

from __future__ import annotations

import re
from dataclasses import dataclass

from email_templates import THAI_MONTHS
from report_footer import EMPLOYEE_FILE_TAG

# Project_Report_Burapol_June2026_FTR_Timesheet.pdf
# Project_Report_Burapol_June2026_FTR_Timesheet_updated_saifon_signed.pdf
PDF_NAME_RE = re.compile(
    rf"^Project_Report_{re.escape(EMPLOYEE_FILE_TAG)}_"
    r"(?P<month>[A-Za-z]+)(?P<year>\d{4})_FTR_Timesheet"
    r".*\.pdf$",
    re.IGNORECASE,
)

SUBJECT_PREFIX = "ส่งรายงานการปฏิบัติงานประจำเดือน"

THAI_TO_EN = {v: k for k, v in THAI_MONTHS.items()}


@dataclass(frozen=True)
class ParsedPdfAttachment:
    filename: str
    month_en: str
    year: int


def parse_pdf_filename(filename: str) -> ParsedPdfAttachment | None:
    match = PDF_NAME_RE.match(filename)
    if not match:
        return None
    month = match.group("month")
    if month not in THAI_MONTHS:
        # normalize first letter upper rest lower e.g. june -> June
        month = month[:1].upper() + month[1:].lower()
        if month not in THAI_MONTHS:
            return None
    return ParsedPdfAttachment(
        filename=filename,
        month_en=month,
        year=int(match.group("year")),
    )


def parse_subject_month_year(subject: str) -> tuple[str, int] | None:
    """Parse '...เดือนมิถุนายน 2569' from subject."""
    if SUBJECT_PREFIX not in subject:
        return None
    for th, en in THAI_TO_EN.items():
        pattern = rf"{re.escape(th)}\s+(?P<year>\d{{4}})"
        match = re.search(pattern, subject)
        if match:
            year_be = int(match.group("year"))
            year = year_be - 543
            return en, year
    return None


def normalize_email(address: str) -> str:
    return address.strip().lower()


def email_local_part(address: str) -> str:
    return normalize_email(address).split("@", 1)[0]
