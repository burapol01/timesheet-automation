"""

สร้าง Outlook draft ครบ workflow 3 role



  python tools/email/draft_workflow.py --month July --all

  python tools/email/draft_workflow.py --month July --step manager

  python tools/email/draft_workflow.py --month July --step accounting

"""



from __future__ import annotations



import argparse

import shutil

import sys

from pathlib import Path



ROOT = Path(__file__).resolve().parents[2]

sys.path.insert(0, str(ROOT / "lib"))



from email_config import (  # noqa: E402

    ACCOUNTING_CC,

    ACCOUNTING_TO,

    APPROVER_DISPLAY,

    APPROVER_EMAIL,

    MANAGER_DISPLAY,

    MANAGER_EMAIL,

)

from email_templates import (  # noqa: E402

    body_manager_review,

    body_to_accounting,

    body_to_approver,

    subject_manager_review,

    subject_to_accounting,

    subject_to_approver,

)

from outlook_draft import DraftMail, create_outlook_draft  # noqa: E402

from paths import DEFAULT_REPORT_YEAR, ensure_dirs  # noqa: E402

from pdf_export import (  # noqa: E402

    PDF_STAGE_EMPLOYEE,

    PDF_STAGE_MANAGER_SIGNED,

    PDF_STAGE_APPROVER_SIGNED,

    pdf_for_draft_step,

    pdf_output_path,

    snapshot_for_draft,

)



STEPS = ("manager", "approver", "accounting")



_SIMULATE_FROM = {

    "manager_signed": PDF_STAGE_EMPLOYEE,

    "approver_signed": PDF_STAGE_MANAGER_SIGNED,

}





def resolve_pdf_for_step(

    *,

    step: str,

    month: str,

    year: int,

    override: Path | None,

) -> Path:

    if override is not None:

        path = override.resolve()

        if not path.exists():

            raise FileNotFoundError(f"ไม่พบ PDF: {path}")

        return path



    path = pdf_for_draft_step(step=step, month=month, year=year)

    if not path.exists():

        hints = {

            "manager": f"python tools/excel/export_pdf.py --month {month}",

            "approver": "รอ PDF จาก คุณสายฝน (Manager) หรือ fetch_signed_replies.py",

            "accounting": "รอ PDF จาก Achara หรือ fetch_signed_replies.py",

        }

        raise FileNotFoundError(

            f"ไม่พบ PDF [{step}]: {path}\n{hints.get(step, '')}"

        )

    return path





def simulate_signed(*, stage: str, month: str, year: int) -> Path:

    src_stage = _SIMULATE_FROM[stage]

    src = pdf_output_path(month=month, year=year, stage=src_stage)

    dest = pdf_output_path(month=month, year=year, stage=stage)

    if not src.exists():

        raise FileNotFoundError(f"ไม่พบไฟล์ต้นทางจำลอง: {src}")

    dest.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(src, dest)

    print(f"  [simulate] {src_stage} -> {dest}")

    return dest





def ensure_signed_for_step(

    *,

    step: str,

    month: str,

    year: int,

    simulate: bool,

) -> None:

    if step == "manager":

        return

    if step == "approver":

        stage = PDF_STAGE_MANAGER_SIGNED

    else:

        stage = PDF_STAGE_APPROVER_SIGNED

    path = pdf_output_path(month=month, year=year, stage=stage)

    if path.exists():

        return

    if simulate:

        simulate_signed(stage=stage, month=month, year=year)

        return

    raise FileNotFoundError(

        f"ไม่พบ PDF: {path}\n"

        "ใช้ --simulate-signed หรือรัน fetch_signed_replies.py"

    )





def _format_recipient(email: str, display: str | None = None) -> str:

    if display:

        return f"{display} <{email}>"

    return email





def build_drafts(

    *,

    month: str,

    year: int,

    steps: list[str],

    attachments: dict[str, Path],

) -> list[tuple[str, DraftMail]]:

    subject_mgr = subject_manager_review(month_en=month, year=year)

    drafts: list[tuple[str, DraftMail]] = []



    if "manager" in steps:

        drafts.append(

            (

                "manager",

                DraftMail(

                    to=MANAGER_EMAIL,

                    cc="",

                    subject=subject_mgr,

                    body=body_manager_review(month_en=month, year=year),

                    attachments=[attachments["manager"]],

                    display_to=MANAGER_DISPLAY,

                ),

            )

        )



    if "approver" in steps:

        drafts.append(

            (

                "approver",

                DraftMail(

                    to=APPROVER_EMAIL,

                    cc=_format_recipient(MANAGER_EMAIL, MANAGER_DISPLAY),

                    subject=subject_to_approver(month_en=month, year=year),

                    body=body_to_approver(month_en=month, year=year),

                    attachments=[attachments["approver"]],

                    display_to=APPROVER_DISPLAY,

                ),

            )

        )



    if "accounting" in steps:

        cc_parts = [

            _format_recipient(APPROVER_EMAIL, APPROVER_DISPLAY),

            _format_recipient(MANAGER_EMAIL, MANAGER_DISPLAY),

            ACCOUNTING_CC,

        ]

        drafts.append(

            (

                "accounting",

                DraftMail(

                    to=ACCOUNTING_TO,

                    cc="; ".join(cc_parts),

                    subject=subject_to_accounting(month_en=month, year=year),

                    body=body_to_accounting(month_en=month, year=year),

                    attachments=[attachments["accounting"]],

                ),

            )

        )



    return drafts





def main() -> None:

    parser = argparse.ArgumentParser(

        description="Create Outlook drafts for timesheet email workflow"

    )

    parser.add_argument("--month", required=True, help="Sheet name e.g. July")

    parser.add_argument("--year", type=int, default=DEFAULT_REPORT_YEAR)

    parser.add_argument(

        "--step",

        action="append",

        choices=STEPS,

        dest="steps",

        help="ขั้นที่ต้องการ (default: --all = ครบ 3 ขั้น)",

    )

    parser.add_argument("--all", action="store_true", help="สร้าง draft ครบ 3 ขั้น")

    parser.add_argument("--pdf", type=Path, help="Override PDF สำหรับ step manager")

    parser.add_argument("--signed-pdf", type=Path, help="Override PDF สำหรับ approver/accounting")

    parser.add_argument(

        "--simulate-signed",

        action="store_true",

        help="จำลอง PDF ลงนาม (copy จากขั้นก่อนหน้า)",

    )

    parser.add_argument("--dry-run", action="store_true")

    parser.add_argument("--open", action="store_true", help="เปิด draft สุดท้ายใน Outlook")

    args = parser.parse_args()



    steps = list(STEPS) if args.all or not args.steps else args.steps

    ensure_dirs()



    for step in steps:

        if step in ("approver", "accounting"):

            ensure_signed_for_step(

                step=step,

                month=args.month,

                year=args.year,

                simulate=args.simulate_signed,

            )



    attachments: dict[str, Path] = {}

    for step in steps:

        override = None

        if step == "manager" and args.pdf:

            override = args.pdf

        elif step in ("approver", "accounting") and args.signed_pdf:

            override = args.signed_pdf

        src = resolve_pdf_for_step(

            step=step,

            month=args.month,

            year=args.year,

            override=override,

        )

        if not args.dry_run:

            snap = snapshot_for_draft(step=step, month=args.month, year=args.year, source=src)

            attachments[step] = snap

            print(f"  [draft] {step} -> {snap.name}")

        else:

            attachments[step] = src



    drafts = build_drafts(

        month=args.month,

        year=args.year,

        steps=steps,

        attachments=attachments,

    )



    if args.dry_run:

        print("=== DRY RUN ===")

        for name, mail in drafts:

            print(f"\n--- {name} ---")

            if mail.display_to:

                print(f"To:      {mail.display_to} <{mail.to}>")

            else:

                print(f"To:      {mail.to}")

            print(f"Cc:      {mail.cc or '(none)'}")

            print(f"Subject: {mail.subject}")

            print(f"Attach:  {mail.attachments[0].name}")

            print(f"Path:    {mail.attachments[0]}")

            print("--- Body ---")

            print(mail.body.replace("\r\n", "\n"))

        return



    for i, (name, mail) in enumerate(drafts):

        open_last = args.open and i == len(drafts) - 1

        create_outlook_draft(mail, open_draft=open_last)

        print(f"Draft [{name}] saved")



    print(f"Done: {len(drafts)} draft(s) in Outlook > Drafts")





if __name__ == "__main__":

    main()


