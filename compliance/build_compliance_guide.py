"""
Solecare Glasgow — Main Compliance Guide (overview document).
Run with: python3 build_compliance_guide.py
Output: ./Compliance_Guide.docx
"""

import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from build_word_forms import (
    a4_doc, add_doc_title, add_section_heading, add_teal_rule,
    body_para, note_para, notice_box, footer_line, checkbox_para,
    form_table, signature_table, shade_cell, cell_border_all,
    set_cell_margins, cell_no_border, hex_str,
    TEAL, DARK, BG, MUTED, TEXT_C, WHITE, RED_LIGHT, RED_TEXT
)
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT

OUT = os.path.dirname(__file__)

def chapter_heading(doc, number, title):
    """Large teal chapter heading with number badge."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    num_cell = tbl.rows[0].cells[0]
    shade_cell(num_cell, hex_str(TEAL))
    cell_no_border(num_cell)
    set_cell_margins(num_cell, top=100, bottom=100, left=120, right=120)
    p = num_cell.paragraphs[0]
    r = p.add_run(str(number))
    r.font.name = 'Calibri Light'
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Force narrow column
    num_cell.width = Cm(1.2)

    title_cell = tbl.rows[0].cells[1]
    shade_cell(title_cell, hex_str(DARK))
    cell_no_border(title_cell)
    set_cell_margins(title_cell, top=100, bottom=100, left=180, right=180)
    p2 = title_cell.paragraphs[0]
    r2 = p2.add_run(title)
    r2.font.name = 'Calibri Light'
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = WHITE
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trH = OxmlElement('w:trHeight')
    trH.set(qn('w:val'), '560')
    trH.set(qn('w:hRule'), 'atLeast')
    trPr.append(trH)

    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(6)

def two_col_table(doc, rows_data, col1_label='', col2_label=''):
    """Simple two-column info table."""
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = 'Table Grid'
    tbl.columns[0].width = Cm(5.5)
    tbl.columns[1].width = Cm(11)
    if col1_label:
        hrow = tbl.add_row()
        for i, h in enumerate([col1_label, col2_label]):
            c = hrow.cells[i]
            shade_cell(c, hex_str(TEAL))
            cell_border_all(c)
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            p = c.paragraphs[0]
            r = p.add_run(h)
            r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    for i, (label, val) in enumerate(rows_data):
        row = tbl.add_row()
        lc = row.cells[0]
        shade_cell(lc, hex_str(BG) if i % 2 == 0 else 'F8FBFB')
        cell_border_all(lc)
        set_cell_margins(lc, top=60, bottom=60, left=100, right=100)
        p = lc.paragraphs[0]
        r = p.add_run(label)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = DARK
        vc = row.cells[1]
        shade_cell(vc, 'FFFFFF')
        cell_border_all(vc)
        set_cell_margins(vc, top=60, bottom=60, left=100, right=100)
        p2 = vc.paragraphs[0]
        r2 = p2.add_run(val)
        r2.font.name = 'Calibri'; r2.font.size = Pt(9); r2.font.color.rgb = TEXT_C
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '320')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
    doc.add_paragraph()

def bullet(doc, text, indent=0, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if bold_prefix:
        r1 = p.add_run(bold_prefix + ' ')
        r1.font.name = 'Calibri'; r1.font.size = Pt(10); r1.font.bold = True; r1.font.color.rgb = DARK
    run = p.add_run(text)
    run.font.name = 'Calibri'; run.font.size = Pt(10); run.font.color.rgb = TEXT_C

def page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(docx_break_type(run))

def docx_break_type(run):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)
    return None  # we already appended it

def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)

def file_card(doc, filename, filetype, description):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    badge_cell = tbl.rows[0].cells[0]
    badge_bg = hex_str(TEAL) if 'xlsx' in filetype.lower() else hex_str(DARK)
    shade_cell(badge_cell, badge_bg)
    cell_no_border(badge_cell)
    set_cell_margins(badge_cell, top=80, bottom=80, left=120, right=120)
    p = badge_cell.paragraphs[0]
    r = p.add_run(filetype)
    r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    badge_cell.width = Cm(1.5)

    desc_cell = tbl.rows[0].cells[1]
    shade_cell(desc_cell, hex_str(BG))
    cell_border_all(desc_cell)
    set_cell_margins(desc_cell, top=80, bottom=80, left=140, right=120)
    p2 = desc_cell.paragraphs[0]
    r2 = p2.add_run(filename + '  ')
    r2.font.name = 'Calibri'; r2.font.size = Pt(10); r2.font.bold = True; r2.font.color.rgb = DARK
    r3 = p2.add_run(description)
    r3.font.name = 'Calibri'; r3.font.size = Pt(9); r3.font.color.rgb = MUTED

    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trH = OxmlElement('w:trHeight')
    trH.set(qn('w:val'), '400')
    trH.set(qn('w:hRule'), 'atLeast')
    trPr.append(trH)

    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(4)


# ════════════════════════════════════════════════════════════════════════════
# BUILD THE GUIDE
# ════════════════════════════════════════════════════════════════════════════
def build_guide():
    doc = a4_doc()

    # ── Cover ────────────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cover = tbl.rows[0].cells[0]
    shade_cell(cover, hex_str(DARK))
    cell_no_border(cover)
    set_cell_margins(cover, top=600, bottom=600, left=300, right=300)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trH = OxmlElement('w:trHeight')
    trH.set(qn('w:val'), '4200')
    trH.set(qn('w:hRule'), 'atLeast')
    trPr.append(trH)

    for line, size, bold in [
        ("Solecare Glasgow", 28, True),
        ("", 8, False),
        ("Private Mobile Podiatry Practice", 14, False),
        ("Compliance and Administration Pack", 14, True),
        ("", 8, False),
        ("September 2026", 10, False),
    ]:
        p = cover.add_paragraph(line)
        for run in p.runs:
            run.font.name = 'Calibri Light'
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = WHITE
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    sub_line = cover.add_paragraph(
        "Prepared for an HCPC-registered podiatrist setting up a sole-trader mobile practice in Scotland."
    )
    for run in sub_line.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0xCC, 0xE5, 0xE6)
    sub_line.paragraph_format.space_before = Pt(8)

    doc.add_paragraph()
    notice_box(doc,
        "This pack contains the information, templates, and logs you need to run a compliant private mobile podiatry practice. "
        "It does not replace professional judgement or legal advice. "
        "Review the annual calendar section and update contact details in the Privacy Notice and Lone Worker Emergency document before seeing your first patient.")

    add_page_break(doc)

    # ── What's in this pack ──────────────────────────────────────────────────
    chapter_heading(doc, "A", "What's in this Pack")

    body_para(doc,
        "This pack was built specifically for a small mobile podiatry practice using single-use instruments and working from a laptop and iPad. "
        "All files are in Microsoft Office formats. The consent forms and policy documents are designed to be completed in Word on a laptop, "
        "or on an iPad using Microsoft Word (included in your Microsoft 365 subscription). "
        "Clients can sign forms on the iPad using a finger or Apple Pencil in the Word app. "
        "All forms also print cleanly if a patient prefers a paper copy.")

    add_section_heading(doc, "Consent Forms and Patient Documents  (compliance/templates/)")
    add_teal_rule(doc)
    forms = [
        ("Consent_Routine.docx", "DOCX", "Patient registration and general treatment consent. One per patient at first appointment. Update when clinical status changes significantly."),
        ("Consent_NailSurgery.docx", "DOCX", "Nail surgery procedure-specific consent. Completed at a separate pre-operative appointment. One per surgical episode."),
        ("Consent_Photography.docx", "DOCX", "Clinical photography consent. One per photography session — treatment consent does not cover photography."),
        ("Privacy_Notice.docx", "DOCX", "UK GDPR-compliant patient privacy notice. Fill in your HCPC number, ICO number, and contact details before distributing. Give to every patient at first contact."),
        ("Complaints_Procedure.docx", "DOCX", "One-page patient-facing complaints procedure plus a practitioner log section. Post on your website if you have one."),
    ]
    for fname, ftype, desc in forms:
        file_card(doc, fname, ftype, desc)

    add_section_heading(doc, "Policy and Protocol Documents  (compliance/templates/)")
    add_teal_rule(doc)
    policies = [
        ("Infection_Control_Policy.docx", "DOCX", "Full written IPC policy including a domiciliary-specific section. Sign annually. Required by HCPC Standards 6.1 and 6.2."),
        ("Risk_Assessment.docx", "DOCX", "H&S risk assessment covering general hazards, domiciliary hazards, vehicle and waste transport, and lone worker risks. COSHH section included."),
        ("Sharps_Protocol.docx", "DOCX", "Single-use sharps management procedure and injury response. Includes immediate post-injury record form."),
        ("Lone_Worker_Emergency.docx", "DOCX", "Emergency contact action plan and escalation procedure. Print and give a copy to your named emergency contact. Both parties sign."),
    ]
    for fname, ftype, desc in policies:
        file_card(doc, fname, ftype, desc)

    add_section_heading(doc, "Excel Working Logs  (compliance/logs/)")
    add_teal_rule(doc)
    logs = [
        ("Financial_Records.xlsx", "XLSX", "Income, Expenses, and Mileage tabs. Mileage rate auto-calculates 45p/25p threshold. Annual Summary tab for Self Assessment prep."),
        ("Clinical_Waste_Log.xlsx", "XLSX", "Transport Log, Consignment Notes log, and Carrier Register. Supports EA duty of care requirement. Retain consignment notes 3 years."),
        ("Daily_Movement_Log.xlsx", "XLSX", "Lone worker daily schedule. Share with emergency contact before first visit each day. Prints on one landscape page."),
        ("IPC_Visit_Log.xlsx", "XLSX", "Per-visit infection control checklist and sharps container assembly/disposal log."),
        ("Medicines_Log.xlsx", "XLSX", "Local anaesthetic administration record, temperature log, and stock register."),
        ("Incident_Log.xlsx", "XLSX", "All incidents and near-misses with RIDDOR-reportable flag. Detailed Sharps Injury sheet included."),
        ("CPD_Log.xlsx", "XLSX", "CPD record for HCPC audit readiness. Auto-totals hours. Includes HCPC audit guidance note."),
    ]
    for fname, ftype, desc in logs:
        file_card(doc, fname, ftype, desc)

    add_page_break(doc)

    # ── NHS to Private: what changes ─────────────────────────────────────────
    chapter_heading(doc, "B", "What Changes Moving from NHS to Private Practice")

    body_para(doc,
        "Most of what you know from 40 years in the NHS applies directly. Clinical standards, patient care, and professional obligations are the same. "
        "The difference is that the NHS provided the compliance infrastructure institutionally. "
        "In private practice, you provide it yourself — scaled to a one-person mobile practice, which is far simpler than Trust-level systems.")

    two_col_table(doc, [
        ("NHS Resolution / Crown indemnity", "Your own professional indemnity policy — buy before your first patient"),
        ("Trust ICO registration", "Individual ICO registration — £52/year, register at ico.org.uk"),
        ("Trust Caldicott Guardian and DPO", "You are the data controller. No DPO required at sole-trader scale."),
        ("IPCT and trust IPC policy", "Your own written IPC policy — Infection_Control_Policy.docx in this pack"),
        ("Estates clinical waste disposal", "Your own contract with a licensed EA-registered carrier"),
        ("Trust PGDs for local anaesthetics", "Not needed — your HCPC POM-A annotation is the statutory authority under Schedule 17 HMR 2012"),
        ("CQC compliance framework", "No CQC registration required for routine sole-trader podiatry"),
        ("NHS pension automatic NI contributions", "Self Assessment — Class 4 and Class 2 NI calculated through your tax return"),
        ("Payroll — employer handles tax", "You pay your own income tax via Self Assessment by 31 January each year"),
        ("PALS complaints route", "Direct response to patient — HCPC is the formal escalation route"),
        ("Occupational health for BBV post-exposure", "Your own documented protocol — Sharps_Protocol.docx in this pack"),
    ], col1_label="NHS provided", col2_label="Your responsibility in private practice")

    add_page_break(doc)

    # ── Before you see your first patient ────────────────────────────────────
    chapter_heading(doc, "C", "Before You See Your First Patient — Checklist")

    notice_box(doc, "Complete every item on this list before your first clinical appointment.")

    checklist = [
        ("Confirm POM-A annotation is visible on your public HCPC record at hcpc-uk.org.",
         "Section 1 of this guide. Without POM-A confirmed, do not administer local anaesthetics."),
        ("Buy professional indemnity and public liability insurance.",
         "Everywhen, Balens, Incision Indemnity, or RCPod membership (James Hallam). Section 4."),
        ("Register as a sole trader for Self Assessment with HMRC.",
         "gov.uk/register-for-self-assessment. Register by 5 October after your first earning tax year. Section 3."),
        ("Register with the ICO (£52/year).",
         "ico.org.uk/register. Required because you process patient health records electronically. Section 5."),
        ("Upgrade car insurance to Class 3 business use. Disclose clinical waste transport.",
         "Get written confirmation from your insurer. Section 9."),
        ("Register as a Lower Tier waste carrier with the Environment Agency (free).",
         "wastecarriersregistration.service.gov.uk. Covers transporting your own clinical waste. Section 9."),
        ("Complete ADR 1.3 general awareness training (~£20 online).",
         "Legal requirement for anyone transporting clinical waste (UN 3291), regardless of quantity. Section 9."),
        ("Fill in and sign the Infection Control Policy.",
         "Infection_Control_Policy.docx — add your name and date and sign. Section 8."),
        ("Complete and test your lone worker check-in arrangement.",
         "Give a copy of Lone_Worker_Emergency.docx to your named emergency contact. Both sign it. Section 9."),
        ("Fill in the brackets in Privacy_Notice.docx with your details.",
         "HCPC number, ICO number, contact email. Give to every patient at first contact."),
        ("Set up a Microsoft 365 Business Basic subscription.",
         "~£60/year. Gives you a business email address and OneDrive for encrypted record storage. Section 5."),
        ("Arrange a clinical waste collection contract with an EA-registered carrier.",
         "Verify EA registration at environment.data.gov.uk/public-register before use."),
    ]
    for item, detail in checklist:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(2)
        r_box = p.add_run('☐  ')
        r_box.font.name = 'Calibri'; r_box.font.size = Pt(12); r_box.font.color.rgb = TEAL
        r_item = p.add_run(item + '  ')
        r_item.font.name = 'Calibri'; r_item.font.size = Pt(10); r_item.font.bold = True; r_item.font.color.rgb = DARK
        r_detail = p.add_run(detail)
        r_detail.font.name = 'Calibri'; r_detail.font.size = Pt(9); r_detail.font.italic = True; r_detail.font.color.rgb = MUTED

    add_page_break(doc)

    # ── Compliance reference sections ─────────────────────────────────────────
    chapter_heading(doc, "1", "HCPC Registration and POM-A")

    body_para(doc,
        "Your HCPC registration and professional obligations are the same in private practice as they were in the NHS. "
        "The one registration-specific thing that matters differently for private practice:")
    notice_box(doc,
        "Check that your POM-A annotation appears on your public HCPC record at hcpc-uk.org. "
        "This is your legal authority to administer local anaesthetics under Schedule 17 of the Human Medicines Regulations 2012. "
        "In private sole-trader practice there is no organisation to underwrite a PGD — the Schedule 17 statutory exemption covers it directly. "
        "If POM-A is absent, contact HCPC to add it before administering any local anaesthetic.")
    body_para(doc,
        "If your registration has lapsed, the readmission process scales with time off the register: "
        "no additional updating under 2 years; 30 days (210 hours) including supervised practice for 2 to 5 years; "
        "60 days for over 5 years. Cost: £431.69 readmission plus two-year registration fee (2025 rates).")
    body_para(doc,
        "CPD in private practice follows the same HCPC qualitative framework — no minimum hours, but keep the log from day one. "
        "Use CPD_Log.xlsx.")
    bold_body = lambda t, b: body_para(doc, f"{t}: {b}")
    body_para(doc,
        "RCPod membership: not a legal requirement, but the £480/year full membership bundles £15m professional indemnity via James Hallam, "
        "access to Clinical Standards 5 and 9 (Infection Control and Domiciliary Care — both directly relevant to your practice), "
        "consent form templates, and CPD resources. Compare against standalone insurance quotes.")

    add_page_break(doc)
    chapter_heading(doc, "2", "Business Setup and Tax")

    notice_box(doc, "This section covers territory that is genuinely new — you have never needed to do any of this as an NHS employee.")

    add_section_heading(doc, "Sole trader structure")
    add_teal_rule(doc)
    body_para(doc,
        "Sole trader is the right structure. A limited company adds Companies House filings, statutory accounts, "
        "corporation tax returns, and director obligations — none of it justified for a small supplementary-income practice.")

    add_section_heading(doc, "Registering with HMRC")
    add_teal_rule(doc)
    body_para(doc,
        "Go to gov.uk/register-for-self-assessment and use the 'Register if you're self-employed' route. "
        "You need your National Insurance number. HMRC issues a Unique Taxpayer Reference (UTR) by post in about 10 working days. "
        "Register by 5 October following the first tax year in which your self-employment income exceeds £1,000.")

    add_section_heading(doc, "Self Assessment")
    add_teal_rule(doc)
    body_para(doc,
        "File a tax return online by 31 January each year covering the previous tax year (6 April to 5 April). "
        "It asks for your total self-employment income, your allowable expenses, and your profit. HMRC calculates the tax. "
        "From the second year, you also pay advance 'payments on account' towards the following year's bill: "
        "half by 31 January, half by 31 July.")
    notice_box(doc,
        "NHS pension and self-employment profit are combined to calculate total taxable income. "
        "If your NHS pension already uses your personal allowance (£12,570 in 2026-27), "
        "self-employment profit is taxable from the first pound. "
        "This is straightforward — it is simply how income tax works with multiple income sources.")

    add_section_heading(doc, "VAT")
    add_teal_rule(doc)
    body_para(doc,
        "Podiatry is VAT-exempt under Schedule 9, Group 7 of the VAT Act 1994. "
        "Exempt income does not count toward the £90,000 registration threshold. "
        "You will not register for VAT and cannot reclaim VAT on purchases.")

    add_section_heading(doc, "Allowable expenses")
    add_teal_rule(doc)
    two_col_table(doc, [
        ("HCPC and RCPod fees", "Fully allowable"),
        ("Professional indemnity and public liability insurance", "Fully allowable"),
        ("CPD costs for maintaining existing competence", "Fully allowable"),
        ("Single-use instruments and consumables", "Fully allowable"),
        ("PPE — gloves, aprons, masks, drapes", "Fully allowable"),
        ("Clinical waste contract", "Fully allowable"),
        ("Business email (Microsoft 365)", "Fully allowable"),
        ("Insulated medicines bag, lockable vehicle storage", "Fully allowable"),
        ("Lone worker app subscription", "Fully allowable"),
        ("Mileage", "45p/mile for first 10,000 business miles, then 25p/mile (HMRC approved rate, no receipts needed)"),
        ("Home office proportion", "Calculate % of home used for admin — allowable proportion of utilities, broadband, heating"),
    ], col1_label="Expense", col2_label="Notes")

    add_section_heading(doc, "Records (Financial_Records.xlsx)")
    add_teal_rule(doc)
    body_para(doc,
        "Keep every income entry (invoice issued, payment received), every expense with receipt, "
        "and a contemporaneous mileage log. Estimates made after the fact are not accepted by HMRC. "
        "Retain everything for 5 years after the 31 January Self Assessment filing deadline for that year.")

    add_page_break(doc)
    chapter_heading(doc, "3", "Insurance")
    body_para(doc,
        "Your NHS indemnity ended on your last working day. Buy insurance before you see any patient.")
    two_col_table(doc, [
        ("Professional indemnity", "Statutory condition of HCPC registration (SI 2014/1887). Providers: Everywhen (~£98/year), Balens, Incision Indemnity, or RCPod membership (James Hallam £15m). MDDUS covers only doctors and dentists."),
        ("Public liability", "Not legally mandated but expected for home visits. Typically bundled with PI."),
        ("Employers liability", "Not required — no employees."),
        ("Nail surgery at home visits", "Check policy wording specifically. Some PI insurers restrict surgery to clinical premises."),
        ("Home insurance", "Notify home insurer if clinical equipment or medicines are stored at home."),
        ("Car insurance", "See Chapter 7 (Mobile Practice)."),
    ])

    add_page_break(doc)
    chapter_heading(doc, "4", "Data Protection and Patient Records")
    add_section_heading(doc, "ICO Registration")
    add_teal_rule(doc)
    body_para(doc,
        "Any sole trader processing patient records electronically must register individually with the ICO — "
        "this does not carry over from your NHS employer. Register at ico.org.uk/register. Annual fee: £52 (Tier 1), £47 by direct debit.")
    add_section_heading(doc, "Patient record retention periods")
    add_teal_rule(doc)
    two_col_table(doc, [
        ("Adults", "8 years from last contact"),
        ("Children", "Until the patient's 25th birthday, or 26th if aged 17 at end of treatment"),
        ("Deceased", "8 years from date of death"),
    ])
    add_section_heading(doc, "Records system")
    add_teal_rule(doc)
    body_para(doc,
        "Paper records are legally acceptable. HCPC does not mandate electronic systems. "
        "Your NHS record habits translate directly and are fully compliant. "
        "Digital records require individual password-protected login, encrypted storage, and a documented backup schedule. "
        "OneDrive (included in Microsoft 365 Business Basic) is encrypted at rest and meets UK GDPR requirements — "
        "Microsoft holds a standard Data Processing Agreement.")
    notice_box(doc,
        "Business email is required for patient communications — consumer email accounts (Gmail, Hotmail) are not acceptable for patient data. "
        "Microsoft 365 Business Basic (~£60/year) provides a business email address, OneDrive, and the full Office apps on up to 5 devices including your iPad.")

    add_page_break(doc)
    chapter_heading(doc, "5", "Consent")
    body_para(doc,
        "The Montgomery standard and your consent obligations are identical in private practice. "
        "The difference is that you design your own forms. This pack provides three Word templates:")
    bullet(doc, "Consent_Routine.docx — patient registration and general treatment consent (one per patient)")
    bullet(doc, "Consent_NailSurgery.docx — procedure-specific consent (one per surgical episode, completed at a separate pre-op appointment)")
    bullet(doc, "Consent_Photography.docx — clinical photography consent (one per image session)")
    body_para(doc,
        "Consent records are retained on the same schedule as clinical records (8 years for adults).")

    add_page_break(doc)
    chapter_heading(doc, "6", "Infection Control")
    body_para(doc,
        "Reference standards: NHS NIPCM v2.10 (May 2024) — publicly available from NHS England; "
        "RCPod Clinical Standards 5 and 9 — available from your member area at rcpod.org.uk.")
    body_para(doc,
        "Because you use single-use instruments exclusively, there is no decontamination burden — "
        "no autoclave, no cycle logs, no HTM 01-05 compliance. "
        "The compliance burden shifts to: (a) sourcing sterile single-use instruments from a licensed medical device supplier; "
        "(b) correct clinical waste disposal of used instruments; (c) your documented IPC policy.")
    notice_box(doc,
        "The written IPC policy must include a domiciliary-specific section. "
        "A generic clinic-based policy is not sufficient for mobile practice. "
        "RCPod Clinical Standard 9 is the specific benchmark. "
        "Infection_Control_Policy.docx in this pack covers both.")
    add_section_heading(doc, "Nail surgery in a domestic setting")
    add_teal_rule(doc)
    body_para(doc,
        "Nail surgery carries higher infection risk than routine podiatry. Before each nail surgery home visit, "
        "satisfy yourself (and document) that the environment is suitable: adequate lighting, no contamination risk, "
        "ability to maintain a sterile field, patient consent to home setting for surgery. "
        "Check your PI insurance policy explicitly covers nail surgery in domestic settings.")

    add_page_break(doc)
    chapter_heading(doc, "7", "Mobile Practice: Vehicle, Lone Working, and Waste Transport")
    add_section_heading(doc, "Car insurance")
    add_teal_rule(doc)
    body_para(doc,
        "Class 3 business use is the appropriate level for visiting multiple patients per day. "
        "Standard social/domestic/pleasure cover is void for driving between multiple patient addresses for work. "
        "Additionally disclose to your insurer that you transport small quantities of clinical waste. Get written confirmation.")
    add_section_heading(doc, "Transporting clinical waste")
    add_teal_rule(doc)
    two_col_table(doc, [
        ("Lower Tier EA waste carrier registration", "Free, permanent, no renewal. wastecarriersregistration.service.gov.uk. Register before first journey."),
        ("ADR 1.3 general awareness training", "Legal requirement regardless of quantity. Short online course ~£20. Keep certificate. ADR Education, Human Focus."),
        ("Outer transport container", "Sealed sharps bins and clinical waste bags must go inside a rigid, lockable, leak-proof box. Loose bags in the boot are prohibited under ADR."),
        ("Fire extinguisher", "2 kg dry powder in the vehicle when carrying clinical waste. The only additional vehicle equipment required below ADR threshold."),
        ("What you do NOT need", "Full ADR Driver Certificate (5-day course). Orange plates. ADR transport documents. The 333 kg UN 3291 threshold is far beyond what a single practitioner carries."),
        ("Consignment notes", "One Hazardous Waste Consignment Note per carrier collection of sharps and infectious waste. Retain 3 years. Log in Clinical_Waste_Log.xlsx."),
    ])
    add_section_heading(doc, "Medicines in the vehicle")
    add_teal_rule(doc)
    body_para(doc,
        "Local anaesthetics must be stored below 25°C (SmPC requirement). "
        "UK car boots regularly exceed 37°C in summer. An insulated medicines bag is required in warm weather — "
        "this is a patient safety issue, not just a paperwork one. "
        "Local anaesthetics are not controlled drugs; a lockable insulated bag in the boot covers the POM security requirement.")
    add_section_heading(doc, "Lone working")
    add_teal_rule(doc)
    body_para(doc,
        "There is no specific UK lone worker statute. Obligations derive from HSWA 1974 and MHSWR 1999. "
        "Add a lone worker section to your risk assessment (Risk_Assessment.docx) rather than a separate document.")
    bullet(doc, "Daily: share movement log with named emergency contact before first visit (Daily_Movement_Log.xlsx)")
    bullet(doc, "Check in after each patient. Send end-of-day confirmation.")
    bullet(doc, "Emergency contact has Lone_Worker_Emergency.docx and knows the escalation steps.")
    bullet(doc, "Apps: Lookout Call (~£2/month, telephony-based, used by NHS community nursing) or MyTeamSafe (from £2.50/month, automated SMS escalation).")
    bullet(doc, "Carry portable footrest and folding stool. Sustained awkward posture across multiple house calls is a real occupational health risk.")

    add_page_break(doc)
    chapter_heading(doc, "8", "Health and Safety")
    two_col_table(doc, [
        ("General risk assessment", "Risk_Assessment.docx in this pack. Includes lone worker section. Review annually and after any incident."),
        ("COSHH assessments", "One per hazardous substance. SDS from manufacturer. COSHHmate.co.uk (~£50/year) auto-populates from SDS data."),
        ("Written H&S policy statement", "One page. Not legally required for a sole trader without employees, but expected by insurers."),
        ("Sharps protocol", "Sharps_Protocol.docx in this pack. Review annually."),
        ("Accident/incident log", "Incident_Log.xlsx in this pack. RIDDOR-notifiable events reported to HSE within 15 days."),
        ("Manual handling", "Equipment weight target under 15 kg. Rucksack carry. Portable footrest and stool for domiciliary posture."),
    ])

    add_page_break(doc)
    chapter_heading(doc, "9", "Annual Compliance Calendar")

    calendar_rows = [
        ("SETUP (one-time, before first patient)", "Insurance; HMRC registration; ICO registration; EA Lower Tier waste carrier; ADR 1.3 training; car insurance upgrade; IPC policy signed; lone worker arrangement tested"),
        ("31 January", "Self Assessment tax return filed. Tax paid. Payment on account due."),
        ("31 July", "Second payment on account due."),
        ("ICO renewal (annual)", "Pay £52 ICO registration fee."),
        ("Insurance renewal (annual)", "Renew PI and PL. Re-confirm car insurance covers business use and clinical waste transport. Get written confirmation."),
        ("HCPC renewal (every 2 years)", "Declare indemnity. Submit CPD profile if selected for audit."),
        ("Annually — January", "Review all risk assessments. Re-sign IPC policy. Review privacy notice. Audit patient records against retention schedule."),
        ("Monthly", "Stock expiry check. Sharps container fill levels. Medicines temperature log review."),
        ("Each working day", "Complete and share Daily_Movement_Log.xlsx with emergency contact. Check-in and check-out logged."),
        ("Each patient appointment", "Clinical record entry. Consent documented. Waste segregated and sealed at point of use."),
        ("Each nail surgery episode", "Procedure-specific consent at separate pre-op appointment. Verify insurer covers home surgery."),
        ("Each carrier collection", "Sign and retain Hazardous Waste Consignment Note. Log in Clinical_Waste_Log.xlsx."),
    ]
    two_col_table(doc, calendar_rows, col1_label="When", col2_label="Action")

    add_page_break(doc)
    chapter_heading(doc, "10", "Estimated Annual Running Costs")

    cost_rows = [
        ("HCPC registration", "~£123/year"),
        ("Professional indemnity + public liability (standalone)", "£100 to £400/year"),
        ("or: RCPod membership (includes PI via James Hallam)", "£480/year"),
        ("ICO data protection registration", "£52/year"),
        ("Microsoft 365 Business Basic", "~£60/year"),
        ("Clinical waste collection (quarterly, low volume)", "£100 to £250/year"),
        ("Consumables: gloves, aprons, masks, drapes, spill kit", "£200 to £400/year"),
        ("Single-use instruments", "£150 to £400/year"),
        ("Lone worker app (MyTeamSafe or Lookout Call)", "£25 to £35/year"),
        ("Car insurance business use uplift", "£20 to £50/year"),
        ("ADR 1.3 training (one-off)", "£20 to £50 once"),
        ("EA Lower Tier waste carrier registration (one-off)", "Free"),
        ("Insulated medicines bag, lockable boot box (one-off)", "£45 to £130 once"),
        ("Total ongoing year 2+ (without RCPod membership)", "~£850 to £1,450/year"),
        ("Total ongoing year 2+ (with RCPod membership)", "~£1,100 to £1,700/year"),
    ]
    two_col_table(doc, cost_rows, col1_label="Item", col2_label="Indicative cost (5-15 patients/week)")

    add_page_break(doc)
    chapter_heading(doc, "11", "Key Contacts")

    contacts = [
        ("HCPC", "Registration, POM-A annotation check, CPD guidance", "hcpc-uk.org  |  0300 500 4472"),
        ("Royal College of Podiatry (RCPod)", "Clinical Standards 5 and 9, consent templates, CPD", "rcpod.org.uk  |  membersarea.rcpod.org.uk"),
        ("ICO", "Data protection registration, privacy notice generator", "ico.org.uk  |  0303 123 1113"),
        ("HMRC", "Self Assessment registration and filing", "gov.uk/register-for-self-assessment  |  0300 200 3310"),
        ("Environment Agency", "Lower Tier waste carrier registration, carrier register", "wastecarriersregistration.service.gov.uk"),
        ("HSE", "INDG73 (lone worker guidance), COSHH, RIDDOR reporting", "hse.gov.uk  |  0300 003 1647"),
        ("CQC", "Scope advice if uncertain about regulated activities", "03000 616161  |  cqc.org.uk"),
        ("Everywhen", "Specialist podiatry professional indemnity insurance", "everywhen.co.uk"),
        ("MyTeamSafe", "Lone worker check-in app (recommended for sole traders)", "myteamsafe.co.uk"),
        ("Lookout Call", "Lone worker check-in — telephony-based, NHS community familiar", "lookoutcall.co.uk"),
        ("EA carrier public register", "Verify any waste carrier's registration before use", "environment.data.gov.uk/public-register"),
    ]
    tbl = doc.add_table(rows=0, cols=3)
    tbl.style = 'Table Grid'
    hrow = tbl.add_row()
    for i, h in enumerate(["Body", "Purpose", "Contact"]):
        c = hrow.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    for i, (body, purpose, contact) in enumerate(contacts):
        row = tbl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '360')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        bg = hex_str(BG) if i % 2 == 0 else 'FFFFFF'
        for j, val in enumerate([body, purpose, contact]):
            c = row.cells[j]
            shade_cell(c, bg)
            cell_border_all(c, color='D4E3E4')
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            p = c.paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Calibri'
            r.font.size = Pt(9)
            r.font.bold = (j == 0)
            r.font.color.rgb = DARK if j == 0 else TEXT_C
    tbl.columns[0].width = Cm(4)
    tbl.columns[1].width = Cm(6)
    tbl.columns[2].width = Cm(6.5)

    doc.add_paragraph()
    footer_line(doc,
        "Solecare Glasgow — Private Mobile Podiatry Compliance and Administration Pack  |  Version 1.0  |  September 2026  |  Review annually")
    footer_line(doc,
        "This document does not constitute legal or regulatory advice. Verify all requirements with the relevant regulatory bodies before practice commences.")

    doc.save(os.path.join(OUT, "Compliance_Guide.docx"))
    print("  Compliance_Guide.docx — done")


if __name__ == "__main__":
    print("Building compliance guide...")
    build_guide()
