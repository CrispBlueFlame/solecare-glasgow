"""
Solecare Glasgow — Word consent forms and short policy documents.
Run with: python3 build_word_forms.py
Outputs .docx files to ./templates/
"""

import os
from docx import Document
from docx.shared import RGBColor, Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = os.path.join(os.path.dirname(__file__), "templates")

# ── Colours ──────────────────────────────────────────────────────────────────
TEAL      = RGBColor(0x2E, 0x7D, 0x83)
DARK      = RGBColor(0x1F, 0x5A, 0x5F)
BG        = RGBColor(0xF0, 0xF7, 0xF7)
BORDER_C  = RGBColor(0xD4, 0xE3, 0xE4)
TEXT_C    = RGBColor(0x1F, 0x1F, 0x1F)
MUTED     = RGBColor(0x4A, 0x4A, 0x4A)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
RED_LIGHT = RGBColor(0xFD, 0xEC, 0xEA)
RED_TEXT  = RGBColor(0x7C, 0x28, 0x20)

def hex_str(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

# ── Page setup ────────────────────────────────────────────────────────────────
def a4_doc():
    doc = Document()
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin  = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin   = Cm(2)
    section.bottom_margin = Cm(2)
    # Default paragraph style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = TEXT_C
    return doc

# ── XML helpers for cell shading and borders ──────────────────────────────────
def shade_cell(cell, hex_colour):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_colour)
    tcPr.append(shd)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            tag = OxmlElement(f'w:{edge}')
            tag.set(qn('w:val'), val.get('val', 'single'))
            tag.set(qn('w:sz'), val.get('sz', '4'))
            tag.set(qn('w:color'), val.get('color', 'D4E3E4'))
            tcBorders.append(tag)
    tcPr.append(tcBorders)

def cell_border_all(cell, color='D4E3E4', sz='4'):
    v = {'val': 'single', 'sz': sz, 'color': color}
    set_cell_border(cell, top=v, bottom=v, left=v, right=v)

def cell_no_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'bottom', 'left', 'right']:
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'none')
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=80, bottom=80, left=108, right=108):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        m = OxmlElement(f'w:{edge}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    tcPr.append(tcMar)

# ── Common heading/text helpers ───────────────────────────────────────────────
def add_doc_title(doc, title, subtitle=None):
    """Dark teal banner with white title text."""
    # Use a 1-column table as a banner
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, hex_str(DARK))
    cell_no_border(cell)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(title)
    run.font.name = 'Calibri Light'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = WHITE
    if subtitle:
        p2 = cell.add_paragraph(subtitle)
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.runs[0] if p2.runs else p2.add_run(subtitle)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0xCC, 0xE5, 0xE6)
        r2.font.italic = True
    doc.add_paragraph()

def add_section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = 'Calibri Light'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = DARK
    # Teal underline via bottom border on paragraph — simulate with a short table
    return p

def add_teal_rule(doc):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, hex_str(TEAL))
    cell_no_border(cell)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcH = OxmlElement('w:tcH')
    trH = OxmlElement('w:trH')
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trHe = OxmlElement('w:trHeight')
    trHe.set(qn('w:val'), '80')
    trHe.set(qn('w:hRule'), 'exact')
    trPr.append(trHe)
    doc.add_paragraph()

def body_para(doc, text, color=None, italic=False, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.color.rgb = color or TEXT_C
    run.font.italic = italic
    return p

def note_para(doc, text):
    return body_para(doc, text, color=MUTED, italic=True, size=9)

# ── Form field table ──────────────────────────────────────────────────────────
def form_table(doc, fields, cols=2):
    """
    fields: list of (label, hint) or (label,) for full-width fields.
    cols: 2 = label|value side by side; 1 = stacked
    Produces a clean table that Word allows typing into on laptop/iPad.
    """
    ncols = cols * 2  # label col + value col (repeated)
    tbl = doc.add_table(rows=0, cols=ncols)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Set column widths proportionally (label 35%, value 65% of half-width)
    half = Cm(8.25)  # half usable width (~16.5cm usable on A4 with 2cm margins)
    full = Cm(16.5)
    for i, col in enumerate(tbl.columns):
        if ncols == 4:
            col.width = Cm(3.5) if i % 2 == 0 else Cm(4.75)
        else:
            col.width = Cm(4.5) if i == 0 else full

    row_idx = 0
    i = 0
    while i < len(fields):
        f = fields[i]
        label = f[0]
        hint = f[1] if len(f) > 1 else ""
        full_width = f[2] if len(f) > 2 else False

        if full_width or cols == 1:
            row = tbl.add_row()
            # Merge all cols for label
            lbl_cell = row.cells[0]
            for c in range(1, ncols):
                lbl_cell = lbl_cell.merge(row.cells[c])
            shade_cell(lbl_cell, hex_str(BG))
            cell_border_all(lbl_cell)
            set_cell_margins(lbl_cell, top=60, bottom=60, left=108, right=108)
            p = lbl_cell.paragraphs[0]
            r = p.add_run(label)
            r.font.name = 'Calibri'
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = DARK
            if hint:
                rh = p.add_run(f'  {hint}')
                rh.font.name = 'Calibri'
                rh.font.size = Pt(8)
                rh.font.italic = True
                rh.font.color.rgb = MUTED
            # Value row
            row2 = tbl.add_row()
            val_cell = row2.cells[0]
            for c in range(1, ncols):
                val_cell = val_cell.merge(row2.cells[c])
            shade_cell(val_cell, 'FFFFFF')
            cell_border_all(val_cell)
            set_cell_margins(val_cell, top=100, bottom=100, left=108, right=108)
            val_cell.paragraphs[0].add_run('')
            # Make the row taller for writing space
            trPr = row2._tr.get_or_add_trPr()
            trH = OxmlElement('w:trHeight')
            trH.set(qn('w:val'), '480')
            trH.set(qn('w:hRule'), 'atLeast')
            trPr.append(trH)
            i += 1
        else:
            # Two fields side by side
            row = tbl.add_row()
            pair = fields[i:i+2]
            for j, pf in enumerate(pair):
                pl = pf[0]
                ph = pf[1] if len(pf) > 1 else ""
                lbl_c = row.cells[j * 2]
                val_c = row.cells[j * 2 + 1] if j * 2 + 1 < len(row.cells) else row.cells[j * 2]
                shade_cell(lbl_c, hex_str(BG))
                cell_border_all(lbl_c)
                set_cell_margins(lbl_c, top=80, bottom=80, left=108, right=108)
                lp = lbl_c.paragraphs[0]
                lr = lp.add_run(pl)
                lr.font.name = 'Calibri'
                lr.font.size = Pt(9)
                lr.font.bold = True
                lr.font.color.rgb = DARK
                if ph:
                    lrh = lp.add_run(f'\n{ph}')
                    lrh.font.size = Pt(8)
                    lrh.font.italic = True
                    lrh.font.color.rgb = MUTED
                shade_cell(val_c, 'FFFFFF')
                cell_border_all(val_c)
                set_cell_margins(val_c, top=80, bottom=80, left=108, right=108)
                val_c.paragraphs[0].add_run('')
            if len(pair) == 1 and ncols == 4:
                # Fill remaining cells
                shade_cell(row.cells[2], hex_str(BG))
                cell_border_all(row.cells[2])
                shade_cell(row.cells[3], 'FFFFFF')
                cell_border_all(row.cells[3])
            i += len(pair)
    doc.add_paragraph()
    return tbl

def signature_table(doc, signatories):
    """
    signatories: list of (role_label,)
    Creates signature/date/name rows — works for iPad pencil signing in Word.
    """
    ncols = len(signatories) * 3  # name | sig | date per signatory
    tbl = doc.add_table(rows=0, cols=len(signatories) * 3)
    tbl.style = 'Table Grid'

    # Header row
    hrow = tbl.add_row()
    usable = Cm(16.5)
    col_w = usable / len(signatories)
    for j, (role,) in enumerate(signatories):
        c = hrow.cells[j * 3]
        end_c = hrow.cells[j * 3 + 2]
        merged = c.merge(end_c)
        shade_cell(merged, hex_str(TEAL))
        cell_border_all(merged, color=hex_str(TEAL))
        set_cell_margins(merged, top=80, bottom=80, left=108, right=108)
        p = merged.paragraphs[0]
        r = p.add_run(role)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = WHITE

    for label, width_hint in [("Full name:", 5), ("Signature:", 8), ("Date:", 3)]:
        row = tbl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '560')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        for j in range(len(signatories)):
            lbl_c = row.cells[j * 3]
            shade_cell(lbl_c, hex_str(BG))
            cell_border_all(lbl_c)
            set_cell_margins(lbl_c, top=80, bottom=80, left=108, right=108)
            p = lbl_c.paragraphs[0]
            r = p.add_run(label)
            r.font.name = 'Calibri'
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = DARK

            mid_c = row.cells[j * 3 + 1]
            shade_cell(mid_c, 'FFFFFF')
            cell_border_all(mid_c)
            end_c = row.cells[j * 3 + 2]
            merged_val = mid_c.merge(end_c)
            shade_cell(merged_val, 'FFFFFF')
            cell_border_all(merged_val)
            set_cell_margins(merged_val, top=80, bottom=80, left=108, right=108)

    doc.add_paragraph()

def checkbox_para(doc, text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run('☐  ')
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = TEAL
    run2 = p.add_run(text)
    run2.font.name = 'Calibri'
    run2.font.size = Pt(10)
    run2.font.color.rgb = TEXT_C
    return p

def notice_box(doc, text, bg=None, text_color=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, hex_str(bg or BG))
    set_cell_border(cell,
        top={'val': 'single', 'sz': '12', 'color': hex_str(TEAL)},
        bottom={'val': 'single', 'sz': '12', 'color': hex_str(TEAL)},
        left={'val': 'single', 'sz': '12', 'color': hex_str(TEAL)},
        right={'val': 'single', 'sz': '12', 'color': hex_str(TEAL)},
    )
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(9)
    run.font.color.rgb = text_color or TEXT_C
    doc.add_paragraph()

def footer_line(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(8)
    run.font.italic = True
    run.font.color.rgb = MUTED
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ════════════════════════════════════════════════════════════════════════════
# CONSENT — PATIENT REGISTRATION & ROUTINE TREATMENT
# ════════════════════════════════════════════════════════════════════════════
def build_consent_routine():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow",
        "Patient Registration and General Treatment Consent")
    note_para(doc,
        "Please complete all sections. This form is used for registration and for consent to routine podiatry treatment (nail care, skin care, routine assessment). "
        "A separate form is required before any nail surgery procedure. You may complete this form on screen or print and handwrite. Your information is stored in accordance with UK GDPR — see our Privacy Notice.")

    add_section_heading(doc, "1. Personal Details")
    add_teal_rule(doc)
    form_table(doc, [
        ("Full name:", "As per ID"),
        ("Date of birth:", "DD/MM/YYYY"),
        ("Address:", "Including postcode", True),
        ("Mobile number:", ""),
        ("Email address:", "For appointment reminders and forms only"),
        ("How did you hear about us?:", ""),
    ])

    add_section_heading(doc, "2. GP and Emergency Contact")
    add_teal_rule(doc)
    form_table(doc, [
        ("GP name:", ""),
        ("GP practice name and address:", "", True),
        ("GP phone:", ""),
        ("Emergency contact name:", ""),
        ("Emergency contact phone:", ""),
        ("Relationship:", ""),
    ])

    add_section_heading(doc, "3. Medical History")
    add_teal_rule(doc)
    note_para(doc,
        "Please list any conditions, diagnoses, or health matters relevant to your foot health and treatment. "
        "Include any history of diabetes, peripheral vascular disease, neuropathy, bleeding disorders, anticoagulant therapy, or immune-suppressing conditions.")
    form_table(doc, [
        ("Relevant medical conditions / diagnoses:", "", True),
        ("Current medications (including blood thinners, steroids, immunosuppressants):", "", True),
        ("Known allergies (including latex, antiseptics, dressings, local anaesthetics):", "", True),
        ("Any previous adverse reactions to podiatry treatment?", "Please describe", True),
    ])

    add_section_heading(doc, "4. Circulation and Sensation")
    add_teal_rule(doc)
    note_para(doc, "These questions help your podiatrist plan safe treatment.")
    form_table(doc, [
        ("Do you experience cold feet, colour changes, or pain in legs when walking?", "Y / N / Details"),
        ("Do you experience numbness, tingling, or loss of sensation in your feet?", "Y / N / Details"),
        ("Are you a current or recent smoker?", "Y / N"),
        ("Date of last diabetic foot check (if applicable):", ""),
    ])

    add_section_heading(doc, "5. Consent to Routine Treatment")
    add_teal_rule(doc)
    notice_box(doc,
        "Your podiatrist will discuss your treatment with you and answer any questions before and during each appointment. "
        "You may withdraw this consent at any time, and your care will continue. This consent covers routine (non-surgical) podiatry treatment only.")
    checkbox_para(doc, "I confirm I have provided accurate information above and will inform my podiatrist of any changes.")
    checkbox_para(doc, "I consent to routine podiatry treatment including assessment, nail care, skin care, and appropriate dressings.")
    checkbox_para(doc, "I understand that a separate consent form will be completed before any nail surgery procedure.")
    checkbox_para(doc, "I have received the Privacy Notice and understand how my information will be stored and used.")
    checkbox_para(doc, "I consent to my podiatrist contacting my GP if they identify a clinical concern requiring medical review.")

    doc.add_paragraph()
    add_section_heading(doc, "6. Signatures")
    add_teal_rule(doc)
    note_para(doc, "Please sign below. On an iPad, tap the Signature field and use Apple Pencil or your finger to draw your signature in Microsoft Word.")
    signature_table(doc, [("Patient (or legal guardian if under 16)",), ("Practitioner",)])

    footer_line(doc, "Solecare Glasgow  |  Review annually or when clinical status changes significantly  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Consent_Routine.docx"))
    print("  Consent_Routine.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# CONSENT — NAIL SURGERY
# ════════════════════════════════════════════════════════════════════════════
def build_consent_nail_surgery():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow",
        "Nail Surgery Procedure Consent — One form per surgical episode")
    note_para(doc,
        "This form must be completed at a pre-operative appointment, not immediately before surgery. "
        "Montgomery v Lanarkshire [2015] standard applies: all material risks must be explained and your questions answered before you sign.")

    add_section_heading(doc, "1. Patient and Procedure Details")
    add_teal_rule(doc)
    form_table(doc, [
        ("Patient full name:", ""),
        ("Date of birth:", "DD/MM/YYYY"),
        ("Date of this consent form:", "DD/MM/YYYY"),
        ("Proposed surgery date:", "DD/MM/YYYY  (if known)"),
        ("Procedure(s) planned:", "e.g. Partial nail avulsion with phenolisation, right hallux", True),
        ("Digit(s) / side:", "e.g. Right hallux (big toe)"),
        ("Local anaesthetic to be used:", "e.g. Lidocaine 2%"),
        ("Any known allergy to local anaesthetic?", "Y / N / Type:"),
    ])

    add_section_heading(doc, "2. Information Provided to Patient")
    add_teal_rule(doc)
    note_para(doc, "Your podiatrist confirms that the following has been explained and discussed:")
    checkbox_para(doc, "The nature of the procedure and what it involves step by step.")
    checkbox_para(doc, "Why this procedure is recommended and what alternatives exist (including no treatment).")
    checkbox_para(doc, "Local anaesthetic: how it is administered, expected discomfort, and duration of numbness.")
    checkbox_para(doc, "Common risks: post-operative infection, bleeding, bruising, temporary nail regrowth changes.")
    checkbox_para(doc, "Less common risks: incomplete phenolisation requiring further treatment, persistent nail abnormality, allergic reaction.")
    checkbox_para(doc, "Post-operative care instructions including dressing changes, footwear, activity, and signs of infection.")
    checkbox_para(doc, "When to seek urgent advice: signs of spreading infection, unexpected bleeding, allergic reaction.")
    checkbox_para(doc, "The patient's right to withdraw consent at any time, including on the day of surgery.")

    add_section_heading(doc, "3. Patient's Questions and Answers")
    add_teal_rule(doc)
    note_para(doc, "Record any specific questions the patient asked and how they were answered:")
    form_table(doc, [
        ("Questions asked:", "", True),
        ("Answers given:", "", True),
    ])

    add_section_heading(doc, "4. Patient Consent")
    add_teal_rule(doc)
    notice_box(doc,
        "I confirm that I have been given sufficient time to consider this information, that my questions have been answered to my satisfaction, "
        "and that I understand I may withdraw consent at any time without it affecting the quality of my care.")
    checkbox_para(doc, "I consent to the procedure(s) described above under local anaesthetic.")
    checkbox_para(doc, "I understand the risks explained to me and accept them as part of this procedure.")
    checkbox_para(doc, "I have received and understood the post-operative care instructions.")
    checkbox_para(doc, "I consent to my GP being informed of this procedure.")

    doc.add_paragraph()
    add_section_heading(doc, "5. Signatures")
    add_teal_rule(doc)
    note_para(doc,
        "Both signatures are required before surgery proceeds. If the patient is under 16, a parent or guardian with parental responsibility must sign.")
    signature_table(doc, [("Patient (or parent/guardian if under 16)",), ("Practitioner (HCPC registered)",)])

    add_section_heading(doc, "6. On the Day of Surgery — Verification")
    add_teal_rule(doc)
    note_para(doc, "Complete immediately before the procedure:")
    form_table(doc, [
        ("Consent form reviewed with patient?", "Y / N"),
        ("Patient confirms consent still given?", "Y / N"),
        ("Allergies re-checked?", "Y / N"),
        ("LA batch number and expiry checked?", "Y / N — Batch no.:"),
        ("Surgical site confirmed (correct toe/side)?", "Y / N — Site confirmed:"),
    ])
    signature_table(doc, [("Practitioner pre-op check",)])

    footer_line(doc, "Solecare Glasgow  |  Retain with patient record for minimum 8 years (adult) or until patient's 26th birthday (child)  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Consent_NailSurgery.docx"))
    print("  Consent_NailSurgery.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# CONSENT — CLINICAL PHOTOGRAPHY
# ════════════════════════════════════════════════════════════════════════════
def build_consent_photography():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow",
        "Clinical Photography Consent")
    note_para(doc,
        "Clinical photographs are sometimes taken to document a patient's condition or monitor treatment progress. "
        "This is a separate consent to treatment consent. One form per photography session.")

    add_section_heading(doc, "1. Patient Details")
    add_teal_rule(doc)
    form_table(doc, [
        ("Patient full name:", ""),
        ("Date of birth:", "DD/MM/YYYY"),
        ("Date photographs taken:", "DD/MM/YYYY"),
        ("Body area photographed:", "e.g. Right hallux, plantar surface both feet"),
    ])

    add_section_heading(doc, "2. Purpose of Photography")
    add_teal_rule(doc)
    note_para(doc, "Please tick all that apply:")
    checkbox_para(doc, "Clinical record — to document the condition and monitor progress.")
    checkbox_para(doc, "Referral — to accompany a referral letter to another clinician.")
    checkbox_para(doc, "Audit / quality improvement — anonymous, not used outside the practice.")
    checkbox_para(doc, "CPD / case study — images will be anonymised before any use.")
    checkbox_para(doc, "Other (please specify below):")
    form_table(doc, [("Other purpose:", "", True)])

    add_section_heading(doc, "3. Storage and Access")
    add_teal_rule(doc)
    notice_box(doc,
        "Photographs will be stored securely in your encrypted clinical record on OneDrive (Microsoft 365 Business), "
        "accessible only by your podiatrist. They will not be shared with any third party without your explicit written consent, "
        "except where required by law or professional duty to protect health and safety. "
        "Photographs will be retained for the same period as your clinical records and then securely deleted.")

    add_section_heading(doc, "4. Consent")
    add_teal_rule(doc)
    checkbox_para(doc, "I consent to clinical photographs being taken for the purpose(s) ticked above.")
    checkbox_para(doc, "I understand how the photographs will be stored and who will have access.")
    checkbox_para(doc, "I understand I can withdraw consent and request deletion of photographs at any time by contacting my podiatrist.")
    form_table(doc, [("Any conditions or restrictions I wish to apply:", "", True)])

    doc.add_paragraph()
    signature_table(doc, [("Patient (or parent/guardian if under 16)",), ("Practitioner",)])
    footer_line(doc, "Solecare Glasgow  |  Retain with patient record  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Consent_Photography.docx"))
    print("  Consent_Photography.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# PRIVACY NOTICE
# ════════════════════════════════════════════════════════════════════════════
def build_privacy_notice():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Privacy Notice",
        "How we collect, use, and protect your personal information")
    note_para(doc,
        "This notice explains your rights and how we use your information. Please read it before your first appointment. "
        "A copy is available on request at any time.")

    sections = [
        ("Who we are",
         "Solecare Glasgow is a sole-trader mobile podiatry practice operated by [PRACTITIONER NAME], "
         "HCPC-registered chiropodist/podiatrist (registration number [HCPC NUMBER]). "
         "[YOUR BUSINESS ADDRESS / CONTACT EMAIL]. "
         "We are registered with the Information Commissioner's Office (ICO), registration number [ICO NUMBER]."),

        ("What information we collect",
         "We collect: your name, date of birth, contact details and address; "
         "medical history, current medications, and allergy information; "
         "clinical assessment findings and treatment records; "
         "appointment history; payment information (not stored in clinical records); "
         "and where consent is given, clinical photographs."),

        ("Why we use it and our legal basis",
         "We use your information to provide podiatry care. Our lawful basis is:\n"
         "• Personal data (Article 6(1)(b) UK GDPR): necessary to perform the service you have contracted.\n"
         "• Health data (Article 9(2)(h) UK GDPR): necessary for healthcare purposes by an HCPC-registered professional.\n"
         "We do not use your data for marketing without separate consent."),

        ("How long we keep it",
         "Adult records: 8 years from your last appointment.\n"
         "Children's records: until the patient's 25th birthday (or 26th if aged 17 at the end of treatment).\n"
         "Records are securely deleted at the end of the retention period."),

        ("Who sees your information",
         "Your records are accessed only by your podiatrist. We will share with other clinicians only when clinically necessary "
         "(e.g. GP referral) or required by law. We do not sell your data or share it with third parties for commercial purposes.\n"
         "Records are stored on Microsoft 365 Business (encrypted, UK/EEA data centres; Microsoft holds a standard Data Processing Agreement with us)."),

        ("Your rights",
         "You have the right to: access your records (subject access request — respond within 1 month, no charge); "
         "correct inaccurate information; request deletion (subject to our legal retention obligations); "
         "restrict or object to processing; data portability.\n"
         "To exercise any right, contact us at [CONTACT EMAIL]. "
         "If you are unhappy with how we handle your data, you can complain to the ICO at ico.org.uk or 0303 123 1113."),

        ("Cookies and website",
         "If you contact us via our website, the site may store a cookie for your session. No personal health data is held on the website."),

        ("Changes to this notice",
         "We review this notice annually. The current version will always be available on request. Last reviewed: September 2026."),
    ]

    for heading, body in sections:
        add_section_heading(doc, heading)
        add_teal_rule(doc)
        body_para(doc, body)

    doc.add_paragraph()
    notice_box(doc,
        "Practitioner: [PRACTITIONER NAME]  |  HCPC: [NUMBER]  |  ICO: [NUMBER]  |  Email: [EMAIL]  |  "
        "Please complete the fields in [brackets] before distributing this notice.")

    footer_line(doc, "Solecare Glasgow  |  Privacy Notice  |  Version 1.0  |  September 2026  |  Next review: September 2027")
    doc.save(os.path.join(OUT, "Privacy_Notice.docx"))
    print("  Privacy_Notice.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# COMPLAINTS PROCEDURE
# ════════════════════════════════════════════════════════════════════════════
def build_complaints():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Complaints Procedure",
        "We take all concerns seriously. Here is how to raise one.")

    body_para(doc,
        "We aim to provide professional, safe, and respectful podiatry care at every appointment. "
        "If you have a concern about any aspect of your treatment or our service, please tell us. "
        "Most concerns can be resolved quickly with a conversation.")

    add_section_heading(doc, "How to raise a concern")
    add_teal_rule(doc)
    body_para(doc, "Step 1 — Contact us directly")
    body_para(doc,
        "Speak to your podiatrist at any time, or contact us in writing at [CONTACT EMAIL]. "
        "We will acknowledge your concern within 3 working days and respond fully within 20 working days. "
        "If we need more time, we will tell you why and give a revised date.")
    body_para(doc, "Step 2 — If you remain unsatisfied")
    body_para(doc,
        "If we cannot resolve your concern to your satisfaction, you have the right to escalate to the "
        "Health and Care Professions Council (HCPC), which regulates podiatrists in the UK:")

    form_table(doc, [
        ("HCPC — fitness to practise concerns:", "fitness@hcpc-uk.org  |  0300 500 4472  |  hcpc-uk.org"),
        ("ICO — data protection concerns:", "casework@ico.org.uk  |  0303 123 1113  |  ico.org.uk"),
    ])

    add_section_heading(doc, "Our commitment")
    add_teal_rule(doc)
    checkbox_para(doc, "We will acknowledge every complaint within 3 working days.")
    checkbox_para(doc, "We will investigate thoroughly and respond honestly.")
    checkbox_para(doc, "We will not treat you differently as a result of raising a concern.")
    checkbox_para(doc, "We keep a confidential log of all complaints to improve our service.")

    add_section_heading(doc, "Complaint log — for practitioner use")
    add_teal_rule(doc)
    form_table(doc, [
        ("Date received:", ""),
        ("Nature of complaint (brief summary):", "", True),
        ("Date acknowledged:", ""),
        ("Action taken / response:", "", True),
        ("Date resolved / closed:", ""),
        ("Learning identified:", "", True),
    ])

    footer_line(doc, "Solecare Glasgow  |  Complaints Procedure  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Complaints_Procedure.docx"))
    print("  Complaints_Procedure.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# LONE WORKER EMERGENCY PROCEDURES
# ════════════════════════════════════════════════════════════════════════════
def build_lone_worker_emergency():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Lone Worker Emergency Procedures",
        "Give a copy to your named emergency contact. Keep a copy in your kit bag.")

    notice_box(doc,
        "This document is the emergency contact's action plan. Complete the details below before the first working day. "
        "Review contact details and escalation steps at least annually.",
        bg=BG)

    add_section_heading(doc, "Practitioner details")
    add_teal_rule(doc)
    form_table(doc, [
        ("Practitioner name:", ""),
        ("Mobile (primary):", ""),
        ("Vehicle registration:", ""),
        ("Vehicle make/model/colour:", ""),
        ("Professional indemnity insurer:", ""),
        ("PI insurer emergency number:", ""),
    ])

    add_section_heading(doc, "Emergency contact details")
    add_teal_rule(doc)
    form_table(doc, [
        ("Emergency contact name:", ""),
        ("Relationship:", ""),
        ("Mobile:", ""),
        ("Home / alternative number:", ""),
    ])

    add_section_heading(doc, "Daily procedure")
    add_teal_rule(doc)
    body_para(doc,
        "Before each working day, the practitioner will share a copy of the Daily Movement Log with the emergency contact "
        "(by text, email, or WhatsApp). The log includes: all patient addresses for the day, appointment times, "
        "and expected finish times. The practitioner will send a check-in message after each patient visit and "
        "an end-of-day confirmation when home safely.")

    add_section_heading(doc, "Escalation steps — if a check-in is missed")
    add_teal_rule(doc)

    steps = [
        ("Step 1 — Wait 15 minutes past the expected check-in time.",
         "A delayed appointment or traffic may be the cause. Do not call the patient address at this stage."),
        ("Step 2 — Call the practitioner's mobile.",
         "If no answer, leave a message and try again after 10 minutes."),
        ("Step 3 — If no answer after two attempts (25 minutes total),",
         "call the last known patient address from the Movement Log. Ask whether the practitioner has left."),
        ("Step 4 — If location is still unknown after a further 15 minutes (40 minutes total),",
         "call 999. Give them: practitioner name, vehicle description and registration, last known address from the Movement Log, and this document."),
        ("Step 5 — Notify the PI insurer",
         "using the number above. They can advise on further steps."),
    ]
    for step_title, step_body in steps:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(0)
        r1 = p.add_run(step_title + " ")
        r1.font.name = 'Calibri'
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = DARK
        r2 = p.add_run(step_body)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(10)
        r2.font.color.rgb = TEXT_C

    doc.add_paragraph()
    notice_box(doc,
        "If in doubt at any stage — call 999. Better a false alarm than a delayed response.",
        bg=RED_LIGHT, text_color=RED_TEXT)

    add_section_heading(doc, "Signatures — both parties confirm they have read and agreed this plan")
    add_teal_rule(doc)
    signature_table(doc, [("Practitioner",), ("Emergency contact",)])

    footer_line(doc, "Solecare Glasgow  |  Lone Worker Emergency Procedures  |  Review annually  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Lone_Worker_Emergency.docx"))
    print("  Lone_Worker_Emergency.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# SHARPS PROTOCOL
# ════════════════════════════════════════════════════════════════════════════
def build_sharps_protocol():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Sharps Protocol",
        "Health and Safety (Sharp Instruments in Healthcare) Regulations 2013  |  Environmental Protection Act 1990")

    add_section_heading(doc, "Principles")
    add_teal_rule(doc)
    body_para(doc,
        "All sharps used in this practice (scalpel blades, needles, lancets) are single-use only. "
        "No resheathing of needles. No passing sharps between hands. Dispose at the point of use, immediately.")

    add_section_heading(doc, "Container assembly and management")
    add_teal_rule(doc)
    items = [
        "Assemble a new sharps container at the start of each working period. Record in the Sharps Container Log (IPC_Visit_Log.xlsx).",
        "Check the fill level before each visit. Seal the container when it reaches the three-quarter full mark — never overfill.",
        "The container must be UN3291-approved and BS 7320 compliant (yellow lid, rigid sides).",
        "Seal, label, and date the container when three-quarters full. Record date and container ref in the log.",
        "Transport sealed containers inside a rigid, lockable outer transport box in the vehicle. "
        "Never transport loose bags. Never transport open or overfilled containers.",
        "Store sealed containers at the base address in a secure, labelled location separate from household waste until carrier collection.",
        "Arrange carrier collection promptly. Issue a Hazardous Waste Consignment Note. Retain the signed copy for 3 years.",
    ]
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(item)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.color.rgb = TEXT_C

    add_section_heading(doc, "Sharps injury response — IMMEDIATE ACTION")
    add_teal_rule(doc)
    notice_box(doc,
        "1. STOP the procedure immediately.\n"
        "2. Remove gloves carefully.\n"
        "3. Wash the wound under running water for at least 2 minutes. Do not suck the wound.\n"
        "4. Encourage gentle bleeding if the wound is puncture-type — do not squeeze.\n"
        "5. Cover with a waterproof plaster.\n"
        "6. Document immediately in Incident_Log.xlsx and the Sharps Injury Detail sheet.\n"
        "7. Call the patient's GP or A&E within 1 hour for BBV (blood-borne virus) risk assessment and advice on PEP if indicated.\n"
        "8. Notify your professional indemnity insurer if the patient's blood-borne virus status is unknown.\n"
        "9. Review risk assessment and update sharps handling procedure if the injury reveals a hazard.",
        bg=RED_LIGHT, text_color=RED_TEXT)

    add_section_heading(doc, "Sharps injury record — complete immediately")
    add_teal_rule(doc)
    note_para(doc, "Also complete the full Sharps Injury Detail sheet in Incident_Log.xlsx.")
    form_table(doc, [
        ("Date and time:", ""),
        ("Device type and procedure:", ""),
        ("Visible blood contamination?", "Y / N"),
        ("Patient reference:", ""),
        ("Immediate action taken:", "", True),
        ("Medical advice sought (time/provider):", ""),
        ("BBV risk assessment outcome:", "", True),
        ("Follow-up required?", "Y / N / Dates:"),
    ])
    signature_table(doc, [("Practitioner signature and date",)])

    footer_line(doc, "Solecare Glasgow  |  Sharps Protocol  |  Review annually  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Sharps_Protocol.docx"))
    print("  Sharps_Protocol.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# RUN ALL
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("Building Word forms and short documents...")
    build_consent_routine()
    build_consent_nail_surgery()
    build_consent_photography()
    build_privacy_notice()
    build_complaints()
    build_lone_worker_emergency()
    build_sharps_protocol()
    print("All Word forms built.")
