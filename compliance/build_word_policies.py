"""
Solecare Glasgow — Infection Control Policy and Risk Assessment.
Run with: python3 build_word_policies.py
Outputs .docx files to ./templates/
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
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

OUT = os.path.join(os.path.dirname(__file__), "templates")

def bullet(doc, text, indent=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10)
    run.font.color.rgb = TEXT_C

def bold_body(doc, label, body):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r1 = p.add_run(label + ' ')
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = DARK
    r2 = p.add_run(body)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(10)
    r2.font.color.rgb = TEXT_C

def risk_table(doc, risks):
    """
    risks: list of (hazard, who, controls, likelihood, severity, risk_level, additional)
    """
    tbl = doc.add_table(rows=0, cols=7)
    tbl.style = 'Table Grid'
    hdr = tbl.add_row()
    headers = ['Hazard', 'Who at risk', 'Existing controls', 'L\n(1-3)', 'S\n(1-3)', 'Risk\n(LxS)', 'Additional action']
    widths = [Cm(3), Cm(2.2), Cm(4.5), Cm(1), Cm(1), Cm(1), Cm(3.8)]
    for i, (h, w) in enumerate(zip(headers, widths)):
        c = hdr.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'
        r.font.size = Pt(8)
        r.font.bold = True
        r.font.color.rgb = WHITE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, row_data in enumerate(risks):
        row = tbl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '360')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        bg = hex_str(BG) if i % 2 == 0 else 'FFFFFF'
        for j, val in enumerate(row_data):
            c = row.cells[j]
            shade_cell(c, bg)
            cell_border_all(c, color='D4E3E4')
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            p = c.paragraphs[0]
            r = p.add_run(str(val))
            r.font.name = 'Calibri'
            r.font.size = Pt(8)
            r.font.color.rgb = TEXT_C
            if j in [3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if j == 5:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Colour the risk score
                score = int(val) if str(val).isdigit() else 0
                r.font.bold = True
                if score >= 6:
                    r.font.color.rgb = RED_TEXT
                elif score >= 3:
                    r.font.color.rgb = RGBColor(0xB8, 0x5C, 0x00)
                else:
                    r.font.color.rgb = RGBColor(0x1B, 0x5E, 0x20)
    doc.add_paragraph()


# ════════════════════════════════════════════════════════════════════════════
# INFECTION CONTROL POLICY
# ════════════════════════════════════════════════════════════════════════════
def build_ipc_policy():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Infection Control Policy",
        "Covers standard infection control precautions and domiciliary (home visit) practice")

    notice_box(doc,
        "Reference standards: NHS National Infection Prevention and Control Manual (NIPCM) v2.10, May 2024  |  "
        "RCPod Clinical Standards 5 (Infection Control) and 9 (Domiciliary Care)  |  "
        "HCPC Standards of Conduct, Performance and Ethics 2024, Standards 6.1 and 6.2")

    form_table(doc, [
        ("Practitioner:", ""),
        ("HCPC registration no.:", ""),
        ("Policy version:", "1.0"),
        ("Date written:", "September 2026"),
        ("Next review date:", "September 2027  (review annually or after any IPC incident)"),
    ])

    add_section_heading(doc, "1. Scope and purpose")
    add_teal_rule(doc)
    body_para(doc,
        "This policy applies to all podiatry treatment provided by Solecare Glasgow, including all domiciliary (home visit) appointments. "
        "It sets out the standard infection control precautions (SICPs) used at every patient contact and the additional measures "
        "specific to working in uncontrolled domestic environments. "
        "Compliance with this policy is mandatory at every visit.")

    add_section_heading(doc, "2. Instruments — single-use only")
    add_teal_rule(doc)
    body_para(doc,
        "All instruments used in this practice are single-use, sterile, and sourced from a licensed medical device supplier. "
        "No reusable instruments are in use. No decontamination equipment (autoclave) is required or held. "
        "The single-use symbol on packaging (a circle with a line through a '2') confirms single-use designation. "
        "Used instruments are disposed of as clinical waste immediately after use. "
        "Instruments are never reused between patients under any circumstance.")

    add_section_heading(doc, "3. Hand hygiene")
    add_teal_rule(doc)
    body_para(doc,
        "Hand hygiene is performed at the five WHO moments: before patient contact, before a clean/aseptic procedure, "
        "after body fluid exposure, after patient contact, and after contact with the patient's surroundings.")
    bold_body(doc, "ABHR (alcohol-based hand rub, minimum 60% alcohol):",
        "used when hands are not visibly soiled and there is no known C. difficile or norovirus risk.")
    bold_body(doc, "Soap and water (patient's kitchen or bathroom sink):",
        "mandatory when hands are visibly soiled; after body fluid contact; suspected C. difficile or norovirus exposure.")
    body_para(doc,
        "In domiciliary settings, the patient's domestic sink is the accepted clinical equivalent of a clinical hand-wash basin. "
        "This is a documented control, not a compliance gap. ABHR is carried at all times as a supplement.")
    body_para(doc, "Hand hygiene is performed after removing PPE (doffing) before leaving the treatment area.")

    add_section_heading(doc, "4. PPE selection")
    add_teal_rule(doc)
    body_para(doc, "PPE is selected according to the procedure risk:")

    ppe_rows = [
        ("Routine nail care, skin assessment, dressings", "Nitrile gloves + disposable apron"),
        ("Nail surgery, debridement with bleeding risk", "Nitrile gloves + apron + Type IIR fluid-resistant mask + eye protection"),
        ("Body fluid spill management", "Heavy nitrile gloves + apron + face shield"),
        ("All visits where household infectious illness reported", "Gloves + apron + Type IIR mask (minimum)"),
    ]
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = 'Table Grid'
    hrow = tbl.add_row()
    for i, h in enumerate(["Procedure", "Minimum PPE"]):
        c = hrow.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    for i, (proc, ppe) in enumerate(ppe_rows):
        row = tbl.add_row()
        for j, val in enumerate([proc, ppe]):
            c = row.cells[j]
            shade_cell(c, hex_str(BG) if i % 2 == 0 else 'FFFFFF')
            cell_border_all(c)
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            p = c.paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.color.rgb = TEXT_C
    doc.add_paragraph()

    body_para(doc,
        "Gloves are changed between patients. Gloves are changed within a patient appointment if moving from a contaminated to a clean procedure. "
        "Gloves are never reused or washed. PPE is donned before patient contact and doffed before leaving the treatment area. "
        "All used PPE is disposed of as clinical waste.")

    add_section_heading(doc, "5. Domiciliary clean field")
    add_teal_rule(doc)
    body_para(doc,
        "A disposable paper drape (minimum 45 x 60 cm) is used as a clean field at every visit before instruments are placed out. "
        "The drape is placed on any available surface — kitchen table, tray, footstool. "
        "The surface beneath is never treated as clean. "
        "Instruments remain on the drape until used, and the drape is disposed of as clinical waste after use. "
        "Clean and contaminated items are never mixed. The practitioner works in one direction — clean to contaminated — and does not cross back.")

    add_section_heading(doc, "6. Nail surgery skin preparation")
    add_teal_rule(doc)
    body_para(doc,
        "Skin antisepsis for nail surgery follows NICE NG125: 0.5% chlorhexidine gluconate in 70% isopropyl alcohol "
        "(e.g. ChloraPrep applicator wand or equivalent). Apply to the surgical site with a single-use sterile applicator. "
        "Allow a minimum of 30 seconds to air-dry before incision. Do not blot or fan. Do not use spirits alone. "
        "Failure to allow full drying creates a fire risk and reduces antiseptic efficacy. "
        "A sterile drape is applied after the site has dried.")

    add_section_heading(doc, "7. Body fluid spill management")
    add_teal_rule(doc)
    bold_body(doc, "Blood or body fluid on hard surface:", "")
    bullet(doc, "Don heavy nitrile gloves, apron, face shield before approaching the spill.")
    bullet(doc, "Apply absorbent paper towels or granules to the spill.")
    bullet(doc, "Apply 10,000 ppm sodium hypochlorite solution or wipes; leave 3 minutes contact time.")
    bullet(doc, "Remove material and clean surface. Dispose as clinical waste.")
    bullet(doc, "Remove PPE. Perform hand hygiene.")
    bold_body(doc, "Urine spill or spill on soft furnishings:",
        "Use 1,000 ppm hypochlorite or detergent only — 10,000 ppm hypochlorite damages soft surfaces and reacts with urine.")
    body_para(doc, "Document all spill events in IPC_Visit_Log.xlsx.")

    add_section_heading(doc, "8. Clinical waste")
    add_teal_rule(doc)
    body_para(doc,
        "All clinical waste is segregated at point of generation and never placed in the patient's household waste bin. "
        "Waste streams: sharps into UN3291-approved sharps container at point of use; infectious soft waste into yellow/orange clinical waste bags; "
        "offensive waste (dry nail clippings, unsoiled PPE) into tiger-stripe bags. "
        "Sealed bags are transported in a rigid, lockable outer container in the vehicle. "
        "See Clinical_Waste_Log.xlsx for consignment note records.")

    add_section_heading(doc, "9. Pre-visit screening")
    add_teal_rule(doc)
    body_para(doc,
        "Before a first visit and after any reported household illness, conduct a brief telephone screening call: "
        "Is there active diarrhoea or vomiting in the household? Has any household member been recently diagnosed with a communicable infection? "
        "Are there any recent changes to the patient's health? If active gastrointestinal illness is reported, defer the appointment by at least 48 hours after the last symptom.")

    add_section_heading(doc, "10. Hepatitis B immunity")
    add_teal_rule(doc)
    body_para(doc,
        "Documented evidence of Hepatitis B immunity is held on file. If immunity cannot be confirmed, the practitioner will consult their GP before treating patients.")

    add_section_heading(doc, "11. Annual review and CPD")
    add_teal_rule(doc)
    body_para(doc,
        "This policy is reviewed and re-signed annually, and after any IPC incident. "
        "At least one IPC-related CPD entry is completed per calendar year and logged in CPD_Log.xlsx.")

    doc.add_paragraph()
    signature_table(doc, [("Practitioner — annual review signature and date",)])
    footer_line(doc, "Solecare Glasgow  |  Infection Control Policy  |  Review annually  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Infection_Control_Policy.docx"))
    print("  Infection_Control_Policy.docx — done")


# ════════════════════════════════════════════════════════════════════════════
# HEALTH AND SAFETY RISK ASSESSMENT
# ════════════════════════════════════════════════════════════════════════════
def build_risk_assessment():
    doc = a4_doc()
    add_doc_title(doc, "Solecare Glasgow — Health and Safety Risk Assessment",
        "HSWA 1974  |  MHSWR 1999  |  COSHH 2002  |  Manual Handling Operations Regulations 1992  |  HSE INDG73")

    notice_box(doc,
        "Risk scoring: Likelihood (L) 1=unlikely, 2=possible, 3=likely  |  "
        "Severity (S) 1=minor, 2=moderate, 3=serious  |  "
        "Risk = L x S  |  1-2 = Low (green)  |  3-4 = Medium (amber)  |  6-9 = High (red — immediate action)")

    form_table(doc, [
        ("Practitioner:", ""),
        ("Practice type:", "Mobile sole-trader domiciliary podiatry — single-use instruments"),
        ("Assessment date:", ""),
        ("Next review date:", "Annual, or after any incident or significant change in practice"),
    ])

    add_section_heading(doc, "Section A — General and Domiciliary Hazards")
    add_teal_rule(doc)
    risks_a = [
        ("Sharps injury — scalpel blades, needles",
         "Practitioner", "Single-use only. Dispose at point of use into UN3291 sharps bin. Never resheath. Written protocol.", 1, 3, 3, "Annual sharps protocol review"),
        ("Blood/body fluid exposure",
         "Practitioner", "PPE: gloves + apron + mask/eye protection for surgical procedures. Spill kit carried. Post-exposure protocol documented.", 1, 3, 3, "BBV immunisation confirmed"),
        ("Infection — patient cross-contamination",
         "Patient, Practitioner", "SICPs at all visits. Single-use instruments. Clean field. Pre-visit screening call. Full IPC policy in place.", 1, 3, 3, "Annual IPC policy review"),
        ("Unstable seating / working posture (domiciliary)",
         "Practitioner", "Portable footrest and folding stool carried. Assess environment on arrival. Do not treat from floor.", 2, 2, 4, "MSK check-in annually"),
        ("Manual handling — equipment transport",
         "Practitioner", "Kit bag weight target under 15 kg. Use rucksack (two-shoulder carry). Heavy items carried separately.", 2, 2, 4, "Weigh kit quarterly"),
        ("Trip hazards in patient home (rugs, cables, pets, clutter)",
         "Practitioner", "Dynamic risk assessment on arrival. Request clear pathway to treatment area. Move obstacles before treating.", 2, 1, 2, "Record at-risk premises"),
        ("Poor lighting in patient home",
         "Practitioner, Patient", "Carry portable examination light for inadequate-lighting environments.", 2, 2, 4, "Check light at each visit"),
        ("Aggressive or unexpected behaviour — patient or household",
         "Practitioner", "Pre-visit screening call for new patients. Flagged patients register maintained. Lone worker check-in system active.", 1, 3, 3, "De-escalation training"),
        ("Aggressive pets",
         "Practitioner", "Ask at booking whether pets are present. Request pets are secured before visit. Add to flagged register.", 2, 2, 4, "Update register after each visit"),
        ("Practitioner medical emergency (no colleague present)",
         "Practitioner", "Lone worker check-in system — emergency contact has schedule and escalation steps. Mobile phone always charged and accessible.", 1, 3, 3, "Annual emergency review"),
    ]
    risk_table(doc, risks_a)

    add_section_heading(doc, "Section B — Vehicle and Waste Transport Hazards")
    add_teal_rule(doc)
    risks_b = [
        ("Clinical waste unsecured in vehicle",
         "Public, Practitioner", "Rigid lockable outer transport box in boot. Sharps in sealed UN3291 container inside box. Never loose bags in vehicle.", 1, 3, 3, "Check before each journey"),
        ("Local anaesthetic temperature excursion (>25°C)",
         "Patient", "Insulated medicines bag in vehicle. Temperature log maintained. Never left in hot vehicle unattended.", 2, 2, 4, "Seasonal check of bag efficacy"),
        ("Road traffic collision during clinical waste transport",
         "Public, Practitioner", "Lower Tier EA waste carrier registered. ADR 1.3 training completed. 2kg dry powder fire extinguisher in vehicle.", 1, 3, 3, "Annual insurance review"),
        ("Contamination of clean supplies",
         "Patient", "Separate labelled containers for clean and contaminated items. Never mixed.", 1, 3, 3, "Check segregation system weekly"),
    ]
    risk_table(doc, risks_b)

    add_section_heading(doc, "Section C — Lone Worker Specific Hazards")
    add_teal_rule(doc)
    risks_c = [
        ("No colleague awareness of location",
         "Practitioner", "Daily movement log shared with named emergency contact before first visit. Check-in after each patient. End-of-day confirmation.", 1, 3, 3, "Test escalation steps annually"),
        ("Delayed check-in response — contact not available",
         "Practitioner", "Emergency contact is a reliable named individual. Alternative emergency contact identified.", 1, 3, 3, "Review contact annually"),
        ("Working in unfamiliar areas / late visits",
         "Practitioner", "No lone visits after dark to unfamiliar addresses. Google Maps pre-checked. Park in visible location.", 2, 2, 4, ""),
    ]
    risk_table(doc, risks_c)

    add_section_heading(doc, "COSHH — Substances used in practice")
    add_teal_rule(doc)
    body_para(doc,
        "A separate COSHH assessment is completed for each hazardous substance using the manufacturer's Safety Data Sheet (SDS). "
        "Substances requiring COSHH assessment in this practice include: "
        "0.5% chlorhexidine in 70% isopropyl alcohol (skin prep); "
        "sodium hypochlorite solution (spill disinfection); "
        "surface disinfectant wipes (Clinell or equivalent); "
        "alcohol-based hand rub; "
        "local anaesthetic agents (lidocaine, etc.).")
    note_para(doc, "COSHH assessments are filed separately. COSHHmate.co.uk is the recommended tool for generating assessments from SDS data.")

    add_section_heading(doc, "Review and signature")
    add_teal_rule(doc)
    body_para(doc,
        "This risk assessment must be reviewed: at least annually; after any incident, near-miss, or RIDDOR-reportable event; "
        "after any significant change to practice, patient group, or working environment.")
    signature_table(doc, [("Practitioner — assessment signature and date",)])

    # Blank review log table
    add_section_heading(doc, "Review log")
    add_teal_rule(doc)
    rl = doc.add_table(rows=0, cols=3)
    rl.style = 'Table Grid'
    rh = rl.add_row()
    for i, h in enumerate(["Review date", "Changes made", "Signature"]):
        c = rh.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    for i in range(5):
        row = rl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '400')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        for j in range(3):
            c = row.cells[j]
            shade_cell(c, hex_str(BG) if i % 2 == 0 else 'FFFFFF')
            cell_border_all(c, color='D4E3E4')
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
    rl.columns[0].width = Cm(3)
    rl.columns[1].width = Cm(10)
    rl.columns[2].width = Cm(3.5)

    footer_line(doc, "Solecare Glasgow  |  H&S Risk Assessment  |  Review annually  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Risk_Assessment.docx"))
    print("  Risk_Assessment.docx — done")


if __name__ == "__main__":
    print("Building policy documents...")
    build_ipc_policy()
    build_risk_assessment()
    print("Policy documents done.")
