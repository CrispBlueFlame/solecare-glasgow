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

def risk_table(doc, risks, widths=None):
    """
    risks: list of (hazard, who, controls, likelihood, severity, risk_level, additional)
    widths: optional list of 7 Cm() values — defaults to portrait; pass landscape widths when building that doc.
    """
    if widths is None:
        widths = [Cm(3), Cm(2.2), Cm(4.5), Cm(1), Cm(1), Cm(1), Cm(3.8)]
    tbl = doc.add_table(rows=0, cols=7)
    tbl.style = 'Table Grid'
    hdr = tbl.add_row()
    headers = ['Hazard', 'Who at risk', 'Existing controls', 'L\n(1-3)', 'S\n(1-3)', 'Risk\n(LxS)', 'Additional action / review']
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
# HEALTH AND SAFETY RISK ASSESSMENT  (A4 landscape)
# ════════════════════════════════════════════════════════════════════════════
def build_risk_assessment():
    doc = a4_doc()

    # ── Override to A4 landscape ─────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width   = Cm(29.7)
    section.page_height  = Cm(21)
    section.left_margin  = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin   = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    sectPr = section._sectPr
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        sectPr.append(pgSz)
    pgSz.set(qn('w:orient'), 'landscape')

    # Usable width: 29.7 - 3 = 26.7 cm
    # Risk table column widths summing to 26.7 cm:
    LS_WIDTHS = [Cm(4.5), Cm(2.8), Cm(10.5), Cm(1.1), Cm(1.1), Cm(1.1), Cm(5.6)]

    add_doc_title(doc, "Solecare Glasgow — Health and Safety Risk Assessment",
        "HSWA 1974  |  MHSWR 1999  |  COSHH 2002  |  Manual Handling Operations Regulations 1992  |  HSE INDG73  |  Sharps Regulations 2013")

    notice_box(doc,
        "Risk scoring  —  Likelihood (L): 1 = unlikely, 2 = possible, 3 = likely  |  "
        "Severity (S): 1 = minor injury/near-miss, 2 = moderate harm, 3 = serious harm or death  |  "
        "Risk rating = L × S  |  Score 1-2 = LOW (acceptable)  |  3-4 = MEDIUM (monitor, review controls)  |  6-9 = HIGH (immediate action required)")

    form_table(doc, [
        ("Practitioner name:", ""),
        ("HCPC registration no.:", ""),
        ("Practice type:", "Mobile sole-trader, domiciliary house calls, single-use instruments only"),
        ("Assessment date:", ""),
        ("Next review date:", "Annual minimum, or after any incident, near-miss, or significant change"),
    ])

    # ── Section A ─────────────────────────────────────────────────────────────
    add_section_heading(doc, "Section A — Clinical and Domiciliary Environment Hazards")
    add_teal_rule(doc)
    risks_a = [
        ("Sharps injury — scalpel blades, needles, lancets",
         "Practitioner",
         "Single-use instruments only. Dispose directly into UN3291-compliant yellow-lidded sharps bin at point of use — never set down first. "
         "Never resheath needles. Never pass sharps between hands. Carry portable sharps bin to every visit. "
         "Refer to written Sharps Protocol (Sharps_Protocol.docx).",
         1, 3, 3,
         "Annual review of sharps protocol. Hep B immunity confirmed on file. Post-injury BBV protocol documented."),

        ("Blood or body fluid exposure — splash or contact",
         "Practitioner",
         "Gloves and apron worn at all times. Type IIR mask and eye protection added for nail surgery and any splash-risk procedure. "
         "Carry spill kit to every visit (absorbent granules, 10,000 ppm hypochlorite wipes, face shield, heavy gloves, clinical waste bag). "
         "Post-exposure protocol: wash immediately, encourage bleeding, cover, seek BBV advice within 1 hour.",
         1, 3, 3,
         "BBV immunisation on file. Incident_Log.xlsx completed for each event."),

        ("Cross-infection between patients",
         "Patient, Practitioner",
         "Standard Infection Control Precautions (SICPs) applied at every visit per NHS NIPCM v2.10. "
         "Single-use sterile instruments from licensed medical device supplier. Disposable paper drape clean field before each treatment. "
         "PPE changed between patients. Pre-visit screening call for any reported household illness.",
         1, 3, 3,
         "Annual IPC policy review and re-signature. IPC CPD logged annually."),

        ("Infection risk — nail surgery in domestic setting",
         "Patient, Practitioner",
         "0.5% chlorhexidine in 70% isopropyl alcohol (ChloraPrep or equivalent) applied to surgical site per NICE NG125. "
         "Minimum 30 seconds air-dry before incision — do not blot. Sterile drape applied. Adequate lighting confirmed before proceeding. "
         "Do not perform nail surgery in environments where a sterile field cannot be maintained.",
         2, 3, 6,
         "Document pre-surgery environment assessment for each home surgery episode. Check PI insurer covers domiciliary nail surgery."),

        ("Unstable or unsafe working posture — patient seating",
         "Practitioner",
         "Portable folding stool and adjustable footrest carried to every visit. "
         "Practitioner never treats from the floor. "
         "If suitable seating cannot be established, reschedule and note in patient record.",
         2, 2, 4,
         "Annual MSK self-assessment. Ergonomics reviewed at each new patient address."),

        ("Manual handling — transporting equipment to and from vehicle",
         "Practitioner",
         "Total kit weight target: under 15 kg. Rucksack used for even two-shoulder load distribution. "
         "Heavy items (sharps waste box, folding stool) carried separately and not combined with clinical bag. "
         "Practitioner does not carry kit and climb stairs simultaneously.",
         2, 2, 4,
         "Weigh kit quarterly. Reassess if new items added."),

        ("Trip or slip hazards in patient's home (rugs, cables, pets, clutter, wet floors)",
         "Practitioner",
         "Dynamic risk assessment conducted on arrival at every home visit before entering treatment area. "
         "Request clear pathway from door to treatment area at booking. "
         "Move portable trip hazards (mats, cables) before beginning treatment. "
         "Wet floor — ask patient to dry before proceeding.",
         2, 2, 4,
         "Flagged-premises register maintained for repeat-visit hazards."),

        ("Poor or insufficient lighting in patient's home",
         "Practitioner, Patient",
         "Portable LED examination light (rechargeable) carried to all visits. "
         "Do not perform nail surgery if lighting is inadequate and cannot be supplemented.",
         2, 2, 4,
         "Check battery charge before each working day."),

        ("Aggressive, threatening, or unexpected behaviour — patient or household member",
         "Practitioner",
         "Pre-visit telephone screening call for all new patients. "
         "Lone worker check-in system active — emergency contact has schedule and knows escalation steps. "
         "Flagged patients register updated after any concerning behaviour. "
         "Exit strategy: practitioner seats self nearest door, keeps kit bag between self and patient.",
         1, 3, 3,
         "Conflict de-escalation training (Suzy Lamplugh Trust or equivalent). Annual review of flagged register."),

        ("Aggressive or uncontrolled pets",
         "Practitioner",
         "Ask at booking: 'Do you have pets?' Request all pets are secured before practitioner enters. "
         "If pets cannot be secured on arrival, reschedule appointment. Add to flagged premises register.",
         2, 2, 4,
         "Flagged premises register updated at each visit."),

        ("Practitioner medical emergency with no colleague present",
         "Practitioner",
         "Lone worker check-in system operational every working day. "
         "Emergency contact has current-day movement log including all patient addresses, appointment times, and escalation steps. "
         "Mobile phone fully charged and accessible throughout each visit.",
         1, 3, 3,
         "Test escalation procedure annually. Update emergency contact if circumstances change."),
    ]
    risk_table(doc, risks_a, LS_WIDTHS)

    # ── Section B ─────────────────────────────────────────────────────────────
    add_section_heading(doc, "Section B — Vehicle, Medicines, and Clinical Waste Transport Hazards")
    add_teal_rule(doc)
    risks_b = [
        ("Clinical waste unsecured or unsealed in vehicle",
         "Public, Practitioner",
         "All sealed sharps bins and clinical waste bags carried inside a rigid, lockable, leak-proof outer transport box in vehicle boot. "
         "Loose bags directly in boot are prohibited under ADR and EA duty of care. "
         "Transport box locked when vehicle is unattended.",
         1, 3, 3,
         "Check transport box integrity before each journey. ADR 1.3 training certificate held on file."),

        ("Local anaesthetic temperature excursion above 25°C (SmPC limit)",
         "Patient",
         "Local anaesthetics transported in a validated insulated medicines bag at all times during vehicle journeys. "
         "Bag is never left in an unshaded vehicle in warm weather. "
         "Temperature log maintained in Medicines_Log.xlsx. Any product exposed to >25°C for extended periods is not used.",
         2, 2, 4,
         "Insulated bag efficacy re-assessed each spring/summer. Any temperature excursion logged and pharmacist consulted."),

        ("Road traffic collision while transporting clinical waste",
         "Public, Practitioner",
         "Lower Tier Environment Agency waste carrier registration held (wastecarriersregistration.service.gov.uk). "
         "ADR 1.3 general awareness training completed and certificated. "
         "2 kg dry powder fire extinguisher carried in vehicle whenever clinical waste is transported. "
         "Car insurance confirmed as Class 3 business use with clinical waste transport declared.",
         1, 3, 3,
         "Annual insurance renewal with clinical waste transport re-declared. Fire extinguisher serviced annually."),

        ("Cross-contamination of clean supplies by used/contaminated items",
         "Patient",
         "Strict clean/dirty segregation in vehicle: dedicated lidded box for clean sterile supplies; "
         "separate sealed container for used/contaminated items. Containers labelled clearly. Never mixed.",
         1, 2, 2,
         "Segregation system checked before each working day."),

        ("Theft or loss of medicines from unattended vehicle",
         "Public",
         "Medicines are not left in an unattended vehicle. Lockable insulated bag transferred to practitioner's person or secured premises when vehicle is left. "
         "Local anaesthetics are not controlled drugs but are Prescription Only Medicines — access must be controlled.",
         1, 2, 2,
         "Check vehicle is locked and medicines removed at each stop."),
    ]
    risk_table(doc, risks_b, LS_WIDTHS)

    # ── Section C ─────────────────────────────────────────────────────────────
    add_section_heading(doc, "Section C — Lone Worker and Personal Safety Hazards  (HSE INDG73)")
    add_teal_rule(doc)
    risks_c = [
        ("No colleague awareness of practitioner location throughout the working day",
         "Practitioner",
         "Daily movement log (Daily_Movement_Log.xlsx) completed and shared with named emergency contact before first visit each day. "
         "Includes all patient names/addresses, appointment times, and expected finish times. "
         "Check-in message sent after each patient. End-of-day 'home safe' confirmation sent.",
         1, 3, 3,
         "Emergency contact briefed on escalation steps (Lone_Worker_Emergency.docx). Test procedure annually."),

        ("Emergency contact unavailable or fails to respond to missed check-in",
         "Practitioner",
         "Primary emergency contact identified. Secondary emergency contact (alternative named person) identified as fallback. "
         "Escalation steps define time intervals for each action, ending in 999 call if location unconfirmed.",
         1, 3, 3,
         "Review and update emergency contact details annually or if circumstances change."),

        ("Working in unfamiliar areas, isolated addresses, or after dark",
         "Practitioner",
         "Route pre-planned using Google Maps before departure. "
         "Appointments in unfamiliar or isolated areas scheduled during daylight hours only. "
         "Vehicle parked in visible, well-lit location. Practitioner ID visible.",
         2, 2, 4,
         "Review visit scheduling policy annually."),

        ("Lone working without medical support if practitioner becomes ill",
         "Practitioner",
         "Lone worker app (Lookout Call or MyTeamSafe) active throughout each working day — provides automated escalation if check-in is missed. "
         "GP contact details on person. Medical emergency: call 999 if able, or trigger lone worker app alert.",
         1, 3, 3,
         "App subscription active and tested monthly."),
    ]
    risk_table(doc, risks_c, LS_WIDTHS)

    # ── COSHH ─────────────────────────────────────────────────────────────────
    add_section_heading(doc, "COSHH — Hazardous Substances Register  (COSHH Regulations 2002)")
    add_teal_rule(doc)
    body_para(doc,
        "A COSHH assessment is required for each hazardous substance used in this practice. "
        "Obtain the Safety Data Sheet (SDS) from the manufacturer for each substance and assess: hazard, exposure route, control measures, PPE, and emergency action. "
        "Use COSHHmate.co.uk (~£50/year) to auto-populate assessments from SDS data. Assessments are filed separately and reviewed annually.")

    # COSHH substances table
    coshh_rows = [
        ("0.5% chlorhexidine gluconate in 70% isopropyl alcohol", "Nail surgery skin prep", "Flammable; eye/skin irritant", "Allow full air-dry before incision. Gloves. No ignition sources near application site."),
        ("Sodium hypochlorite solution (10,000 ppm)", "Blood spill disinfection on hard surfaces", "Corrosive; harmful vapour; reacts with urine and acids", "Gloves, apron, face shield. Do not use on soft furnishings or urine spills. Ventilate area."),
        ("Sodium hypochlorite solution (1,000 ppm) or detergent", "Spill management on soft furnishings / urine spills", "Low irritant at this concentration", "Gloves. Do not mix with higher concentration hypochlorite."),
        ("Surface disinfectant wipes (e.g. Clinell Universal)", "General surface decontamination", "Skin/eye irritant", "Gloves. Allow contact time per manufacturer. Dispose as clinical waste."),
        ("Alcohol-based hand rub (≥60% ethanol)", "Hand hygiene", "Flammable; drying to skin", "Store away from heat sources. Allow full evaporation before any open flame. Moisturise regularly."),
        ("Lidocaine HCl injection (e.g. 1% or 2% solution)", "Local anaesthesia for nail surgery", "Systemic toxicity risk if overdosed or intravascular", "Check max dose per body weight. Aspirate before injection. Do not exceed single-procedure limits. Sharps disposal immediately post-use."),
    ]
    ctbl = doc.add_table(rows=0, cols=4)
    ctbl.style = 'Table Grid'
    ch = ctbl.add_row()
    for i, h in enumerate(["Substance", "Used for", "Hazard summary", "Controls and PPE"]):
        c = ch.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    cwidths = [Cm(5.5), Cm(4), Cm(5.5), Cm(11.7)]
    for i, row_data in enumerate(coshh_rows):
        row = ctbl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '400')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        bg = hex_str(BG) if i % 2 == 0 else 'FFFFFF'
        for j, val in enumerate(row_data):
            c = row.cells[j]
            shade_cell(c, bg)
            cell_border_all(c, color='D4E3E4')
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
            p = c.paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Calibri'; r.font.size = Pt(8); r.font.color.rgb = TEXT_C
    for i, w in enumerate(cwidths):
        ctbl.columns[i].width = w
    doc.add_paragraph()
    note_para(doc, "SDS documents held on file for each substance above. COSHH assessments filed separately at [location]. Review annually or when product changes.")

    # ── Review and signature ──────────────────────────────────────────────────
    add_section_heading(doc, "Signature and Review Log")
    add_teal_rule(doc)
    body_para(doc,
        "Review triggers: annually as a minimum; after any RIDDOR-reportable event, near-miss, or significant incident; "
        "after any change to working environment, patient group, or practice scope.")
    signature_table(doc, [("Practitioner — initial assessment signature and date",)])

    rl = doc.add_table(rows=0, cols=4)
    rl.style = 'Table Grid'
    rh = rl.add_row()
    for i, h in enumerate(["Review date", "Trigger for review", "Changes made to assessment", "Signature"]):
        c = rh.cells[i]
        shade_cell(c, hex_str(TEAL))
        cell_border_all(c)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE
    for i in range(6):
        row = rl.add_row()
        trPr = row._tr.get_or_add_trPr()
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '440')
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)
        for j in range(4):
            c = row.cells[j]
            shade_cell(c, hex_str(BG) if i % 2 == 0 else 'FFFFFF')
            cell_border_all(c, color='D4E3E4')
            set_cell_margins(c, top=60, bottom=60, left=100, right=100)
    rl.columns[0].width = Cm(2.5)
    rl.columns[1].width = Cm(4.5)
    rl.columns[2].width = Cm(16)
    rl.columns[3].width = Cm(3.7)

    footer_line(doc, "Solecare Glasgow  |  H&S Risk Assessment  |  A4 Landscape  |  Review annually or after any incident  |  Version 1.0  |  September 2026")
    doc.save(os.path.join(OUT, "Risk_Assessment.docx"))
    print("  Risk_Assessment.docx — done")


if __name__ == "__main__":
    print("Building policy documents...")
    build_ipc_policy()
    build_risk_assessment()
    print("Policy documents done.")
