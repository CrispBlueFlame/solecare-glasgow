"""
Solecare Glasgow — compliance Excel log builder.
Run with: python3 build_excel_logs.py
Outputs 7 .xlsx files to ./logs/
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, Protection
)
from openpyxl.styles.numbers import FORMAT_DATE_DATETIME
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime, os

OUT = os.path.join(os.path.dirname(__file__), "logs")

# ── Colour palette ──────────────────────────────────────────────────────────
TEAL       = "2E7D83"
DARK       = "1F5A5F"
BG         = "F0F7F7"
BORDER_CLR = "D4E3E4"
MUTED      = "4A4A4A"
TEXT       = "1F1F1F"
WHITE      = "FFFFFF"
YELLOW     = "FFF8DC"   # alert/flag rows

# ── Style helpers ────────────────────────────────────────────────────────────
def fill(hex_):     return PatternFill("solid", fgColor=hex_)
def side(hex_=BORDER_CLR): return Side(style="thin", color=hex_)
def border():       return Border(left=side(), right=side(), top=side(), bottom=side())
def hdr_font():     return Font(name="Calibri", bold=True, color=WHITE, size=11)
def sub_font():     return Font(name="Calibri", bold=True, color=DARK, size=10)
def body_font():    return Font(name="Calibri", color=TEXT, size=10)
def muted_font():   return Font(name="Calibri", color=MUTED, size=9, italic=True)
def title_font():   return Font(name="Calibri Light", bold=True, color=DARK, size=14)

def header_row(ws, row, cols):
    """Write a teal header row. cols = list of (col_index, label, width)."""
    for col, label, width in cols:
        c = ws.cell(row=row, column=col, value=label)
        c.font = hdr_font()
        c.fill = fill(TEAL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border()
        ws.column_dimensions[get_column_letter(col)].width = width

def subheader_row(ws, row, label, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=label)
    c.font = sub_font()
    c.fill = fill(BG)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = border()
    ws.row_dimensions[row].height = 18

def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws.cell(row=1, column=1, value=title)
    t.font = Font(name="Calibri Light", bold=True, color=WHITE, size=15)
    t.fill = fill(DARK)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(row=2, column=1, value=subtitle)
    s.font = Font(name="Calibri", color=MUTED, size=10)
    s.fill = fill(BG)
    s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

def placeholder(ws, row, col, text=""):
    c = ws.cell(row=row, column=col, value=text)
    c.font = muted_font()
    c.fill = fill(WHITE)
    c.border = border()
    c.alignment = Alignment(vertical="center", wrap_text=True)
    return c

def blank_row(ws, row, ncols, shaded=False):
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.fill = fill(BG if shaded else WHITE)
        c.border = border()
        c.font = body_font()

def freeze(ws, cell="A4"):
    ws.freeze_panes = cell

# ════════════════════════════════════════════════════════════════════════════
# 1. FINANCIAL RECORDS
# ════════════════════════════════════════════════════════════════════════════
def build_financial():
    wb = openpyxl.Workbook()

    # ── Income sheet ────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Income"
    title_block(ws, "Solecare Glasgow — Income Record", "Tax year: ___________   |   Sole trader: ___________", 7)
    cols = [
        (1, "Invoice No.", 12), (2, "Date Issued", 13), (3, "Patient Ref.", 16),
        (4, "Service", 22), (5, "Amount £", 12), (6, "Date Paid", 13), (7, "Notes", 22),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")
    for r in range(4, 54):
        blank_row(ws, r, 7, shaded=(r % 2 == 0))
    # Running total row
    tot_row = 54
    ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=4)
    lbl = ws.cell(row=tot_row, column=1, value="TOTAL INCOME")
    lbl.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    lbl.fill = fill(DARK)
    lbl.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    tot = ws.cell(row=tot_row, column=5, value=f"=SUM(E4:E53)")
    tot.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    tot.fill = fill(DARK)
    tot.number_format = "£#,##0.00"
    tot.border = border()
    note = ws.cell(row=tot_row, column=6, value="Auto-sum of column E")
    note.font = muted_font()
    note.fill = fill(DARK)
    for col in [6, 7]:
        ws.cell(row=tot_row, column=col).fill = fill(DARK)

    # ── Expenses sheet ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Expenses")
    title_block(ws2, "Solecare Glasgow — Expenses Record", "Retain all receipts. HMRC requires 5-year retention after filing.", 6)
    cols2 = [
        (1, "Date", 13), (2, "Supplier / Description", 28), (3, "Category", 22),
        (4, "Amount £", 12), (5, "Receipt ref.", 14), (6, "Notes", 22),
    ]
    header_row(ws2, 3, cols2)
    # Category dropdown
    cats = '"HCPC/RCPod fees,Insurance,CPD,Instruments & consumables,PPE,Clinical waste,Business email/software,Vehicle — mileage,Vehicle — equipment,Medicines storage,Lone worker app,Home office,Other"'
    dv = DataValidation(type="list", formula1=cats, showDropDown=False)
    ws2.add_data_validation(dv)
    freeze(ws2, "A4")
    for r in range(4, 54):
        blank_row(ws2, r, 6, shaded=(r % 2 == 0))
        dv.add(ws2.cell(row=r, column=3))

    tot2_row = 54
    ws2.merge_cells(start_row=tot2_row, start_column=1, end_row=tot2_row, end_column=3)
    lbl2 = ws2.cell(row=tot2_row, column=1, value="TOTAL EXPENSES")
    lbl2.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    lbl2.fill = fill(DARK)
    lbl2.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    tot2 = ws2.cell(row=tot2_row, column=4, value="=SUM(D4:D53)")
    tot2.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    tot2.fill = fill(DARK)
    tot2.number_format = "£#,##0.00"
    tot2.border = border()
    ws2.cell(row=tot2_row, column=5).fill = fill(DARK)
    ws2.cell(row=tot2_row, column=6).fill = fill(DARK)

    # ── Mileage sheet ───────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Mileage")
    title_block(ws3, "Solecare Glasgow — Mileage Log", "45p/mile (first 10,000 business miles per tax year), then 25p/mile. Log every journey contemporaneously.", 7)
    cols3 = [
        (1, "Date", 13), (2, "Start point", 20), (3, "Destination / Patient postcode", 28),
        (4, "Purpose", 22), (5, "Miles", 10), (6, "Rate", 10), (7, "Allowable £", 13),
    ]
    header_row(ws3, 3, cols3)
    freeze(ws3, "A4")
    # Running total row at top for quick reference
    ws3.insert_rows(3, 1)  # shift — recalculate after
    # Re-do after insert
    ws3.delete_rows(3, 1)
    header_row(ws3, 3, cols3)
    for r in range(4, 104):
        blank_row(ws3, r, 7, shaded=(r % 2 == 0))
        # Rate formula: =IF(SUM($E$4:E{r})>10000, 0.25, 0.45)
        rate_cell = ws3.cell(row=r, column=6)
        rate_cell.value = f'=IF(SUM($E$4:E{r})>10000,0.25,0.45)'
        rate_cell.font = body_font()
        rate_cell.number_format = "£0.00"
        rate_cell.border = border()
        rate_cell.fill = fill(BG if r % 2 == 0 else WHITE)
        val_cell = ws3.cell(row=r, column=7)
        val_cell.value = f'=IF(E{r}="","",E{r}*F{r})'
        val_cell.font = body_font()
        val_cell.number_format = "£#,##0.00"
        val_cell.border = border()
        val_cell.fill = fill(BG if r % 2 == 0 else WHITE)

    tot3_row = 104
    ws3.merge_cells(start_row=tot3_row, start_column=1, end_row=tot3_row, end_column=4)
    lbl3 = ws3.cell(row=tot3_row, column=1, value="TOTAL MILEAGE ALLOWANCE")
    lbl3.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    lbl3.fill = fill(DARK)
    lbl3.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    tot3_miles = ws3.cell(row=tot3_row, column=5, value="=SUM(E4:E103)")
    tot3_miles.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    tot3_miles.fill = fill(DARK)
    tot3_miles.border = border()
    tot3_val = ws3.cell(row=tot3_row, column=7, value="=SUM(G4:G103)")
    tot3_val.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    tot3_val.fill = fill(DARK)
    tot3_val.number_format = "£#,##0.00"
    tot3_val.border = border()
    for col in [6]:
        ws3.cell(row=tot3_row, column=col).fill = fill(DARK)

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("Summary")
    title_block(ws4, "Solecare Glasgow — Annual Summary", "Review before Self Assessment filing (31 January deadline).", 3)
    rows = [
        ("Total income", "=Income!E54"),
        ("Total expenses", "=Expenses!D54"),
        ("Total mileage allowance", "=Mileage!G104"),
        ("Total deductible costs", "=B6+B7"),
        ("Net profit (taxable)", "=B5-B8"),
    ]
    labels = ["Gross income", "Expenses total", "Mileage allowance",
              "Total deductions", "Net taxable profit"]
    section_labels = [
        (5, "Income"),
        (6, "Deductions"),
        (7, ""),
        (8, ""),
        (9, "Result"),
    ]
    data = [
        (5, "Total income (from Income sheet)",       "=Income!E54"),
        (6, "Total expenses (from Expenses sheet)",   "=Expenses!D54"),
        (7, "Total mileage allowance (from Mileage)","=Mileage!G104"),
        (8, "Total deductions",                       "=B6+B7"),
        (9, "Net taxable profit",                     "=B5-B8"),
    ]
    for r, label, formula in data:
        lc = ws4.cell(row=r, column=1, value=label)
        lc.font = sub_font()
        lc.fill = fill(BG)
        lc.border = border()
        lc.alignment = Alignment(indent=1, vertical="center")
        ws4.row_dimensions[r].height = 20
        vc = ws4.cell(row=r, column=2, value=formula)
        vc.font = Font(name="Calibri", bold=(r == 9), color=DARK, size=11)
        vc.fill = fill(WHITE)
        vc.number_format = "£#,##0.00"
        vc.border = border()
        ws4.column_dimensions["A"].width = 40
        ws4.column_dimensions["B"].width = 18

    note_r = 11
    ws4.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=2)
    n = ws4.cell(row=note_r, column=1,
        value="Note: VAT does not apply — podiatry is VAT-exempt (VAT Act 1994, Schedule 9 Group 7). Do not register for VAT.")
    n.font = muted_font()
    n.fill = fill(YELLOW)
    n.alignment = Alignment(wrap_text=True, indent=1, vertical="center")
    ws4.row_dimensions[note_r].height = 30

    wb.save(os.path.join(OUT, "Financial_Records.xlsx"))
    print("  Financial_Records.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 2. CLINICAL WASTE LOG
# ════════════════════════════════════════════════════════════════════════════
def build_waste_log():
    wb = openpyxl.Workbook()

    # ── Transport log ────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Transport Log"
    title_block(ws, "Solecare Glasgow — Clinical Waste Transport Log",
        "Required under Environmental Protection Act 1990 s.34 duty of care. Retain this record for 3 years.", 8)
    cols = [
        (1, "Date", 12), (2, "Patient postcode(s)", 20), (3, "Sharps (Y/N)", 12),
        (4, "Infectious waste (Y/N)", 16), (5, "Offensive waste (Y/N)", 16),
        (6, "Outer container ref.", 16), (7, "Est. weight (kg)", 14), (8, "Notes", 20),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")
    yn = '"Y,N"'
    for field in ["C", "D", "E"]:
        dv = DataValidation(type="list", formula1=yn)
        ws.add_data_validation(dv)
        for r in range(4, 104):
            dv.add(ws[f"{field}{r}"])
    for r in range(4, 104):
        blank_row(ws, r, 8, shaded=(r % 2 == 0))

    # ── Consignment Notes log ────────────────────────────────────────────────
    ws2 = wb.create_sheet("Consignment Notes")
    title_block(ws2, "Solecare Glasgow — Hazardous Waste Consignment Notes",
        "One row per collection. Retain original signed note for 3 years (sharps + infectious). Waste Transfer Notes for offensive waste: 2 years.", 7)
    cols2 = [
        (1, "Collection date", 14), (2, "Carrier name", 22), (3, "Carrier reg. (CBDU...)", 20),
        (4, "Waste type", 18), (5, "Consignment/WT note ref.", 22), (6, "Signed (Y/N)", 12), (7, "Notes", 20),
    ]
    header_row(ws2, 3, cols2)
    freeze(ws2, "A4")
    wtype = '"Sharps (HW consignment note),Infectious soft waste (HW consignment note),Offensive/tiger-stripe (Waste Transfer Note)"'
    dv2 = DataValidation(type="list", formula1=wtype)
    ws2.add_data_validation(dv2)
    dv3 = DataValidation(type="list", formula1=yn)
    ws2.add_data_validation(dv3)
    for r in range(4, 104):
        blank_row(ws2, r, 7, shaded=(r % 2 == 0))
        dv2.add(ws2[f"D{r}"])
        dv3.add(ws2[f"F{r}"])

    # ── Carrier register ─────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Carrier Register")
    title_block(ws3, "Solecare Glasgow — Licensed Carrier Register",
        "Verify each carrier's EA registration at environment.data.gov.uk/public-register before first use. Re-verify annually.", 5)
    cols3 = [
        (1, "Carrier name", 26), (2, "EA reg. number (CBDU...)", 22),
        (3, "Services", 24), (4, "Last verified date", 18), (5, "Notes", 22),
    ]
    header_row(ws3, 3, cols3)
    for r in range(4, 14):
        blank_row(ws3, r, 5, shaded=(r % 2 == 0))

    wb.save(os.path.join(OUT, "Clinical_Waste_Log.xlsx"))
    print("  Clinical_Waste_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 3. DAILY MOVEMENT LOG  (lone worker)
# ════════════════════════════════════════════════════════════════════════════
def build_movement_log():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Movement Log Template"

    title_block(ws, "Solecare Glasgow — Daily Movement Log",
        "Complete before first visit each day. Share with emergency contact. HSE INDG73 lone worker requirement.", 8)

    # Day info block
    info_rows = [
        (3, "Date:", ""),
        (4, "Emergency contact name:", ""),
        (5, "Emergency contact phone:", ""),
        (6, "Vehicle reg.:", ""),
    ]
    for r, label, val in info_rows:
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = sub_font()
        lc.fill = fill(BG)
        lc.border = border()
        lc.alignment = Alignment(indent=1, vertical="center")
        ws.row_dimensions[r].height = 18
        vc = ws.cell(row=r, column=2, value=val)
        vc.font = body_font()
        vc.fill = fill(WHITE)
        vc.border = border()
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)

    # Column headers
    cols = [
        (1, "Visit #", 8), (2, "Patient ref.", 16), (3, "Address / postcode", 26),
        (4, "Appt time", 12), (5, "Expected finish", 14), (6, "Check-in time", 13),
        (7, "Check-out time", 13), (8, "Notes / incidents", 24),
    ]
    header_row(ws, 7, cols)
    freeze(ws, "A8")

    for r in range(8, 18):
        blank_row(ws, r, 8, shaded=(r % 2 == 0))
        ws.row_dimensions[r].height = 20

    # End-of-day row
    end_r = 18
    ws.merge_cells(start_row=end_r, start_column=1, end_row=end_r, end_column=5)
    ed = ws.cell(row=end_r, column=1, value="End-of-day confirmation sent to emergency contact (time):")
    ed.font = sub_font()
    ed.fill = fill(BG)
    ed.border = border()
    ed.alignment = Alignment(indent=1, vertical="center")
    ec = ws.cell(row=end_r, column=6, value="")
    ec.font = body_font()
    ec.fill = fill(WHITE)
    ec.border = border()
    ws.merge_cells(start_row=end_r, start_column=6, end_row=end_r, end_column=8)
    ws.row_dimensions[end_r].height = 22

    # Escalation procedure note
    note_r = 20
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=8)
    note = ws.cell(row=note_r, column=1,
        value="ESCALATION: If check-in is missed — emergency contact calls practitioner mobile. If no answer within 15 minutes, calls last known patient address. If still no answer within further 15 minutes, calls 999 and provides this log.")
    note.font = Font(name="Calibri", color="7C2820", size=10, bold=True)
    note.fill = fill("FDECEA")
    note.alignment = Alignment(wrap_text=True, indent=1, vertical="center")
    ws.row_dimensions[note_r].height = 40

    ws.column_dimensions["A"].width = 8
    ws.print_area = "A1:H22"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"

    wb.save(os.path.join(OUT, "Daily_Movement_Log.xlsx"))
    print("  Daily_Movement_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 4. IPC VISIT LOG
# ════════════════════════════════════════════════════════════════════════════
def build_ipc_log():
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "IPC Visit Log"
    title_block(ws, "Solecare Glasgow — Infection Control Visit Log",
        "Complete per patient visit. Supports HCPC Standards 6.1–6.2 and RCPod Clinical Standards 5 and 9.", 9)
    cols = [
        (1, "Date", 12), (2, "Patient ref.", 14), (3, "Gloves (Y/N)", 12),
        (4, "Apron (Y/N)", 10), (5, "Mask (Y/N)", 10), (6, "Clean field used", 12),
        (7, "Spill event", 10), (8, "Waste bagged & sealed", 14), (9, "Notes", 22),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")
    yn = '"Y,N"'
    for col_letter in ["C", "D", "E", "F", "G", "H"]:
        dv = DataValidation(type="list", formula1=yn)
        ws.add_data_validation(dv)
        for r in range(4, 104):
            dv.add(ws[f"{col_letter}{r}"])
    for r in range(4, 104):
        blank_row(ws, r, 9, shaded=(r % 2 == 0))
        ws.row_dimensions[r].height = 18

    ws2 = wb.create_sheet("Sharps Container Log")
    title_block(ws2, "Solecare Glasgow — Sharps Container Log",
        "Assemble a new container, log it. Close and dispose when three-quarters full. Never overfill.", 6)
    cols2 = [
        (1, "Container ref.", 16), (2, "Date assembled", 15), (3, "Date sealed / closed", 17),
        (4, "Fill level at sealing", 18), (5, "Waste Transfer/Consignment note ref.", 30), (6, "Notes", 22),
    ]
    header_row(ws2, 3, cols2)
    fill_level = '"<25%,25-50%,50-75% (three-quarters — seal now)"'
    dv4 = DataValidation(type="list", formula1=fill_level)
    ws2.add_data_validation(dv4)
    for r in range(4, 54):
        blank_row(ws2, r, 6, shaded=(r % 2 == 0))
        dv4.add(ws2[f"D{r}"])
        ws2.row_dimensions[r].height = 18

    wb.save(os.path.join(OUT, "IPC_Visit_Log.xlsx"))
    print("  IPC_Visit_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 5. MEDICINES LOG
# ════════════════════════════════════════════════════════════════════════════
def build_medicines_log():
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Administration Log"
    title_block(ws, "Solecare Glasgow — Medicines Administration Log",
        "Record every administration of local anaesthetic. HCPC requirement. Retain for 8 years (adult patient records).", 9)
    cols = [
        (1, "Date", 12), (2, "Patient ref.", 16), (3, "Medicine name", 20),
        (4, "Batch number", 14), (5, "Expiry date", 12), (6, "Dose / volume", 14),
        (7, "Site", 16), (8, "Route", 12), (9, "Practitioner initials", 16),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")
    route_dv = DataValidation(type="list", formula1='"Subcutaneous,Intra-dermal,Ring block,Digital block,Other"')
    ws.add_data_validation(route_dv)
    for r in range(4, 104):
        blank_row(ws, r, 9, shaded=(r % 2 == 0))
        route_dv.add(ws[f"H{r}"])
        ws.row_dimensions[r].height = 18

    ws2 = wb.create_sheet("Temperature Log")
    title_block(ws2, "Solecare Glasgow — Medicines Temperature Log",
        "Local anaesthetics must be stored below 25°C (per SmPC). Log any excursion above 25°C. Notify dispensing pharmacy if product may be compromised.", 7)
    cols2 = [
        (1, "Date", 12), (2, "Time", 10), (3, "Storage location", 22),
        (4, "Temperature (°C)", 16), (5, "Above 25°C? (Y/N)", 16), (6, "Action taken", 24), (7, "Initials", 12),
    ]
    header_row(ws2, 3, cols2)
    yn = '"Y,N"'
    dv5 = DataValidation(type="list", formula1=yn)
    ws2.add_data_validation(dv5)
    for r in range(4, 104):
        blank_row(ws2, r, 7, shaded=(r % 2 == 0))
        dv5.add(ws2[f"E{r}"])
        # Conditional fill for excursion rows is best done in Excel directly
        ws2.row_dimensions[r].height = 18

    ws3 = wb.create_sheet("Stock Register")
    title_block(ws3, "Solecare Glasgow — Medicines Stock Register",
        "Log each purchase and use. Retain supplier invoices for 2 years.", 7)
    cols3 = [
        (1, "Date", 12), (2, "Medicine / product", 24), (3, "Batch no.", 14),
        (4, "Expiry", 12), (5, "Qty received", 13), (6, "Qty used", 12), (7, "Supplier invoice ref.", 20),
    ]
    header_row(ws3, 3, cols3)
    for r in range(4, 54):
        blank_row(ws3, r, 7, shaded=(r % 2 == 0))
        ws3.row_dimensions[r].height = 18

    wb.save(os.path.join(OUT, "Medicines_Log.xlsx"))
    print("  Medicines_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 6. INCIDENT LOG
# ════════════════════════════════════════════════════════════════════════════
def build_incident_log():
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Incident Log"
    title_block(ws, "Solecare Glasgow — Incident and Near-Miss Log",
        "Record all incidents and near-misses. RIDDOR-notifiable events must be reported to HSE within 15 days.", 9)
    cols = [
        (1, "Date", 12), (2, "Type", 22), (3, "Location / patient ref.", 20),
        (4, "Description", 30), (5, "Immediate action taken", 26),
        (6, "RIDDOR reportable?", 16), (7, "HSE report ref.", 16),
        (8, "Follow-up action", 22), (9, "Closed date", 12),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")

    type_dv = DataValidation(type="list",
        formula1='"Sharps injury,Blood/body fluid exposure,Patient adverse event,Slip/trip/fall,Aggression/verbal abuse,Data breach,Medicine error,Vehicle incident,Equipment failure,Near-miss,Other"')
    ws.add_data_validation(type_dv)
    yn = '"Y — report to HSE,N"'
    riddor_dv = DataValidation(type="list", formula1=yn)
    ws.add_data_validation(riddor_dv)
    for r in range(4, 54):
        blank_row(ws, r, 9, shaded=(r % 2 == 0))
        type_dv.add(ws[f"B{r}"])
        riddor_dv.add(ws[f"F{r}"])
        ws.row_dimensions[r].height = 20

    # Sharps injury protocol reminder
    note_r = 55
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=9)
    note = ws.cell(row=note_r, column=1,
        value="SHARPS INJURY: Wash site immediately under running water (do not suck). Encourage bleeding. Cover with waterproof dressing. Record above immediately. Seek blood-borne virus (BBV) advice from GP or A&E within 1 hour. Notify PI insurer if relevant.")
    note.font = Font(name="Calibri", color="7C2820", size=10, bold=True)
    note.fill = fill("FDECEA")
    note.alignment = Alignment(wrap_text=True, indent=1, vertical="center")
    ws.row_dimensions[note_r].height = 50

    ws2 = wb.create_sheet("Sharps Injury Detail")
    title_block(ws2, "Solecare Glasgow — Sharps Injury Detailed Record",
        "Complete immediately after a sharps injury in addition to the Incident Log. This record stays confidential.", 6)
    detail_rows = [
        ("Date and time of injury:", ""),
        ("Name of injured person:", ""),
        ("Type of device involved:", ""),
        ("Was the device visibly contaminated?", "Y / N"),
        ("Patient whose blood was on device (if known):", ""),
        ("What were you doing when injured?", ""),
        ("Immediate first aid taken:", ""),
        ("Time GP / A&E attended:", ""),
        ("Advice given / BBV risk assessment outcome:", ""),
        ("Follow-up blood tests required?", "Y / N / Dates:"),
        ("Notified PI insurer?", "Y / N / Date:"),
        ("Risk assessment updated?", "Y / N / Date:"),
    ]
    for i, (label, val) in enumerate(detail_rows):
        r = i + 4
        lc = ws2.cell(row=r, column=1, value=label)
        lc.font = sub_font()
        lc.fill = fill(BG)
        lc.border = border()
        lc.alignment = Alignment(indent=1, vertical="center")
        ws2.row_dimensions[r].height = 22
        vc = ws2.cell(row=r, column=2, value=val)
        vc.font = body_font()
        vc.fill = fill(WHITE)
        vc.border = border()
        vc.alignment = Alignment(vertical="center", wrap_text=True)
        ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 40

    wb.save(os.path.join(OUT, "Incident_Log.xlsx"))
    print("  Incident_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# 7. CPD LOG
# ════════════════════════════════════════════════════════════════════════════
def build_cpd_log():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CPD Log"
    title_block(ws, "Solecare Glasgow — CPD Record",
        "HCPC requires a maintained CPD profile. No minimum hours — qualitative evidence of impact. Start from registration date.", 8)
    cols = [
        (1, "Date", 12), (2, "Activity / title", 30), (3, "CPD type", 22),
        (4, "Provider / source", 22), (5, "Hours", 8),
        (6, "HCPC standard(s) met", 22), (7, "What I learned", 30), (8, "How this benefits my practice", 30),
    ]
    header_row(ws, 3, cols)
    freeze(ws, "A4")
    cpd_type = '"Course/workshop,Self-directed study,Peer review,Clinical audit,Conference,Reflective practice,Online learning,Journal reading,Other"'
    dv = DataValidation(type="list", formula1=cpd_type)
    ws.add_data_validation(dv)
    stds = '"Standard 1 (professional autonomy),Standard 2 (communication),Standard 3 (safety),Standard 4 (clinical),Standard 5 (infection control),Standard 6 (risk management),Multiple"'
    dv2 = DataValidation(type="list", formula1=stds)
    ws.add_data_validation(dv2)
    for r in range(4, 54):
        blank_row(ws, r, 8, shaded=(r % 2 == 0))
        dv.add(ws[f"C{r}"])
        dv2.add(ws[f"F{r}"])
        ws.row_dimensions[r].height = 24

    # Running total hours
    tot_r = 54
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=4)
    lbl = ws.cell(row=tot_r, column=1, value="TOTAL CPD HOURS (this log)")
    lbl.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    lbl.fill = fill(DARK)
    lbl.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    tot = ws.cell(row=tot_r, column=5, value="=SUM(E4:E53)")
    tot.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
    tot.fill = fill(DARK)
    tot.border = border()
    for col in range(6, 9):
        ws.cell(row=tot_r, column=col).fill = fill(DARK)

    note_r = 56
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=8)
    n = ws.cell(row=note_r, column=1,
        value="HCPC audit tip: Select 5–6 entries that best show development and impact. Write 500–1,000 words of reflective summary. HCPC do not count hours — they assess quality of reflection.")
    n.font = muted_font()
    n.fill = fill(YELLOW)
    n.alignment = Alignment(wrap_text=True, indent=1, vertical="center")
    ws.row_dimensions[note_r].height = 36

    wb.save(os.path.join(OUT, "CPD_Log.xlsx"))
    print("  CPD_Log.xlsx — done")


# ════════════════════════════════════════════════════════════════════════════
# RUN ALL
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("Building Excel logs...")
    build_financial()
    build_waste_log()
    build_movement_log()
    build_ipc_log()
    build_medicines_log()
    build_incident_log()
    build_cpd_log()
    print("All Excel logs built.")
