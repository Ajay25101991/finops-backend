import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (Font, Alignment, PatternFill, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter

import os

# ================= PATHS =================
TB_FILE  = r"C:\Users\Ajay Rathod\Desktop\Ajay\Projects\02 Financial Statements\TB_Demo_With_Opening_and_Closing_BALANCED.xlsx"
MAP_FILE = r"C:\Users\Ajay Rathod\Desktop\Ajay\Projects\02 Financial Statements\Mapping_From_TB_Balanced.xlsx"
OUT_FILE = r"C:\Users\Ajay Rathod\Desktop\Ajay\Projects\02 Financial Statements\Automated_FS_v7_Final.xlsx"

# ================= COMPANY CONFIG =================
COMPANY   = "DemoCo Pvt. Ltd."
PERIOD    = "Year Ended 31 December 2025"
CURRENCY  = "INR"
UNIT      = "(Amount in INR)"

# ================= COLORS =================
C_NAVY        = "1F3864"   # Title bar
C_DARK_BLUE   = "2F5496"   # Section headers
C_MID_BLUE    = "4472C4"   # Sub-headers
C_LIGHT_BLUE  = "D6E4F0"   # Column headers
C_PALE_BLUE   = "EBF3FB"   # Alternate row
C_WHITE       = "FFFFFF"
C_GOLD        = "C9A84C"   # Totals highlight
C_GOLD_LIGHT  = "FFF2CC"   # Totals background
C_GREEN_LIGHT = "E2EFDA"   # Positive amount bg
C_GRAY_LIGHT  = "F5F5F5"   # Notes header
C_ORANGE      = "C55A11"   # Warning / contra
C_RED_LIGHT   = "FCE4D6"

# ================= FONTS =================
FT_TITLE      = Font(name="Calibri", bold=True, size=16, color=C_WHITE)
FT_SUBTITLE   = Font(name="Calibri", bold=False, size=10, color=C_WHITE)
FT_COMPANY    = Font(name="Calibri", bold=True, size=18, color=C_WHITE)
FT_HEAD       = Font(name="Calibri", bold=True, size=10, color=C_WHITE)
FT_GROUP      = Font(name="Calibri", bold=True, size=10, color=C_WHITE)
FT_SUBGROUP   = Font(name="Calibri", bold=True, size=10, color="1F3864")
FT_NORMAL     = Font(name="Calibri", size=10)
FT_TOTAL      = Font(name="Calibri", bold=True, size=10, color="1F3864")
FT_GRAND      = Font(name="Calibri", bold=True, size=11, color=C_WHITE)
FT_NOTE_NUM   = Font(name="Calibri", bold=True, size=10, color="2F5496")
FT_ITALIC     = Font(name="Calibri", italic=True, size=9, color="595959")

# ================= FILLS =================
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

# ================= BORDERS =================
thin  = Side(style="thin",   color="BDD7EE")
thick = Side(style="medium", color="2F5496")

def border_bottom_thick():
    return Border(bottom=thick)

def border_thin_all():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def border_bottom_thin():
    return Border(bottom=thin)

def border_top_thick():
    return Border(top=thick)

# ================= NUMBER FORMAT =================
FMT_NUM = '#,##0;(#,##0);"-"'

# ================= HELPERS =================
def autofit(ws, min_width=12, max_width=50):
    for col in ws.columns:
        col_letter = None
        for cell in col:
            if cell.__class__.__name__ != "MergedCell":
                col_letter = cell.column_letter
                break
        if not col_letter:
            continue
        max_len = 0
        for cell in col:
            try:
                if cell.value and cell.__class__.__name__ != "MergedCell":
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        width = max(min_width, min(max_len + 3, max_width))
        ws.column_dimensions[col_letter].width = width


def set_cell(ws, row, col, value=None, font=None, fill_=None,
             align=None, border=None, num_fmt=None, row_height=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:    c.font      = font
    if fill_:   c.fill      = fill_
    if align:   c.alignment = align
    if border:  c.border    = border
    if num_fmt: c.number_format = num_fmt
    if row_height:
        ws.row_dimensions[row].height = row_height
    return c


def title_block(ws, title, subtitle, col_span, start_col=1):
    # Row 1 - Company
    ws.merge_cells(start_row=1, start_column=start_col,
                   end_row=1, end_column=start_col + col_span - 1)
    c = ws.cell(row=1, column=start_col, value=COMPANY)
    c.font = FT_COMPANY
    c.fill = fill(C_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Row 2 - Statement title
    ws.merge_cells(start_row=2, start_column=start_col,
                   end_row=2, end_column=start_col + col_span - 1)
    c = ws.cell(row=2, column=start_col, value=title)
    c.font = FT_TITLE
    c.fill = fill(C_DARK_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    # Row 3 - Period + Currency
    ws.merge_cells(start_row=3, start_column=start_col,
                   end_row=3, end_column=start_col + col_span - 1)
    c = ws.cell(row=3, column=start_col,
                value=f"{subtitle}    |    {UNIT}")
    c.font = FT_SUBTITLE
    c.fill = fill(C_MID_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18


def col_headers(ws, row, labels, start_col=1):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lbl)
        c.font = FT_HEAD
        c.fill = fill(C_MID_BLUE)
        c.alignment = Alignment(horizontal="center" if i > 0 else "left",
                                vertical="center", wrap_text=True)
        c.border = border_thin_all()
    ws.row_dimensions[row].height = 22


def group_row(ws, row, label, col_span, start_col=1):
    ws.merge_cells(start_row=row, start_column=start_col,
                   end_row=row, end_column=start_col + col_span - 1)
    c = ws.cell(row=row, column=start_col, value=f"  {label.upper()}")
    c.font = FT_GROUP
    c.fill = fill(C_DARK_BLUE)
    c.alignment = Alignment(vertical="center")
    c.border = border_thin_all()
    ws.row_dimensions[row].height = 20


def classification_row(ws, row, label, col_span, start_col=1):
    ws.merge_cells(start_row=row, start_column=start_col,
                   end_row=row, end_column=start_col + col_span - 1)
    c = ws.cell(row=row, column=start_col, value=f"    {label}")
    c.font = FT_SUBGROUP
    c.fill = fill(C_LIGHT_BLUE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 18


def data_row(ws, row, label, amount, note_ref, alt=False, indent=6):
    bg = C_PALE_BLUE if alt else C_WHITE
    ws.cell(row=row, column=1, value=f"{'  ' * (indent // 2)}{label}").font = FT_NORMAL
    ws.cell(row=row, column=1).fill = fill(bg)
    ws.cell(row=row, column=1).border = border_bottom_thin()

    c_amt = ws.cell(row=row, column=2, value=amount)
    c_amt.font = FT_NORMAL
    c_amt.fill = fill(bg)
    c_amt.number_format = FMT_NUM
    c_amt.alignment = Alignment(horizontal="right")
    c_amt.border = border_bottom_thin()

    c_note = ws.cell(row=row, column=3, value=note_ref if note_ref else "")
    c_note.font = FT_NOTE_NUM
    c_note.fill = fill(bg)
    c_note.alignment = Alignment(horizontal="center")
    c_note.border = border_bottom_thin()
    ws.row_dimensions[row].height = 16


def subtotal_row(ws, row, label, formula_or_val, col_span=3):
    c = ws.cell(row=row, column=1, value=f"  {label}")
    c.font = FT_TOTAL
    c.fill = fill(C_GOLD_LIGHT)
    c.border = border_bottom_thick()

    c2 = ws.cell(row=row, column=2, value=formula_or_val)
    c2.font = FT_TOTAL
    c2.fill = fill(C_GOLD_LIGHT)
    c2.number_format = FMT_NUM
    c2.alignment = Alignment(horizontal="right")
    c2.border = border_bottom_thick()

    ws.cell(row=row, column=3).fill = fill(C_GOLD_LIGHT)
    ws.cell(row=row, column=3).border = border_bottom_thick()
    ws.row_dimensions[row].height = 18


def grand_total_row(ws, row, label, formula_or_val, col_span=3):
    for col in range(1, col_span + 1):
        c = ws.cell(row=row, column=col)
        c.fill = fill(C_NAVY)
        c.font = FT_GRAND
        c.border = border_thin_all()
    ws.cell(row=row, column=1, value=f"  {label}").font = FT_GRAND
    ws.cell(row=row, column=1).fill = fill(C_NAVY)

    c2 = ws.cell(row=row, column=2, value=formula_or_val)
    c2.font = FT_GRAND
    c2.fill = fill(C_NAVY)
    c2.number_format = FMT_NUM
    c2.alignment = Alignment(horizontal="right")
    ws.row_dimensions[row].height = 22


# ================= LOAD DATA =================
def load_data(tb_path, map_path):
    tb = pd.read_excel(tb_path)
    mp = pd.read_excel(map_path)

    tb["Account"] = tb["Account"].astype(str).str.strip()
    mp["Account"] = mp["Account"].astype(str).str.strip()
    tb["Account_Name"] = tb["Account_Name"].astype(str).str.strip()
    mp["Account_Name"] = mp["Account_Name"].astype(str).str.strip()

    df = tb.merge(mp, on=["Account", "Account_Name"], how="left")
    df["Closing_Balance"] = pd.to_numeric(df["Closing_Balance"], errors="coerce").fillna(0)
    df["Opening_Balance"] = pd.to_numeric(df["Opening_Balance"], errors="coerce").fillna(0)

    for col in ["Statement", "Group", "Classification", "LineItem"]:
        df[col] = df[col].astype(str).str.strip()
    df["SubLineItem"] = df["SubLineItem"].fillna("").astype(str).str.strip()
    df["Notes"] = df["Notes"].fillna("").astype(str).str.strip()

    df["Statement"] = df["Statement"].str.upper()
    return df


# ================= NOTE REFERENCES =================
NOTE_MAP = {
    "Cash and cash equivalents":       "Note 1",
    "Trade receivables":               "Note 2",
    "Inventories":                     "Note 3",
    "Prepayments and other assets":    "Note 4",
    "Property, plant and equipment":   "Note 5",
    "Right-of-use assets":             "Note 6",
    "Intangible assets":               "Note 7",
    "VAT receivable":                  "Note 8",
    "Trade payables":                  "Note 9",
    "Accrued expenses":                "Note 10",
    "Lease liabilities":               "Note 11",
    "VAT payable":                     "Note 12",
    "Borrowings":                      "Note 13",
    "Share capital":                   "Note 14",
    "Retained earnings":               "Note 15",
    "Revenue":                         "Note 16",
    "Cost of sales":                   "Note 17",
    "Salaries and wages":              "Note 18",
    "Depreciation":                    "Note 19",
    "Finance costs":                   "Note 20",
}


# ================= BALANCE SHEET =================
def write_bs(wb, bs, df_raw):
    ws = wb.create_sheet("Balance Sheet")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 10

    title_block(ws, "Statement of Financial Position", PERIOD, 3)
    col_headers(ws, 4, ["Particulars", "Amount (INR)", "Note Ref"], 1)

    row = 5
    grand_rows = {}   # group -> list of subtotal rows for grand total formula

    for grp_order, grp in enumerate(["Assets", "Liabilities", "Equity"]):
        df_g = bs[bs["Group"].str.lower() == grp.lower()]
        if df_g.empty:
            continue

        group_row(ws, row, grp, 3)
        row += 1
        subtotal_rows_for_group = []

        for cls in df_g["Classification"].unique():
            df_c = df_g[df_g["Classification"] == cls]

            classification_row(ws, row, cls, 3)
            row += 1

            line_rows = []
            alt = False
            for _, r in df_c.iterrows():
                note_ref = NOTE_MAP.get(r["LineItem"], "")
                data_row(ws, row, r["LineItem"], r["Amount"], note_ref, alt)
                line_rows.append(row)
                alt = not alt
                row += 1

            # Classification subtotal
            if len(line_rows) > 0:
                sub_formula = f"=SUM(B{line_rows[0]}:B{line_rows[-1]})"
                subtotal_row(ws, row, f"Total {cls} {grp}", sub_formula)
                subtotal_rows_for_group.append(row)
                row += 1

        # Group total
        if subtotal_rows_for_group:
            refs = "+".join([f"B{r}" for r in subtotal_rows_for_group])
            grand_total_row(ws, row, f"TOTAL {grp.upper()}", f"={refs}")
            grand_rows[grp] = row
            row += 2

    # Balance sheet check row
    all_grand = "+".join([f"B{r}" for r in grand_rows.values()])
    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"A{row}:B{row}")
    c = ws.cell(row=row, column=1,
                value=f'=IF(ABS({all_grand})<1,"✔  Balance Sheet Balances","✘  Check Required")')
    c.font = Font(name="Calibri", bold=True, size=10, color="375623")
    c.fill = fill(C_GREEN_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = border_thin_all()

    autofit(ws)
    ws.freeze_panes = "A5"
    return ws


# ================= P&L =================
def write_pl(wb, pl, profit):
    ws = wb.create_sheet("P&L Statement")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 10

    title_block(ws, "Statement of Profit or Loss", PERIOD, 3)
    col_headers(ws, 4, ["Particulars", "Amount (INR)", "Note Ref"], 1)

    row = 5

    order = ["Income", "Contra-Income", "Expense"]
    group_label = {"Income": "Revenue & Income",
                   "Contra-Income": "Contra Revenue",
                   "Expense": "Operating & Finance Expenses"}

    subtotal_refs = {}

    for grp in order:
        df_g = pl[pl["Group"].str.lower() == grp.lower()]
        if df_g.empty:
            continue

        group_row(ws, row, group_label.get(grp, grp), 3)
        row += 1

        line_rows = []
        alt = False
        for _, r in df_g.iterrows():
            note_ref = NOTE_MAP.get(r["LineItem"], "")
            data_row(ws, row, r["LineItem"], r["Amount"], note_ref, alt)
            line_rows.append(row)
            alt = not alt
            row += 1

        sub_formula = f"=SUM(B{line_rows[0]}:B{line_rows[-1]})"
        subtotal_row(ws, row, f"Total {group_label.get(grp, grp)}", sub_formula)
        subtotal_refs[grp] = row
        row += 2

    # ===== FIXED NET PROFIT =====
    inc  = subtotal_refs.get("Income", None)
    con  = subtotal_refs.get("Contra-Income", None)
    exp  = subtotal_refs.get("Expense", None)

    if inc and con and exp:
        profit_formula = f"=B{inc}+B{con}-B{exp}"
    elif inc and exp:
        profit_formula = f"=B{inc}-B{exp}"
    else:
        profit_formula = profit

    grand_total_row(ws, row, "NET PROFIT / (LOSS) FOR THE YEAR", profit_formula)
    row += 2

    c = ws.cell(row=row, column=1, value="  Earnings Per Share (Basic) — Refer Note 21")
    c.font = FT_ITALIC
    ws.row_dimensions[row].height = 16

    autofit(ws)
    ws.freeze_panes = "A5"
    return ws

# ================= NOTES OF ACCOUNTS =================

def notes_header(ws, row, note_num, note_title, col_span=5):
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=col_span)
    c = ws.cell(row=row, column=1,
                value=f"  {note_num}  |  {note_title.upper()}")
    c.font = Font(name="Calibri", bold=True, size=11, color=C_WHITE)
    c.fill = fill(C_DARK_BLUE)
    c.alignment = Alignment(vertical="center")
    c.border = border_thin_all()
    ws.row_dimensions[row].height = 22


def notes_col_header(ws, row, labels, col_span=5):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=i + 1, value=lbl)
        c.font = FT_HEAD
        c.fill = fill(C_MID_BLUE)
        c.alignment = Alignment(horizontal="right" if i > 0 else "left",
                                vertical="center")
        c.border = border_thin_all()
    ws.row_dimensions[row].height = 20


def notes_data_row(ws, row, label, opening, movement, closing, alt=False):
    bg = C_PALE_BLUE if alt else C_WHITE
    cols_vals = [("left", label, FT_NORMAL, None),
                 ("right", opening, FT_NORMAL, FMT_NUM),
                 ("right", movement, FT_NORMAL, FMT_NUM),
                 ("right", closing, FT_NORMAL, FMT_NUM)]
    for i, (align_, val, fnt, nfmt) in enumerate(cols_vals):
        c = ws.cell(row=row, column=i + 1, value=val)
        c.font = fnt
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal=align_)
        c.border = border_bottom_thin()
        if nfmt:
            c.number_format = nfmt
    ws.row_dimensions[row].height = 16


def notes_total_row(ws, row, label, opening, movement, closing):
    vals = [label, opening, movement, closing]
    for i, val in enumerate(vals):
        c = ws.cell(row=row, column=i + 1, value=val)
        c.font = FT_TOTAL
        c.fill = fill(C_GOLD_LIGHT)
        c.border = border_bottom_thick()
        if i > 0:
            c.number_format = FMT_NUM
            c.alignment = Alignment(horizontal="right")
    ws.row_dimensions[row].height = 18


def notes_spacer(ws, row):
    ws.row_dimensions[row].height = 10


def write_notes(wb, df):
    ws = wb.create_sheet("Notes to Accounts")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 15

    # Title
    title_block(ws, "Notes to Financial Statements", PERIOD, 5)
    row = 5

    # ---- Intro blurb ----
    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row=row, column=1,
                value="  All amounts are in INR unless otherwise stated. "
                      "These notes form an integral part of the financial statements.")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    ws.row_dimensions[row].height = 16
    row += 2

    # Determine all BS line items with opening/movement/closing
    bs_df = df[df["Statement"] == "BS"].copy()
    pl_df = df[df["Statement"] == "PL"].copy()

    # ---------- NOTE 1: Cash ----------
    notes_header(ws, row, "Note 1", "Cash and Cash Equivalents", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Opening Balance", "Net Movement", "Closing Balance"], 4)
    row += 1
    accounts = bs_df[bs_df["LineItem"] == "Cash and cash equivalents"]
    alt = False
    sub_rows = []
    for _, r in accounts.iterrows():
        notes_data_row(ws, row, r["Account_Name"],
                       abs(r["Opening_Balance"]),
                       abs(r["Closing_Balance"]) - abs(r["Opening_Balance"]),
                       abs(r["Closing_Balance"]), alt)
        sub_rows.append(row)
        alt = not alt
        row += 1
    if sub_rows:
        tot_open  = accounts["Opening_Balance"].abs().sum()
        tot_close = accounts["Closing_Balance"].abs().sum()
        notes_total_row(ws, row, "  Total Cash and Cash Equivalents",
                        tot_open, tot_close - tot_open, tot_close)
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 2: Trade Receivables ----------
    notes_header(ws, row, "Note 2", "Trade Receivables", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Opening Balance", "Net Movement", "Closing Balance"], 4)
    row += 1
    accounts = bs_df[bs_df["LineItem"] == "Trade receivables"]
    alt = False
    for _, r in accounts.iterrows():
        mv = r["Closing_Balance"] - r["Opening_Balance"]
        notes_data_row(ws, row, r["Account_Name"],
                       r["Opening_Balance"], mv, r["Closing_Balance"], alt)
        alt = not alt
        row += 1
    # Net receivable note
    gross = bs_df[bs_df["Account_Name"] == "Trade Receivables"]["Closing_Balance"].sum()
    prov  = bs_df[bs_df["Account_Name"] == "Allowance for Doubtful Debts"]["Closing_Balance"].sum()
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value=f"  Net Trade Receivable: INR {gross + prov:,.0f}  |  "
                      f"Provision for Bad Debts: INR {abs(prov):,.0f}")
    c.font = FT_ITALIC
    c.fill = fill(C_RED_LIGHT)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 3: Inventories ----------
    notes_header(ws, row, "Note 3", "Inventories", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Opening Balance", "Net Movement", "Closing Balance"], 4)
    row += 1
    accounts = bs_df[bs_df["LineItem"] == "Inventories"]
    alt = False
    for _, r in accounts.iterrows():
        notes_data_row(ws, row, r["Account_Name"],
                       r["Opening_Balance"],
                       r["Closing_Balance"] - r["Opening_Balance"],
                       r["Closing_Balance"], alt)
        alt = not alt
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 5: PPE ----------
    notes_header(ws, row, "Note 5", "Property, Plant and Equipment (PPE)", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Gross Block", "Accumulated Dep.", "Net Book Value"], 4)
    row += 1
    ppe_gross = bs_df[bs_df["Account_Name"] == "Property, Plant & Equipment"]["Closing_Balance"].sum()
    ppe_dep   = bs_df[bs_df["Account_Name"] == "Accumulated Depreciation - PPE"]["Closing_Balance"].sum()
    notes_data_row(ws, row, "  Property, Plant & Equipment", ppe_gross, abs(ppe_dep), ppe_gross + ppe_dep)
    row += 1
    rou_gross = bs_df[bs_df["Account_Name"] == "Right-of-Use Asset"]["Closing_Balance"].sum()
    rou_dep   = bs_df[bs_df["Account_Name"] == "Accumulated Depreciation - ROU"]["Closing_Balance"].sum()
    notes_data_row(ws, row, "  Right-of-Use Asset (IFRS 16)", rou_gross, abs(rou_dep), rou_gross + rou_dep, True)
    row += 1
    int_gross = bs_df[bs_df["Account_Name"] == "Intangible Assets"]["Closing_Balance"].sum()
    int_dep   = bs_df[bs_df["Account_Name"] == "Accumulated Amortization - Intangibles"]["Closing_Balance"].sum()
    notes_data_row(ws, row, "  Intangible Assets", int_gross, abs(int_dep), int_gross + int_dep)
    row += 1
    notes_total_row(ws, row, "  TOTAL",
                    ppe_gross + rou_gross + int_gross,
                    abs(ppe_dep) + abs(rou_dep) + abs(int_dep),
                    (ppe_gross + ppe_dep) + (rou_gross + rou_dep) + (int_gross + int_dep))
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="  Depreciation method: Straight Line Method (SLM). "
                      "Useful life determined per Ind AS 16 / IFRS guidelines.")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 9: Trade Payables ----------
    notes_header(ws, row, "Note 9", "Trade Payables", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Opening Balance", "Net Movement", "Closing Balance"], 4)
    row += 1
    accounts = bs_df[bs_df["LineItem"] == "Trade payables"]
    alt = False
    for _, r in accounts.iterrows():
        op = abs(r["Opening_Balance"])
        cl = abs(r["Closing_Balance"])
        notes_data_row(ws, row, r["Account_Name"], op, cl - op, cl, alt)
        alt = not alt
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 13: Borrowings ----------
    notes_header(ws, row, "Note 13", "Borrowings", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Opening Balance", "Net Movement", "Closing Balance"], 4)
    row += 1
    accounts = bs_df[bs_df["LineItem"] == "Borrowings"]
    alt = False
    for _, r in accounts.iterrows():
        op = abs(r["Opening_Balance"])
        cl = abs(r["Closing_Balance"])
        notes_data_row(ws, row, r["Account_Name"], op, cl - op, cl, alt)
        alt = not alt
        row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="  Bank loans carry floating interest rate linked to MCLR/SOFR. "
                      "Current portion reclassified under Current Liabilities.")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 14: Share Capital ----------
    notes_header(ws, row, "Note 14", "Share Capital", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "No. of Shares", "Face Value (INR)", "Amount (INR)"], 4)
    row += 1
    sc = abs(bs_df[bs_df["LineItem"] == "Share capital"]["Closing_Balance"].sum())
    notes_data_row(ws, row, "  Equity Shares of INR 10 each, fully paid", sc // 10, 10, sc)
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="  All shares rank equally with regard to dividend rights and voting.")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 15: Retained Earnings ----------
    notes_header(ws, row, "Note 15", "Retained Earnings", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Amount (INR)", "", ""], 4)
    row += 1
    re_accounts = bs_df[bs_df["LineItem"] == "Retained earnings"]
    for _, r in re_accounts.iterrows():
        notes_data_row(ws, row, r["Account_Name"],
                       abs(r["Opening_Balance"]), "", abs(r["Closing_Balance"]))
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 16: Revenue ----------
    notes_header(ws, row, "Note 16", "Revenue from Operations", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Current Year (INR)", "", ""], 4)
    row += 1
    rev_accounts = pl_df[pl_df["Group"].str.lower().isin(["income", "contra-income"])]
    alt = False
    rev_rows = []
    for _, r in rev_accounts.iterrows():
        notes_data_row(ws, row, r["Account_Name"], "", "", abs(r["Closing_Balance"]), alt)
        rev_rows.append(row)
        alt = not alt
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 17: Cost of Sales ----------
    notes_header(ws, row, "Note 17", "Cost of Sales", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Current Year (INR)", "", ""], 4)
    row += 1
    cogs = pl_df[pl_df["LineItem"] == "Cost of sales"]
    for _, r in cogs.iterrows():
        notes_data_row(ws, row, r["Account_Name"], "", "", r["Closing_Balance"])
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 18: Employee Benefits ----------
    notes_header(ws, row, "Note 18", "Employee Benefit Expenses", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Current Year (INR)", "", ""], 4)
    row += 1
    emp_items = ["Salaries and wages", "Staff benefits"]
    emp = pl_df[pl_df["LineItem"].isin(emp_items)]
    alt = False
    emp_rows_list = []
    for _, r in emp.iterrows():
        notes_data_row(ws, row, r["Account_Name"], "", "", r["Closing_Balance"], alt)
        emp_rows_list.append(row)
        alt = not alt
        row += 1
    tot_emp = emp["Closing_Balance"].sum()
    notes_total_row(ws, row, "  Total Employee Benefit Expenses", "", "", tot_emp)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 19: Depreciation ----------
    notes_header(ws, row, "Note 19", "Depreciation and Amortization", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Current Year (INR)", "", ""], 4)
    row += 1
    dep_items = pl_df[pl_df["LineItem"].isin(["Depreciation", "Amortization"])]
    alt = False
    for _, r in dep_items.iterrows():
        notes_data_row(ws, row, r["Account_Name"], "", "", r["Closing_Balance"], alt)
        alt = not alt
        row += 1
    tot_dep = dep_items["Closing_Balance"].sum()
    notes_total_row(ws, row, "  Total Depreciation & Amortization", "", "", tot_dep)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 20: Finance Costs ----------
    notes_header(ws, row, "Note 20", "Finance Costs", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Current Year (INR)", "", ""], 4)
    row += 1
    fin = pl_df[pl_df["LineItem"] == "Finance costs"]
    for _, r in fin.iterrows():
        notes_data_row(ws, row, r["Account_Name"], "", "", r["Closing_Balance"])
        row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 21: EPS ----------
    notes_header(ws, row, "Note 21", "Earnings Per Share (EPS)", 4)
    row += 1
    notes_col_header(ws, row, ["Particulars", "Amount (INR)", "", ""], 4)
    row += 1

    total_pl = pl_df["Closing_Balance"].sum() * -1
    sc_val   = abs(bs_df[bs_df["LineItem"] == "Share capital"]["Closing_Balance"].sum())
    shares   = sc_val // 10

    notes_data_row(ws, row, "  Net Profit / (Loss) for the year", "", "", total_pl)
    row += 1
    notes_data_row(ws, row, "  Weighted Average No. of Equity Shares", "", "", shares, True)
    row += 1
    eps = total_pl / shares if shares else 0
    notes_data_row(ws, row, "  Basic Earnings Per Share (INR)", "", "", round(eps, 2))
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="  Diluted EPS = Basic EPS (No dilutive instruments outstanding during the year).")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    row += 1
    notes_spacer(ws, row); row += 2

    # ---------- NOTE 22: Accounting Policies ----------
    notes_header(ws, row, "Note 22", "Significant Accounting Policies", 4)
    row += 1
    policies = [
        ("Basis of Preparation",
         "Financial statements are prepared under the historical cost convention on "
         "accrual basis in accordance with Ind AS / IFRS."),
        ("Revenue Recognition",
         "Revenue is recognized when control of goods or services is transferred to "
         "the customer at an amount that reflects the consideration entitled (Ind AS 115)."),
        ("Property, Plant & Equipment",
         "PPE is stated at cost less accumulated depreciation. Depreciation is provided "
         "on SLM basis over estimated useful life per Schedule II of Companies Act 2013."),
        ("Inventories",
         "Inventories are valued at the lower of cost (weighted average) or net realizable value."),
        ("Financial Instruments",
         "Financial assets and liabilities are recognized at fair value. Trade "
         "receivables are measured at amortized cost."),
        ("Leases (Ind AS 116 / IFRS 16)",
         "The company recognizes right-of-use assets and lease liabilities for all "
         "leases with a term exceeding 12 months."),
        ("Taxation",
         "Current tax is recognized based on the estimated tax payable under applicable "
         "tax laws. Deferred tax is recognized on temporary differences."),
    ]
    alt = False
    for policy_name, policy_text in policies:
        bg = C_PALE_BLUE if alt else C_WHITE
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1,
                    value=f"  {policy_name}:  {policy_text}")
        c.font = FT_NORMAL
        c.fill = fill(bg)
        c.alignment = Alignment(wrap_text=True)
        c.border = border_bottom_thin()
        ws.row_dimensions[row].height = 36
        alt = not alt
        row += 1

    notes_spacer(ws, row); row += 2

    # ---------- NOTE 23: Contingent Liabilities ----------
    notes_header(ws, row, "Note 23", "Contingent Liabilities and Commitments", 4)
    row += 1
    items = [
        "  No material contingent liabilities exist as at the reporting date.",
        "  Capital commitments contracted but not provided: NIL.",
        "  Operating lease commitments: disclosed under Note 11.",
    ]
    for item in items:
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=item)
        c.font = FT_ITALIC
        c.fill = fill(C_GRAY_LIGHT)
        row += 1
    notes_spacer(ws, row); row += 2

    # Auditor's note
    notes_header(ws, row, "Note 24", "Auditor's Remark (For Demo Purposes)", 4)
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="  These financial statements have been prepared using AI-assisted "
                      "automation for demonstration purposes by DemoCo Pvt. Ltd. "
                      "All figures are indicative and should be independently verified before use.")
    c.font = FT_ITALIC
    c.fill = fill(C_PALE_BLUE)
    c.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[row].height = 36
    row += 1

    ws.freeze_panes = "A5"
    autofit(ws)
    return ws


# ================= COVER SHEET =================
def write_cover(wb):
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    for col_letter in ["A", "B", "C", "D", "E"]:
        ws.column_dimensions[col_letter].width = 22

    # Big title block
    ws.merge_cells("A1:E4")
    c = ws.cell(row=1, column=1, value=COMPANY)
    c.font = Font(name="Calibri", bold=True, size=28, color=C_WHITE)
    c.fill = fill(C_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 80

    ws.merge_cells("A5:E6")
    c = ws.cell(row=5, column=1, value="ANNUAL FINANCIAL STATEMENTS")
    c.font = Font(name="Calibri", bold=True, size=18, color=C_WHITE)
    c.fill = fill(C_DARK_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 40

    ws.merge_cells("A7:E7")
    c = ws.cell(row=7, column=1, value=PERIOD)
    c.font = Font(name="Calibri", size=13, color=C_WHITE)
    c.fill = fill(C_MID_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 30

    # Index table
    row = 9
    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row=row, column=1, value="  TABLE OF CONTENTS")
    c.font = Font(name="Calibri", bold=True, size=12, color=C_WHITE)
    c.fill = fill(C_DARK_BLUE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 22
    row += 1

    contents = [
        ("1", "Statement of Financial Position", "Balance Sheet"),
        ("2", "Statement of Profit or Loss", "P&L Statement"),
        ("3", "Notes to Financial Statements (Note 1–24)", "Notes to Accounts"),
        ("4", "Trial Balance (Source Data)", "TB Source"),
        ("5", "Account Mapping File", "Mapping"),
    ]

    for i, (num, title_, sheet_ref) in enumerate(contents):
        bg = C_PALE_BLUE if i % 2 == 0 else C_WHITE
        ws.merge_cells(f"A{row}:B{row}")
        c = ws.cell(row=row, column=1, value=f"  {num}.  {title_}")
        c.font = FT_NORMAL
        c.fill = fill(bg)
        c.border = border_bottom_thin()

        c2 = ws.cell(row=row, column=3, value=sheet_ref)
        c2.font = Font(name="Calibri", size=10, color="2F5496", underline="single")
        c2.fill = fill(bg)
        c2.border = border_bottom_thin()
        c2.alignment = Alignment(horizontal="center")
        row += 1

    row += 1
    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row=row, column=1,
                value=f"  Prepared using Python + openpyxl automation  |  "
                      f"Source: TB + Account Mapping  |  {PERIOD}")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)
    row += 1

    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row=row, column=1,
                value="  All figures in Indian Rupees (INR) unless otherwise stated")
    c.font = FT_ITALIC
    c.fill = fill(C_GRAY_LIGHT)


# ================= TB SHEET =================
def write_tb(wb, df_raw, tb_path=None):
    ws = wb.create_sheet("TB Source")
    ws.sheet_view.showGridLines = False

    title_block(ws, "Trial Balance — Source Data", PERIOD, 8)
    headers = ["Entity", "Period", "Account", "Account Name",
               "Opening Balance", "Debit", "Credit", "Closing Balance"]
    col_headers(ws, 4, headers, 1)

    for col, w in zip("ABCDEFGH", [12, 10, 12, 35, 18, 14, 14, 18]):
        ws.column_dimensions[col].width = w

    row = 5
    tb = pd.read_excel(tb_path or TB_FILE)
    for i, r in tb.iterrows():
        alt = i % 2 == 0
        bg = C_PALE_BLUE if alt else C_WHITE
        vals = [r["Entity"], str(r["Period"]), str(r["Account"]),
                r["Account_Name"], r["Opening_Balance"],
                r["Debit"], r["Credit"], r["Closing_Balance"]]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            c.fill = fill(bg)
            c.font = FT_NORMAL
            c.border = border_bottom_thin()
            if col_idx >= 5:
                c.number_format = FMT_NUM
                c.alignment = Alignment(horizontal="right")
        ws.row_dimensions[row].height = 16
        row += 1

    # Totals
    for col_idx in [5, 6, 7, 8]:
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=row, column=col_idx,
                    value=f"=SUM({col_letter}5:{col_letter}{row-1})")
        c.font = FT_TOTAL
        c.fill = fill(C_GOLD_LIGHT)
        c.number_format = FMT_NUM
        c.alignment = Alignment(horizontal="right")
        c.border = border_bottom_thick()
    ws.cell(row=row, column=4, value="  TOTAL").font = FT_TOTAL
    ws.cell(row=row, column=4).fill = fill(C_GOLD_LIGHT)
    ws.row_dimensions[row].height = 20
    ws.freeze_panes = "A5"


# ================= MAIN =================
# ================= MAIN =================
def main():
    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)

    df = load_data(TB_FILE, MAP_FILE)

    bs_mask = df["Statement"] == "BS"
    pl_mask = df["Statement"] == "PL"

    # ================= FIX 1: BS SIGN =================
    df["BS_Amount"] = df["Closing_Balance"]
    df.loc[bs_mask & df["Group"].str.lower().isin(["liabilities", "equity"]), "BS_Amount"] *= -1

    # ================= FIX 2: PL SIGN =================
    df["PL_Amount"] = df["Closing_Balance"]
    df.loc[pl_mask & df["Group"].str.lower().isin(["income", "contra-income"]), "PL_Amount"] *= -1

    grp_cols = ["Statement", "Group", "Classification", "LineItem", "SubLineItem"]

    bs = (df[bs_mask][grp_cols + ["BS_Amount"]]
          .groupby(grp_cols, as_index=False)["BS_Amount"]
          .sum()
          .rename(columns={"BS_Amount": "Amount"}))

    pl = (df[pl_mask][grp_cols + ["PL_Amount"]]
          .groupby(grp_cols, as_index=False)["PL_Amount"]
          .sum()
          .rename(columns={"PL_Amount": "Amount"}))

    # Profit calculation
    income  = pl[pl["Group"].str.lower()=="income"]["Amount"].sum()
    contra  = pl[pl["Group"].str.lower()=="contra-income"]["Amount"].sum()
    expense = pl[pl["Group"].str.lower()=="expense"]["Amount"].sum()

    net_revenue = income + contra
    profit = net_revenue - expense

    # Add profit to BS
    bs = pd.concat([bs, pd.DataFrame([{
        "Statement": "BS",
        "Group": "Equity",
        "Classification": "Equity",
        "LineItem": "Current year profit/(loss)",
        "SubLineItem": "",
        "Amount": profit
    }])], ignore_index=True)

    wb = Workbook()
    wb.remove(wb.active)

    write_cover(wb)
    write_bs(wb, bs, df)
    write_pl(wb, pl, profit)
    write_notes(wb, df)
    write_tb(wb, df)

    wb.save(OUT_FILE)
    print("DONE:", OUT_FILE)


# ================= API ENTRY POINT =================
def generate(tb_path, map_path, out_path, company=COMPANY, period=PERIOD):
    """Called by FastAPI — accepts file paths, returns the output Excel path."""
    global COMPANY, PERIOD
    COMPANY = company
    PERIOD  = period

    df = load_data(tb_path, map_path)

    bs_mask = df["Statement"] == "BS"
    pl_mask = df["Statement"] == "PL"

    df["BS_Amount"] = df["Closing_Balance"]
    df.loc[bs_mask & df["Group"].str.lower().isin(["liabilities", "equity"]), "BS_Amount"] *= -1

    df["PL_Amount"] = df["Closing_Balance"]
    df.loc[pl_mask & df["Group"].str.lower().isin(["income", "contra-income"]), "PL_Amount"] *= -1

    grp_cols = ["Statement", "Group", "Classification", "LineItem", "SubLineItem"]

    bs = (df[bs_mask][grp_cols + ["BS_Amount"]]
          .groupby(grp_cols, as_index=False)["BS_Amount"]
          .sum()
          .rename(columns={"BS_Amount": "Amount"}))

    pl = (df[pl_mask][grp_cols + ["PL_Amount"]]
          .groupby(grp_cols, as_index=False)["PL_Amount"]
          .sum()
          .rename(columns={"PL_Amount": "Amount"}))

    income  = pl[pl["Group"].str.lower() == "income"]["Amount"].sum()
    contra  = pl[pl["Group"].str.lower() == "contra-income"]["Amount"].sum()
    expense = pl[pl["Group"].str.lower() == "expense"]["Amount"].sum()
    profit  = (income + contra) - expense

    bs = pd.concat([bs, pd.DataFrame([{
        "Statement": "BS", "Group": "Equity", "Classification": "Equity",
        "LineItem": "Current year profit/(loss)", "SubLineItem": "", "Amount": profit
    }])], ignore_index=True)

    wb = Workbook()
    wb.remove(wb.active)

    write_cover(wb)
    write_bs(wb, bs, df)
    write_pl(wb, pl, profit)
    write_notes(wb, df)
    write_tb(wb, df, tb_path)

    wb.save(out_path)
    return out_path


# ================= ENTRY POINT =================
if __name__ == "__main__":
    main()