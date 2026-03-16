"""
Loan Portfolio Dashboard Generator
Generates synthetic loan data and builds a formula-driven Excel dashboard.
Usage: python generate_dashboard.py [--output /path/to/folder]
"""

import argparse
import os
import random
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── CLI ───────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--output", default=".", help="Output directory for the Excel file")
args = parser.parse_args()
OUT_DIR = args.output
os.makedirs(OUT_DIR, exist_ok=True)

# ── 1. GENERATE SYNTHETIC DATA ────────────────────────────────────────────────
np.random.seed(42)
random.seed(42)

NUM = 1000
SEGMENTS   = ["Retail", "SME", "Corporate", "Agri"]
PRODUCTS   = ["Personal", "Home", "Auto", "Business"]
AGE_GROUPS = ["18-30", "31-45", "46-60", "60+"]
RATINGS    = ["AAA", "AA", "A", "BBB", "BB", "B", "C"]
RATE_MAP   = {"AAA": 0.07, "AA": 0.08, "A": 0.09, "BBB": 0.11,
              "BB": 0.13, "B": 0.15, "C": 0.18}

rows = []
for i in range(1, NUM + 1):
    seg    = random.choice(SEGMENTS)
    prod   = random.choice(PRODUCTS)
    age    = random.choice(AGE_GROUPS)
    rating = random.choice(RATINGS)

    if prod == "Home":
        disbursed = int(np.random.randint(150_000, 500_000))
    elif prod == "Business":
        disbursed = int(np.random.randint(50_000, 300_000))
    else:
        disbursed = int(np.random.randint(5_000, 50_000))

    outstanding   = round(disbursed * np.random.uniform(0.1, 0.95), 2)
    interest_rate = round(RATE_MAP[rating] + np.random.uniform(-0.005, 0.005), 4)

    if rating in ["AAA", "AA"]:
        dpd = int(np.random.choice([0, 0, 0, 5, 10], p=[0.8, 0.1, 0.05, 0.03, 0.02]))
    elif rating in ["A", "BBB"]:
        dpd = int(np.random.choice([0, 15, 30, 45], p=[0.7, 0.15, 0.1, 0.05]))
    else:
        dpd = int(np.random.choice([0, 45, 90, 120], p=[0.4, 0.3, 0.2, 0.1]))

    status = "NPA" if dpd > 90 else ("Delinquent" if dpd > 0 else "Active")
    profit = round((outstanding * (interest_rate - 0.03)) / 12, 2)

    rows.append([f"L-{1000+i}", seg, prod, age, rating,
                 disbursed, outstanding, interest_rate, status, dpd, profit])

COLS = ["Loan_ID", "Segment", "Product", "Age_Group", "Credit_Rating",
        "Disbursement_Amt", "Outstanding_Principal", "Interest_Rate",
        "Status", "DPD", "Monthly_Profit"]
df = pd.DataFrame(rows, columns=COLS)

# ── 2. STYLES ─────────────────────────────────────────────────────────────────
NAVY        = PatternFill("solid", fgColor="1F4E79")
BLUE        = PatternFill("solid", fgColor="2E75B6")
LIGHT_BLUE  = PatternFill("solid", fgColor="D6E4F0")
ALT         = PatternFill("solid", fgColor="EBF5FB")
WHITE       = PatternFill("solid", fgColor="FFFFFF")

HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, color="1F4E79", size=16)
LABEL_FONT = Font(bold=True, color="1F4E79", size=10)
KPI_FONT   = Font(bold=True, color="1F4E79", size=12)
SUB_FONT   = Font(bold=True, color="FFFFFF", size=10)

THIN_SIDE   = Side(style="thin",   color="AAAAAA")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE,
                     top=THIN_SIDE,  bottom=THIN_SIDE)

CENTER  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT    = Alignment(horizontal="left",   vertical="center")
RIGHT   = Alignment(horizontal="right",  vertical="center")


def hdr(ws, row, col, value, fill=NAVY, font=HDR_FONT, align=CENTER):
    c = ws.cell(row=row, column=col, value=value)
    c.fill, c.font, c.alignment, c.border = fill, font, align, THIN_BORDER
    return c


def val(ws, row, col, value, fmt=None, fill=WHITE, align=RIGHT):
    c = ws.cell(row=row, column=col, value=value)
    if fmt:
        c.number_format = fmt
    c.fill, c.alignment, c.border = fill, align, THIN_BORDER
    return c


def section_title(ws, row, title, col_start=1, col_end=6):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=title)
    c.fill      = NAVY
    c.font      = Font(bold=True, color="FFFFFF", size=11)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = THIN_BORDER
    ws.row_dimensions[row].height = 22


# ── 3. WORKBOOK ───────────────────────────────────────────────────────────────
wb = Workbook()

# ── DATA SHEET ────────────────────────────────────────────────────────────────
ws_d = wb.active
ws_d.title = "Data"

for ci, name in enumerate(COLS, 1):
    hdr(ws_d, 1, ci, name)

col_fmts = {6: "#,##0.00", 7: "#,##0.00", 8: "0.00%", 11: "#,##0.00"}
for ri, row_data in enumerate(df.itertuples(index=False), 2):
    for ci, v in enumerate(row_data, 1):
        c = ws_d.cell(row=ri, column=ci, value=v)
        c.border = THIN_BORDER
        if ci in col_fmts:
            c.number_format = col_fmts[ci]

col_widths_d = [10, 10, 10, 10, 13, 18, 22, 14, 12, 6, 15]
for i, w in enumerate(col_widths_d, 1):
    ws_d.column_dimensions[get_column_letter(i)].width = w
ws_d.freeze_panes = "A2"

# ── DASHBOARD SHEET ───────────────────────────────────────────────────────────
ws = wb.create_sheet("Dashboard")
ws.sheet_view.showGridLines = False

col_widths_dash = {1: 26, 2: 18, 3: 18, 4: 18, 5: 20, 6: 18}
for col, w in col_widths_dash.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# Formula range aliases (absolute refs into Data sheet)
SEG  = "Data!$B$2:$B$1001"
PROD = "Data!$C$2:$C$1001"
AGE  = "Data!$D$2:$D$1001"
RAT  = "Data!$E$2:$E$1001"
DIS  = "Data!$F$2:$F$1001"
OUT  = "Data!$G$2:$G$1001"
RATE = "Data!$H$2:$H$1001"
STA  = "Data!$I$2:$I$1001"
DPD_ = "Data!$J$2:$J$1001"
PROF = "Data!$K$2:$K$1001"

# ── TITLE ─────────────────────────────────────────────────────────────────────
ws.merge_cells("A1:F1")
c = ws["A1"]
c.value     = "LOAN PORTFOLIO DASHBOARD"
c.font      = TITLE_FONT
c.fill      = LIGHT_BLUE
c.alignment = CENTER
c.border    = THIN_BORDER
ws.row_dimensions[1].height = 36

ws.merge_cells("A2:F2")
c = ws["A2"]
c.value     = (f"Generated: {pd.Timestamp.today().strftime('%d %b %Y')} "
               f"| Records: {NUM:,}")
c.font      = Font(italic=True, color="666666", size=9)
c.fill      = WHITE
c.alignment = CENTER
c.border    = THIN_BORDER

R = 4  # current row cursor

# ── SECTION 1: Executive KPIs ─────────────────────────────────────────────────
section_title(ws, R, "  EXECUTIVE KPIs")
R += 1

kpis = [
    ("Total Loans",               f"=COUNTA({SEG})",                                    "0"),
    ("Total Disbursed (₹)",       f"=SUM({DIS})",                                       "#,##0.00"),
    ("Total Outstanding (₹)",     f"=SUM({OUT})",                                       "#,##0.00"),
    ("NPA Count",                 f'=COUNTIF({STA},"NPA")',                              "0"),
    ("NPA Rate (%)",              f'=COUNTIF({STA},"NPA")/COUNTA({SEG})',                "0.00%"),
    ("Delinquent Count",          f'=COUNTIF({STA},"Delinquent")',                       "0"),
    ("Portfolio Yield (%)",       f"=SUMPRODUCT({RATE},{OUT})/SUM({OUT})",               "0.00%"),
    ("Total Monthly Profit (₹)",  f"=SUM({PROF})",                                      "#,##0.00"),
    ("Avg Loan Size (₹)",         f"=SUM({OUT})/COUNTA({SEG})",                         "#,##0.00"),
    ("Active Loans",              f'=COUNTIF({STA},"Active")',                           "0"),
]

for idx, (label, formula, fmt) in enumerate(kpis):
    row = R + (idx // 2)
    col = 1 if idx % 2 == 0 else 4

    lc = ws.cell(row=row, column=col, value=label)
    lc.font, lc.fill, lc.alignment, lc.border = (
        LABEL_FONT, LIGHT_BLUE, LEFT, THIN_BORDER)

    ws.merge_cells(start_row=row, start_column=col+1,
                   end_row=row,   end_column=col+2)
    vc = ws.cell(row=row, column=col+1, value=formula)
    vc.font, vc.fill, vc.number_format = KPI_FONT, WHITE, fmt
    vc.alignment, vc.border = CENTER, THIN_BORDER

R += (len(kpis) + 1) // 2 + 1

# ── SECTION 2: NPA by Segment ─────────────────────────────────────────────────
section_title(ws, R, "  NPA BY SEGMENT")
R += 1

for ci, h in enumerate(["Segment", "Total Loans", "NPA Count",
                         "NPA %", "Outstanding (NPA ₹)", "Delinquent Count"], 1):
    hdr(ws, R, ci, h, fill=BLUE, font=SUB_FONT)
ws.row_dimensions[R].height = 28
R += 1

for i, seg in enumerate(SEGMENTS):
    fill = ALT if i % 2 == 0 else WHITE
    val(ws, R, 1, seg,                                                      fill=fill, align=LEFT)
    val(ws, R, 2, f'=COUNTIF({SEG},"{seg}")',              "#,##0",         fill=fill)
    val(ws, R, 3, f'=COUNTIFS({SEG},"{seg}",{STA},"NPA")', "#,##0",        fill=fill)
    val(ws, R, 4, f'=IFERROR(COUNTIFS({SEG},"{seg}",{STA},"NPA")'
                  f'/COUNTIF({SEG},"{seg}"),0)',            "0.00%",         fill=fill)
    val(ws, R, 5, f'=SUMIFS({OUT},{SEG},"{seg}",{STA},"NPA")', "#,##0.00", fill=fill)
    val(ws, R, 6, f'=COUNTIFS({SEG},"{seg}",{STA},"Delinquent")', "#,##0", fill=fill)
    R += 1

R += 1

# ── SECTION 3: Portfolio Concentration by Product ─────────────────────────────
section_title(ws, R, "  PORTFOLIO CONCENTRATION BY PRODUCT")
R += 1

for ci, h in enumerate(["Product", "Count", "Outstanding (₹)",
                         "% of Portfolio", "NPA Count", "Avg Interest Rate"], 1):
    hdr(ws, R, ci, h, fill=BLUE, font=SUB_FONT)
ws.row_dimensions[R].height = 28
R += 1

for i, prod in enumerate(PRODUCTS):
    fill = ALT if i % 2 == 0 else WHITE
    val(ws, R, 1, prod,                                                          fill=fill, align=LEFT)
    val(ws, R, 2, f'=COUNTIF({PROD},"{prod}")',               "#,##0",           fill=fill)
    val(ws, R, 3, f'=SUMIF({PROD},"{prod}",{OUT})',           "#,##0.00",        fill=fill)
    val(ws, R, 4, f'=IFERROR(SUMIF({PROD},"{prod}",{OUT})/SUM({OUT}),0)', "0.00%", fill=fill)
    val(ws, R, 5, f'=COUNTIFS({PROD},"{prod}",{STA},"NPA")', "#,##0",           fill=fill)
    val(ws, R, 6, f'=AVERAGEIF({PROD},"{prod}",{RATE})',      "0.00%",           fill=fill)
    R += 1

R += 1

# ── SECTION 4: Age Group vs DPD ───────────────────────────────────────────────
section_title(ws, R, "  AGE GROUP vs DPD ANALYSIS")
R += 1

for ci, h in enumerate(["Age Group", "Count", "Avg DPD",
                         "Delinquent Count", "NPA Count", "NPA %"], 1):
    hdr(ws, R, ci, h, fill=BLUE, font=SUB_FONT)
ws.row_dimensions[R].height = 28
R += 1

for i, age in enumerate(AGE_GROUPS):
    fill = ALT if i % 2 == 0 else WHITE
    val(ws, R, 1, age,                                                             fill=fill, align=LEFT)
    val(ws, R, 2, f'=COUNTIF({AGE},"{age}")',                  "#,##0",           fill=fill)
    val(ws, R, 3, f'=AVERAGEIF({AGE},"{age}",{DPD_})',         "0.0",             fill=fill)
    val(ws, R, 4, f'=COUNTIFS({AGE},"{age}",{STA},"Delinquent")', "#,##0",       fill=fill)
    val(ws, R, 5, f'=COUNTIFS({AGE},"{age}",{STA},"NPA")',     "#,##0",           fill=fill)
    val(ws, R, 6, f'=IFERROR(COUNTIFS({AGE},"{age}",{STA},"NPA")'
                  f'/COUNTIF({AGE},"{age}"),0)',                "0.00%",           fill=fill)
    R += 1

R += 1

# ── SECTION 5: Interest Rate Band Analysis ────────────────────────────────────
section_title(ws, R, "  INTEREST RATE BAND ANALYSIS")
R += 1

for ci, h in enumerate(["Rate Band", "Count", "Avg Rate",
                         "Total Outstanding (₹)", "Total Monthly Profit (₹)", "NPA Count"], 1):
    hdr(ws, R, ci, h, fill=BLUE, font=SUB_FONT)
ws.row_dimensions[R].height = 28
R += 1

bands = [
    ("Low  (< 8%)",
     f"=SUMPRODUCT(({RATE}<0.08)*1)",
     f"=IFERROR(SUMPRODUCT(({RATE}<0.08)*{RATE})/SUMPRODUCT(({RATE}<0.08)*1),0)",
     f"=SUMPRODUCT(({RATE}<0.08)*{OUT})",
     f"=SUMPRODUCT(({RATE}<0.08)*{PROF})",
     f'=SUMPRODUCT(({RATE}<0.08)*({STA}="NPA"))'),
    ("Medium (8–12%)",
     f"=SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*1)",
     f"=IFERROR(SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*{RATE})"
     f"/SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*1),0)",
     f"=SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*{OUT})",
     f"=SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*{PROF})",
     f'=SUMPRODUCT(({RATE}>=0.08)*({RATE}<0.12)*({STA}="NPA"))'),
    ("High  (≥ 12%)",
     f"=SUMPRODUCT(({RATE}>=0.12)*1)",
     f"=IFERROR(SUMPRODUCT(({RATE}>=0.12)*{RATE})/SUMPRODUCT(({RATE}>=0.12)*1),0)",
     f"=SUMPRODUCT(({RATE}>=0.12)*{OUT})",
     f"=SUMPRODUCT(({RATE}>=0.12)*{PROF})",
     f'=SUMPRODUCT(({RATE}>=0.12)*({STA}="NPA"))'),
]

for i, (label, cnt, avg_r, tot_out, tot_prof, npa_cnt) in enumerate(bands):
    fill = ALT if i % 2 == 0 else WHITE
    val(ws, R, 1, label,    fill=fill, align=LEFT)
    val(ws, R, 2, cnt,      "#,##0",    fill=fill)
    val(ws, R, 3, avg_r,    "0.00%",    fill=fill)
    val(ws, R, 4, tot_out,  "#,##0.00", fill=fill)
    val(ws, R, 5, tot_prof, "#,##0.00", fill=fill)
    val(ws, R, 6, npa_cnt,  "#,##0",    fill=fill)
    R += 1

R += 1

# ── SECTION 6: Profitability by Credit Rating ──────────────────────────────────
section_title(ws, R, "  PROFITABILITY BY CREDIT RATING")
R += 1

for ci, h in enumerate(["Credit Rating", "Count", "Avg Interest Rate",
                         "Avg Monthly Profit (₹)", "Total Monthly Profit (₹)", "NPA Count"], 1):
    hdr(ws, R, ci, h, fill=BLUE, font=SUB_FONT)
ws.row_dimensions[R].height = 28
R += 1

for i, rating in enumerate(RATINGS):
    fill = ALT if i % 2 == 0 else WHITE
    val(ws, R, 1, rating,                                                              fill=fill, align=LEFT)
    val(ws, R, 2, f'=COUNTIF({RAT},"{rating}")',               "#,##0",               fill=fill)
    val(ws, R, 3, f'=AVERAGEIF({RAT},"{rating}",{RATE})',      "0.00%",               fill=fill)
    val(ws, R, 4, f'=AVERAGEIF({RAT},"{rating}",{PROF})',      "#,##0.00",            fill=fill)
    val(ws, R, 5, f'=SUMIF({RAT},"{rating}",{PROF})',          "#,##0.00",            fill=fill)
    val(ws, R, 6, f'=COUNTIFS({RAT},"{rating}",{STA},"NPA")', "#,##0",               fill=fill)
    R += 1

# ── 5. SAVE ───────────────────────────────────────────────────────────────────
out_path = os.path.join(OUT_DIR, "loan_portfolio_dashboard.xlsx")
wb.save(out_path)
print(f"Dashboard saved: {out_path}")
