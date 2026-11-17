"""
STEP 3 - Build an Excel workbook that recomputes the key results with visible formulas.

Input : outputs/study_events.csv   (made by step2_analysis.py)
Output: outputs/Recall_Analysis_Calculations.xlsx
Run   : python step3_excel.py      (a few seconds; open the file in Excel afterwards)

Every result cell is an Excel formula, so you can click any number and see how it was calculated.
Excel calculates the formulas when you open the file.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

ev = pd.read_csv("outputs/study_events.csv", dtype={"event_id": str})
ev["quarter"] = pd.PeriodIndex(pd.to_datetime(ev["date"], errors="coerce"), freq="Q").astype(str)

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
WHITE = Font(name=F, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
NOTE = Font(name=F, size=9, italic=True, color="555555")
HEAD = PatternFill("solid", fgColor="1F3864")
KEY = PatternFill("solid", fgColor="FFFF00")
side = Side(style="thin", color="C9C9C9")
BOX = Border(top=side, bottom=side, left=side, right=side)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")

wb = Workbook()


def header(ws, row, labels, col=1):
    for j, t in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=t)
        c.font, c.fill, c.alignment, c.border = WHITE, HEAD, WRAP, BOX
    ws.row_dimensions[row].height = 36


def cell(ws, ref, value, font=BLACK, fmt=None, fill=None):
    c = ws[ref]
    c.value, c.font, c.border = value, font, BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill


def top(ws, title, sub):
    ws["A1"], ws["A2"] = title, sub
    ws["A1"].font, ws["A2"].font = TITLE, NOTE


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# ---------------------------------------------------------------- Read Me
ws = wb.active
ws.title = "Read Me"
top(ws, "Medical Device Recall Prevention: Calculations Workbook",
    "Recomputes the project's results from 13,410 FDA recall events (2013-2025) with visible Excel formulas.")
header(ws, 4, ["Sheet", "What it shows"])
for i, (a, b) in enumerate([
    ("Settings", "Study periods used by every formula (blue = editable)."),
    ("Events", "The 13,410 recall events (data from step 2). One row = one recall event."),
    ("1 Trend", "Class I events and share by year; z-test comparing 2016-2018 with 2023-2025."),
    ("2 Pareto", "Recall events by root-cause category, share and cumulative share."),
    ("3 Chi-square", "Is severity (Class I) linked to root cause? Observed, expected, chi-square, p-value."),
    ("4 Severity & Trend", "Class I rate per category with 95% CI; baseline vs recent per year; growth."),
    ("5 Priority (RPN)", "FMEA-style risk priority number: Occurrence x Severity x Trend, each scaled 1-10."),
    ("6 p-chart", "Quarterly share of Class I events with control limits set on 2013-2018."),
    ("7 Radiology", "The same Pareto for imaging (radiology) devices."),
    ("8 Impact", "How many Class I recalls a year would be avoided if top-priority causes were reduced."),
    ("9 Logistic regression", "Is root cause still linked to Class I after accounting for device class, specialty and period? Odds ratios, model checks and a what-if calculator."),
], start=5):
    cell(ws, f"A{i}", a, BOLD)
    cell(ws, f"B{i}", b)
ws["A17"], ws["A17"].font = "Colour key: blue text = data or editable input; black = formula; yellow = headline result.", NOTE
widths(ws, {"A": 22, "B": 95})

# ---------------------------------------------------------------- Settings
ws = wb.create_sheet("Settings")
top(ws, "Settings", "Change a blue value and every sheet recalculates.")
header(ws, 4, ["Setting", "Value", "Name used in formulas"])
for i, (a, v, n) in enumerate([
    ("Baseline period: first year", 2016, "Base_Start"), ("Baseline period: last year", 2018, "Base_End"),
    ("Recent period: first year", 2023, "Recent_Start"), ("Recent period: last year", 2025, "Recent_End"),
    ("Control-chart Phase I: last year", 2018, "Phase1_End"), ("z for 95% confidence", 1.96, "Z"),
], start=5):
    cell(ws, f"A{i}", a)
    cell(ws, f"B{i}", v, BLUE)
    cell(ws, f"C{i}", n, BOLD)
    wb.defined_names[n] = DefinedName(n, attr_text=f"Settings!$B${i}")
widths(ws, {"A": 34, "B": 10, "C": 22})

# ---------------------------------------------------------------- Events (data)
ws = wb.create_sheet("Events")
top(ws, "Recall events, 2013-2025", "Data from step2_analysis.py. Class I = 1 if the FDA classified the recall Class I.")
cols = ["event_id", "year", "quarter", "date", "firm", "root_cause", "category", "specialty", "products", "class1",
        "devclass", "spec_model", "period"]
header(ws, 4, cols + ["logit (b0 + b's)", "predicted P(Class I)", "log-likelihood part"])
E0 = 5
E1 = E0 + len(ev) - 1
LKP = "VLOOKUP(\"{v}|\"&{c}{r},lr_tab,4,0)"
for i, rec in enumerate(ev[cols].itertuples(index=False), start=E0):
    for j, v in enumerate(rec, start=1):
        ws.cell(row=i, column=j, value=v).font = BLUE
    terms = "+".join(LKP.format(v=v, c=c, r=i) for v, c in
                     (("Root cause", "G"), ("Device class", "K"), ("Specialty", "L"), ("Period", "M")))
    ws.cell(row=i, column=14, value=f'=IF(G{i}="Not determined","",lr_b0+{terms})').number_format = "0.000"
    ws.cell(row=i, column=15, value=f'=IF(N{i}="","",1/(1+EXP(-N{i})))').number_format = "0.000"
    ws.cell(row=i, column=16, value=f'=IF(N{i}="","",J{i}*LN(O{i})+(1-J{i})*LN(1-O{i}))').number_format = "0.000"
ws.freeze_panes = "A5"
widths(ws, {"A": 9, "B": 7, "C": 9, "D": 11, "E": 30, "F": 34, "G": 20, "H": 22, "I": 9, "J": 8,
            "K": 10, "L": 22, "M": 11, "N": 12, "O": 12, "P": 12})
YR, QT, CAT, SPEC, C1 = (f"Events!$B${E0}:$B${E1}", f"Events!$C${E0}:$C${E1}", f"Events!$G${E0}:$G${E1}",
                         f"Events!$H${E0}:$H${E1}", f"Events!$J${E0}:$J${E1}")
for nm, col in (("ev_devclass", "K"), ("ev_specm", "L"), ("ev_period", "M"), ("ev_phat", "O"), ("ev_ll", "P")):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"Events!${col}${E0}:${col}${E1}")
for nm, rng in (("ev_year", YR), ("ev_quarter", QT), ("ev_cat", CAT), ("ev_spec", SPEC), ("ev_class1", C1)):
    wb.defined_names[nm] = DefinedName(nm, attr_text=rng)

# ---------------------------------------------------------------- 1 Trend
ws = wb.create_sheet("1 Trend")
top(ws, "1. The problem: Class I recall events are rising",
    "Events = COUNTIFS(ev_year, year). Class I = SUMIFS(ev_class1, ev_year, year).")
header(ws, 4, ["Year", "Recall events", "Class I events", "Class I share"])
for k, yr in enumerate(range(2013, 2026), start=5):
    cell(ws, f"A{k}", yr, BLUE)
    cell(ws, f"B{k}", f"=COUNTIFS(ev_year,A{k})", fmt="#,##0")
    cell(ws, f"C{k}", f"=SUMIFS(ev_class1,ev_year,A{k})", fmt="#,##0")
    cell(ws, f"D{k}", f"=C{k}/B{k}", fmt="0.0%")
cell(ws, "A19", "Class I change 2020 to 2025", BOLD)
cell(ws, "C19", "=C17/C12-1", fmt="0%", fill=KEY)
cell(ws, "A20", "Recall events change 2020 to 2025", BOLD)
cell(ws, "C20", "=B17/B12-1", fmt="0%")

header(ws, 22, ["Two-proportion z-test", "Baseline", "Recent"])
rows = [("Events", "=COUNTIFS(ev_year,\">=\"&Base_Start,ev_year,\"<=\"&Base_End)",
         "=COUNTIFS(ev_year,\">=\"&Recent_Start,ev_year,\"<=\"&Recent_End)", "#,##0"),
        ("Class I events", "=SUMIFS(ev_class1,ev_year,\">=\"&Base_Start,ev_year,\"<=\"&Base_End)",
         "=SUMIFS(ev_class1,ev_year,\">=\"&Recent_Start,ev_year,\"<=\"&Recent_End)", "#,##0"),
        ("Class I share", "=B24/B23", "=C24/C23", "0.0%")]
for k, (a, b, c, fmt) in enumerate(rows, start=23):
    cell(ws, f"A{k}", a, BOLD)
    cell(ws, f"B{k}", b, fmt=fmt)
    cell(ws, f"C{k}", c, fmt=fmt)
ws["C25"].fill = KEY
for k, (a, f, fmt) in enumerate([
    ("Pooled share", "=(B24+C24)/(B23+C23)", "0.000%"),
    ("z statistic", "=(C25-B25)/SQRT(B26*(1-B26)*(1/B23+1/C23))", "0.00"),
    ("p-value (two-sided)", "=2*(1-NORMSDIST(ABS(B27)))", "0.0E+00"),
], start=26):
    cell(ws, f"A{k}", a, BOLD)
    cell(ws, f"B{k}", f, fmt=fmt)
ws["B27"].fill = KEY
widths(ws, {"A": 34, "B": 15, "C": 15, "D": 14})

# ---------------------------------------------------------------- 2 Pareto
CATS = ["Production process", "Design (hardware)", "Software", "Supplier & material",
        "Labeling & packaging", "Regulatory", "Use error"]
ws = wb.create_sheet("2 Pareto")
top(ws, "2. Pareto of root causes (events with a known cause)",
    "'Not determined' events are counted separately: the firm did not find or report a root cause.")
header(ws, 4, ["Category", "Recall events", "Share", "Cumulative share", "Class I events"])
for k, c in enumerate(CATS, start=5):
    cell(ws, f"A{k}", c, BLUE)
    cell(ws, f"B{k}", f"=COUNTIFS(ev_cat,A{k})", fmt="#,##0")
    cell(ws, f"C{k}", f"=B{k}/$B$12", fmt="0.0%")
    cell(ws, f"D{k}", f"=SUM($C$5:C{k})", fmt="0.0%")
    cell(ws, f"E{k}", f"=SUMIFS(ev_class1,ev_cat,A{k})", fmt="#,##0")
cell(ws, "A12", "Total, known causes", BOLD)
cell(ws, "B12", "=SUM(B5:B11)", BOLD, "#,##0")
cell(ws, "E12", "=SUM(E5:E11)", BOLD, "#,##0")
cell(ws, "A14", "Not determined", BOLD)
cell(ws, "B14", '=COUNTIFS(ev_cat,"Not determined")', fmt="#,##0")
cell(ws, "C14", "=B14/(B12+B14)", fmt="0.0%", fill=KEY)
ws["D14"], ws["D14"].font = "share of all events with no root cause found", NOTE
widths(ws, {"A": 26, "B": 14, "C": 10, "D": 16, "E": 14})

# ---------------------------------------------------------------- 3 Chi-square
ws = wb.create_sheet("3 Chi-square")
top(ws, "3. Chi-square test: is recall severity linked to root cause?",
    "Expected = row total x column total / grand total. Contribution = (Observed - Expected)^2 / Expected.")
header(ws, 4, ["Category", "Class I (observed)", "Not Class I (observed)", "Row total",
               "Class I (expected)", "Not Class I (expected)", "Contribution: Class I", "Contribution: not Class I"])
for k, c in enumerate(CATS, start=5):
    cell(ws, f"A{k}", c, BLUE)
    cell(ws, f"B{k}", f"=SUMIFS(ev_class1,ev_cat,A{k})")
    cell(ws, f"C{k}", f"=COUNTIFS(ev_cat,A{k})-B{k}")
    cell(ws, f"D{k}", f"=B{k}+C{k}")
    cell(ws, f"E{k}", f"=D{k}*$B$12/$D$12", fmt="0.0")
    cell(ws, f"F{k}", f"=D{k}*$C$12/$D$12", fmt="0.0")
    cell(ws, f"G{k}", f"=(B{k}-E{k})^2/E{k}", fmt="0.00")
    cell(ws, f"H{k}", f"=(C{k}-F{k})^2/F{k}", fmt="0.00")
cell(ws, "A12", "Total", BOLD)
for col in "BCD":
    cell(ws, f"{col}12", f"=SUM({col}5:{col}11)", BOLD)
for k, (a, f, fmt) in enumerate([
    ("Chi-square statistic", "=SUM(G5:H11)", "0.0"),
    ("Degrees of freedom", "=(7-1)*(2-1)", "0"),
    ("p-value", "=CHIDIST(B14,B15)", "0.0E+00"),
    ("Cramer's V (strength, 0-1)", "=SQRT(B14/D12)", "0.000"),
    ("Smallest expected count (should be 5 or more)", "=MIN(E5:F11)", "0.0"),
], start=14):
    cell(ws, f"A{k}", a, BOLD)
    cell(ws, f"B{k}", f, fmt=fmt, fill=KEY if k in (14, 16) else None)
widths(ws, {"A": 42, **{L(c): 14 for c in range(2, 9)}})

# ---------------------------------------------------------------- 4 Severity & Trend
ws = wb.create_sheet("4 Severity & Trend")
top(ws, "4. Severity and trend by root-cause category",
    "Per-year values = count in the period / number of years in the period.")
header(ws, 4, ["Category", "Class I rate, 2013-2025", "95% CI ±", "Relative to average",
               "Events / year, baseline", "Events / year, recent", "Class I / year, baseline",
               "Class I / year, recent", "Class I rate, recent", "Growth in Class I / year"])
NB, NR = "(Base_End-Base_Start+1)", "(Recent_End-Recent_Start+1)"
for k, c in enumerate(CATS, start=5):
    cell(ws, f"A{k}", c, BLUE)
    cell(ws, f"B{k}", f"='2 Pareto'!E{k}/'2 Pareto'!B{k}", fmt="0.0%")
    cell(ws, f"C{k}", f"=Z*SQRT(B{k}*(1-B{k})/'2 Pareto'!B{k})", fmt="0.0%")
    cell(ws, f"D{k}", f"=B{k}/$B$13", fmt="0.00")
    cell(ws, f"E{k}", f"=COUNTIFS(ev_cat,A{k},ev_year,\">=\"&Base_Start,ev_year,\"<=\"&Base_End)/{NB}", fmt="0.0")
    cell(ws, f"F{k}", f"=COUNTIFS(ev_cat,A{k},ev_year,\">=\"&Recent_Start,ev_year,\"<=\"&Recent_End)/{NR}", fmt="0.0")
    cell(ws, f"G{k}", f"=SUMIFS(ev_class1,ev_cat,A{k},ev_year,\">=\"&Base_Start,ev_year,\"<=\"&Base_End)/{NB}", fmt="0.0")
    cell(ws, f"H{k}", f"=SUMIFS(ev_class1,ev_cat,A{k},ev_year,\">=\"&Recent_Start,ev_year,\"<=\"&Recent_End)/{NR}", fmt="0.0")
    cell(ws, f"I{k}", f"=IF(F{k}=0,0,H{k}/F{k})", fmt="0.0%")
    cell(ws, f"J{k}", f'=IF(G{k}=0,"n/a",H{k}/G{k}-1)', fmt="0%")
cell(ws, "A13", "Average Class I rate, known causes", BOLD)
cell(ws, "B13", "='2 Pareto'!E12/'2 Pareto'!B12", fmt="0.0%")
widths(ws, {"A": 34, **{L(c): 13 for c in range(2, 11)}})

# ---------------------------------------------------------------- 5 Priority (RPN)
ws = wb.create_sheet("5 Priority (RPN)")
top(ws, "5. Risk priority number (FMEA style)",
    "Score = 1 + 9 x (value - min) / (max - min). RPN = O x S x T. Regulatory is excluded: it is a paperwork issue, not a product failure.")
header(ws, 4, ["Category", "Occurrence: events / year (recent)", "Severity: Class I rate (recent)",
               "Trend: growth in Class I / year", "O score", "S score", "T score", "RPN", "Priority rank",
               "ISO 13485 clause to strengthen", "Class I events / year (recent)"])
RCATS = ["Production process", "Design (hardware)", "Software", "Supplier & material", "Labeling & packaging", "Use error"]
CLAUSE = {"Production process": "7.5 Production and service provision", "Design (hardware)": "7.3 Design and development",
          "Software": "7.3 Design and development + IEC 62304", "Supplier & material": "7.4 Purchasing",
          "Labeling & packaging": "7.5.1 Control of production", "Use error": "7.3 + IEC 62366 usability"}
src = {c: CATS.index(c) + 5 for c in RCATS}
for k, c in enumerate(RCATS, start=5):
    s = src[c]
    cell(ws, f"A{k}", c, BLUE)
    cell(ws, f"B{k}", f"='4 Severity & Trend'!F{s}", fmt="0.0")
    cell(ws, f"C{k}", f"='4 Severity & Trend'!I{s}", fmt="0.0%")
    cell(ws, f"D{k}", f"='4 Severity & Trend'!J{s}", fmt="0%")
    for col, srccol in (("E", "B"), ("F", "C"), ("G", "D")):
        cell(ws, f"{col}{k}", f"=1+9*({srccol}{k}-MIN({srccol}$5:{srccol}$10))/(MAX({srccol}$5:{srccol}$10)-MIN({srccol}$5:{srccol}$10))", fmt="0.0")
    cell(ws, f"H{k}", f"=E{k}*F{k}*G{k}", fmt="0", fill=KEY)
    cell(ws, f"I{k}", f"=RANK(H{k},$H$5:$H$10)")
    cell(ws, f"J{k}", CLAUSE[c])
    cell(ws, f"K{k}", f"='4 Severity & Trend'!H{s}", fmt="0.0")
widths(ws, {"A": 24, "B": 16, "C": 16, "D": 16, "E": 9, "F": 9, "G": 9, "H": 9, "I": 9, "J": 38, "K": 14})

# ---------------------------------------------------------------- 6 p-chart
ws = wb.create_sheet("6 p-chart")
top(ws, "6. p-chart: share of recall events classified Class I, by quarter",
    "p-bar from Phase I (2013 to Phase1_End). UCL/LCL = p-bar ± 3 x SQRT(p-bar x (1 - p-bar) / n).")
header(ws, 4, ["Quarter", "Year", "Events (n)", "Class I (d)", "Share (p)", "UCL", "LCL", "Above UCL?"])
quarters = sorted(ev["quarter"].unique())
Q0 = 5
Q1 = Q0 + len(quarters) - 1
for k, qq in enumerate(quarters, start=Q0):
    cell(ws, f"A{k}", qq, BLUE)
    cell(ws, f"B{k}", f"=VALUE(LEFT(A{k},4))")
    cell(ws, f"C{k}", f"=COUNTIFS(ev_quarter,A{k})")
    cell(ws, f"D{k}", f"=SUMIFS(ev_class1,ev_quarter,A{k})")
    cell(ws, f"E{k}", f"=D{k}/C{k}", fmt="0.0%")
    cell(ws, f"F{k}", f"=$L$5+3*SQRT($L$5*(1-$L$5)/C{k})", fmt="0.0%")
    cell(ws, f"G{k}", f"=MAX(0,$L$5-3*SQRT($L$5*(1-$L$5)/C{k}))", fmt="0.0%")
    cell(ws, f"H{k}", f'=IF(E{k}>F{k},"Yes","")')
cell(ws, "K5", "Center line p-bar (Phase I)", BOLD)
cell(ws, "L5", f'=SUMIFS(D{Q0}:D{Q1},B{Q0}:B{Q1},"<="&Phase1_End)/SUMIFS(C{Q0}:C{Q1},B{Q0}:B{Q1},"<="&Phase1_End)', fmt="0.00%", fill=KEY)
cell(ws, "K6", "Quarters above UCL (all)", BOLD)
cell(ws, "L6", f'=COUNTIF(H{Q0}:H{Q1},"Yes")')
cell(ws, "K7", "Quarters above UCL, 2023-2025", BOLD)
cell(ws, "L7", f'=COUNTIFS(H{Q0}:H{Q1},"Yes",B{Q0}:B{Q1},">=2023")', fill=KEY)
cell(ws, "K8", "Quarters in 2023-2025", BOLD)
cell(ws, "L8", f'=COUNTIFS(B{Q0}:B{Q1},">=2023")')
ch = LineChart()
ch.title = "p-chart: Class I share of recall events"
ch.y_axis.title, ch.x_axis.title = "Share", "Quarter"
ch.height, ch.width = 9, 24
for col in (5, 6):
    ch.add_data(Reference(ws, min_col=col, min_row=4, max_row=Q1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=Q0, max_row=Q1))
ch.series[0].graphicalProperties.line.solidFill = "2A6FDB"
ch.series[1].graphicalProperties.line.solidFill = "C00000"
ch.series[1].graphicalProperties.line.dashStyle = "dash"
ws.add_chart(ch, "K11")
widths(ws, {"A": 9, "B": 7, "C": 10, "D": 10, "E": 9, "F": 9, "G": 9, "H": 10, "K": 30, "L": 10})

# ---------------------------------------------------------------- 7 Radiology
ws = wb.create_sheet("7 Radiology")
top(ws, "7. Imaging devices (radiology specialty): Pareto of root causes", "COUNTIFS on category and specialty = Radiology.")
header(ws, 4, ["Category", "Recall events", "Share", "Cumulative share", "Class I events", "Class I rate"])
RAD = ["Software", "Design (hardware)", "Production process", "Supplier & material", "Regulatory",
       "Labeling & packaging", "Use error"]
for k, c in enumerate(RAD, start=5):
    cell(ws, f"A{k}", c, BLUE)
    cell(ws, f"B{k}", f'=COUNTIFS(ev_cat,A{k},ev_spec,"Radiology")', fmt="#,##0")
    cell(ws, f"C{k}", f"=B{k}/$B$12", fmt="0.0%", fill=KEY if k == 5 else None)
    cell(ws, f"D{k}", f"=SUM($C$5:C{k})", fmt="0.0%")
    cell(ws, f"E{k}", f'=SUMIFS(ev_class1,ev_cat,A{k},ev_spec,"Radiology")')
    cell(ws, f"F{k}", f"=IF(B{k}=0,0,E{k}/B{k})", fmt="0.0%")
cell(ws, "A12", "Total", BOLD)
cell(ws, "B12", "=SUM(B5:B11)", BOLD, "#,##0")
cell(ws, "E12", "=SUM(E5:E11)", BOLD)
widths(ws, {"A": 24, "B": 14, "C": 10, "D": 16, "E": 14, "F": 12})

# ---------------------------------------------------------------- 8 Impact
ws = wb.create_sheet("8 Impact")
top(ws, "8. Impact scenario", "Class I recall events avoided per year if the two top-priority categories were reduced. A scenario, not a forecast.")
cell(ws, "A4", "Class I events per year, recent, top 2 priorities", BOLD)
cell(ws, "B4", "=SUMIFS('5 Priority (RPN)'!K5:K10,'5 Priority (RPN)'!I5:I10,\"<=2\")", fmt="0.0")
cell(ws, "A5", "Class I events per year, recent, all known causes", BOLD)
cell(ws, "B5", "=SUM('4 Severity & Trend'!H5:H11)", fmt="0.0")
cell(ws, "A6", "Top 2 share of Class I events", BOLD)
cell(ws, "B6", "=B4/B5", fmt="0%", fill=KEY)
header(ws, 8, ["Reduction in top-2 categories", "Class I events avoided per year", "Share of all known-cause Class I"])
for k, cut in enumerate((0.10, 0.25, 0.50), start=9):
    cell(ws, f"A{k}", cut, BLUE, "0%")
    cell(ws, f"B{k}", f"=$B$4*A{k}", fmt="0.0", fill=KEY if cut == 0.25 else None)
    cell(ws, f"C{k}", f"=B{k}/$B$5", fmt="0.0%")
widths(ws, {"A": 46, "B": 26, "C": 30})


# ---------------------------------------------------------------- 9 Logistic regression
import json
from openpyxl.worksheet.datavalidation import DataValidation
lt = pd.read_csv("outputs/table_logistic_regression.csv")
LR = json.load(open("outputs/results.json"))["logistic_regression"]
ws = wb.create_sheet("9 Logistic regression")
top(ws, "9. Logistic regression: Class I ~ root cause + device class + specialty + period",
    "Coefficients (blue) were estimated in Python (statsmodels). Every other number is an Excel formula, and the log-likelihood check below re-derives the fit from the 10,640 events.")
header(ws, 4, ["Lookup key", "Variable", "Level", "Coefficient b", "Std error", "z = b / SE", "p-value",
               "Odds ratio = EXP(b)", "OR 95% low", "OR 95% high", "Events", "Class I observed",
               "Class I predicted", "p-value, firm-clustered SE"])
RNG = {"Root cause": "ev_cat", "Device class": "ev_devclass", "Specialty": "ev_specm", "Period": "ev_period"}
T0 = 5
levels_rows = {}
for k, r in enumerate(lt.itertuples(index=False), start=T0):
    lev = r.level.replace(" (reference)", "")
    is_ref = "reference" in r.level
    key = "Intercept" if r.variable == "Intercept" else f"{r.variable}|{lev}"
    cell(ws, f"A{k}", key, NOTE)
    cell(ws, f"B{k}", r.variable, BOLD)
    cell(ws, f"C{k}", lev + ("  (reference)" if is_ref else ""), BLUE if not is_ref else BOLD)
    if not is_ref:
        ws[f"C{k}"].value = lev
    cell(ws, f"D{k}", round(float(r.coef), 6), BLUE, "0.000")
    if is_ref:
        cell(ws, f"E{k}", None)
        cell(ws, f"F{k}", "reference", NOTE)
    else:
        cell(ws, f"E{k}", round(float(r.se), 6), BLUE, "0.000")
        cell(ws, f"F{k}", f"=D{k}/E{k}", fmt="0.00")
        cell(ws, f"G{k}", f"=2*(1-NORMSDIST(ABS(F{k})))", fmt="0.0000")
    if r.variable != "Intercept":
        cell(ws, f"H{k}", f"=EXP(D{k})", fmt="0.00", fill=KEY if lev in ("Design (hardware)", "2023-2025") else None)
        if not is_ref:
            cell(ws, f"I{k}", f"=EXP(D{k}-Z*E{k})", fmt="0.00")
            cell(ws, f"J{k}", f"=EXP(D{k}+Z*E{k})", fmt="0.00")
            cell(ws, f"N{k}", round(float(r.p_value_firm_clustered), 4), BLUE, "0.0000")
        rg = RNG[r.variable]
        cell(ws, f"K{k}", f'=COUNTIFS({rg},"{lev}",ev_cat,"<>Not determined")', fmt="#,##0")
        cell(ws, f"L{k}", f'=SUMIFS(ev_class1,{rg},"{lev}",ev_cat,"<>Not determined")', fmt="#,##0")
        cell(ws, f"M{k}", f'=SUMIFS(ev_phat,{rg},"{lev}")', fmt="0.0")
        levels_rows.setdefault(r.variable, []).append(k)
    else:
        cell(ws, f"K{k}", '=COUNTIFS(ev_cat,"<>Not determined")', fmt="#,##0")
        cell(ws, f"L{k}", '=SUMIFS(ev_class1,ev_cat,"<>Not determined")', fmt="#,##0")
        cell(ws, f"M{k}", "=SUM(ev_phat)", fmt="0.0")
        wb.defined_names["lr_b0"] = DefinedName("lr_b0", attr_text=f"'9 Logistic regression'!$D${k}")
T1 = T0 + len(lt) - 1
wb.defined_names["lr_tab"] = DefinedName("lr_tab", attr_text=f"'9 Logistic regression'!$A${T0}:$D${T1}")

# Model fit block
F0 = T1 + 3
header(ws, F0, ["", "Model check", "", "Excel formula", "Python result", "Difference"])
fit = [
    ("Events in the model (known root cause)", f"=K{T0}", LR["events"], "#,##0"),
    ("Class I events", f"=L{T0}", LR["class1"], "#,##0"),
    ("Average Class I share p-bar", f"=D{F0+2}/D{F0+1}", None, "0.00%"),
    ("Log-likelihood of the model = SUM of Events column P", "=SUM(ev_ll)", LR["log_likelihood"], "0.000"),
    ("Log-likelihood of the null model (no predictors)", f"=D{F0+2}*LN(D{F0+3})+(D{F0+1}-D{F0+2})*LN(1-D{F0+3})", LR["log_likelihood_null"], "0.000"),
    ("Likelihood-ratio chi-square = 2 x (LL model - LL null)", f"=2*(D{F0+4}-D{F0+5})", LR["lr_test_all"]["chi2"], "0.0"),
    ("Degrees of freedom = number of non-reference coefficients", f"=COUNT(E{T0+1}:E{T1})", LR["lr_test_all"]["df"], "0"),
    ("p-value of the whole model = CHIDIST(chi-square, df)", f"=CHIDIST(D{F0+6},D{F0+7})", None, "0.0E+00"),
    ("McFadden pseudo R-squared = 1 - LL model / LL null", f"=1-D{F0+4}/D{F0+5}", LR["mcfadden_r2"], "0.000"),
    ("Sum of predicted probabilities (should equal Class I events)", f"=M{T0}", LR["class1"], "0.0"),
    ("Does root cause matter after the controls? LR chi-square (Python: model with vs without root cause)", None, LR["lr_test_root_cause"]["chi2"], "0.0"),
    ("   its p-value = CHIDIST(chi-square, 6)", f"=CHIDIST(E{F0+11},{LR['lr_test_root_cause']['df']})", None, "0.0E+00"),
    ("AUC: chance the model ranks a random Class I event above a random non-Class I event (Python)", None, LR["auc"], "0.000"),
]
for i, (lab, f, py, fmt) in enumerate(fit, start=F0 + 1):
    cell(ws, f"B{i}", lab, BOLD)
    if f:
        cell(ws, f"D{i}", f, fmt=fmt, fill=KEY if i in (F0 + 4, F0 + 9) else None)
    if py is not None:
        cell(ws, f"E{i}", py, BLUE, fmt)
    if f and py is not None:
        cell(ws, f"F{i}", f"=D{i}-E{i}", fmt="0.000")

# Unadjusted vs adjusted odds ratios for root cause
U0 = F0 + len(fit) + 3
header(ws, U0, ["", "Root cause", "Class I rate", "Odds = rate / (1 - rate)", "Unadjusted OR (vs Production)",
                "Adjusted OR (from the model)", "What changed"])
rc_rows = levels_rows["Root cause"]
ref_row = U0 + 1
for i, k in enumerate(rc_rows, start=U0 + 1):
    cell(ws, f"B{i}", f"=SUBSTITUTE(C{k},\"  (reference)\",\"\")", BOLD)
    cell(ws, f"C{i}", f"=L{k}/K{k}", fmt="0.0%")
    cell(ws, f"D{i}", f"=C{i}/(1-C{i})", fmt="0.000")
    cell(ws, f"E{i}", f"=D{i}/$D${ref_row}", fmt="0.00")
    cell(ws, f"F{i}", f"=H{k}", fmt="0.00")
    if i == ref_row:
        cell(ws, f"G{i}", "reference")
    else:
        cell(ws, f"G{i}", f'=IF(ABS(ABS(LN(F{i}))-ABS(LN(E{i})))<0.05,"about the same",IF(ABS(LN(F{i}))>ABS(LN(E{i})),"further from 1: stronger after adjustment","closer to 1: weaker after adjustment"))')

# What-if calculator
W0 = U0 + len(rc_rows) + 3
header(ws, W0, ["", "What-if calculator", "Choose (drop-down)", "Coefficient used"])
choices = [("Root cause", "Design (hardware)"), ("Device class", "Class 2"),
           ("Specialty", "General Hospital"), ("Period", "2023-2025")]
for i, (var, default) in enumerate(choices, start=W0 + 1):
    cell(ws, f"B{i}", var, BOLD)
    cell(ws, f"C{i}", default, BLUE)
    cell(ws, f"D{i}", f'=VLOOKUP(B{i}&"|"&C{i},lr_tab,4,0)', fmt="0.000")
    rows_ = levels_rows[var]
    # drop-down uses the level names; reference levels carry a label, so list a clean copy in column P
    for j, k in enumerate(rows_):
        ws[f"{L(16 + i - W0 - 1)}{T0 + j}"] = f'=SUBSTITUTE(C{k},"  (reference)","")'
    dv = DataValidation(type="list", formula1=f"=${L(16 + i - W0 - 1)}${T0}:${L(16 + i - W0 - 1)}${T0 + len(rows_) - 1}", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"C{i}")
cell(ws, f"B{W0+5}", "Intercept b0", BOLD)
cell(ws, f"D{W0+5}", "=lr_b0", fmt="0.000")
cell(ws, f"B{W0+6}", "logit = b0 + the four coefficients", BOLD)
cell(ws, f"D{W0+6}", f"=SUM(D{W0+1}:D{W0+5})", fmt="0.000")
cell(ws, f"B{W0+7}", "Predicted probability of Class I = 1 / (1 + EXP(-logit))", BOLD)
cell(ws, f"D{W0+7}", f"=1/(1+EXP(-D{W0+6}))", fmt="0.0%", fill=KEY)
for j, h in enumerate(["Root cause levels", "Device class levels", "Specialty levels", "Period levels"]):
    c_ = ws.cell(row=T0 - 1, column=16 + j, value=h)
    c_.font, c_.fill, c_.alignment = WHITE, HEAD, WRAP
widths(ws, {"A": 12, "B": 34, "C": 28, "D": 14, "E": 12, "F": 12, "G": 12, "H": 13, "I": 11, "J": 11,
            "K": 10, "L": 11, "M": 11, "N": 14, "P": 20, "Q": 14, "R": 24, "S": 12})
LR_ROWS = dict(T0=T0, T1=T1, F0=F0, U0=U0, W0=W0)

# ---------------------------------------------------------------- Plain-English explanations on every sheet
# Each sheet gets a "HOW THIS SHEET WORKS" panel: what the sheet does, then every column or cell
# with its formula written out in words, for a reader who is new to the subject.
PANEL_HEAD = Font(name=F, size=11, bold=True, color="FFFFFF")
PANEL_FILL = PatternFill("solid", fgColor="2A6FDB")
PANEL_BODY = PatternFill("solid", fgColor="EEF3FC")
PANEL_TEXT = Font(name=F, size=10)
PANEL_SUB = Font(name=F, size=10, bold=True, color="1F3864")


def explain(sheet, col, row, lines, width=95):
    """Write an explanation panel starting at (col, row). Lines starting with '#' are sub-headings."""
    ws = wb[sheet]
    ws.column_dimensions[col].width = width
    c = ws[f"{col}{row}"]
    c.value, c.font, c.fill = "HOW THIS SHEET WORKS (plain English)", PANEL_HEAD, PANEL_FILL
    for i, text in enumerate(lines, start=row + 1):
        c = ws[f"{col}{i}"]
        sub = text.startswith("#")
        c.value = text.lstrip("# ")
        c.font = PANEL_SUB if sub else PANEL_TEXT
        c.fill = PANEL_BODY
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = max(15, 15 * (len(c.value) // int(width * 1.1) + 1))


explain("Read Me", "B", 19, [
    "# What this workbook is",
    "It repeats the project's main calculations in Excel, using real formulas. Click any black number and the formula bar shows how it was calculated.",
    "The data is 13,410 FDA medical device recall events from 2013 to 2025. One event = one problem that made a company fix or take back a device.",
    "# Words used in the sheets",
    "Class I: the FDA's most serious recall rating (could cause serious injury or death). In the data, class1 = 1 means Class I, 0 means Class II or III.",
    "Root cause: the reason the company found for the problem. About 40 FDA labels are grouped into 8 areas (column 'category' in Events).",
    "Baseline period: 2016-2018 (normal years before COVID). Recent period: 2023-2025 (latest full years).",
    "# Names used in formulas",
    "To keep formulas readable, some ranges have names. ev_year = the year column in Events; ev_cat = the root-cause area; ev_class1 = 1 if Class I; ev_spec = medical specialty; ev_quarter = quarter (like 2024Q3).",
    "Base_Start, Base_End, Recent_Start, Recent_End, Phase1_End and Z are single cells on the Settings sheet.",
    "# The two Excel functions used most",
    "COUNTIFS(range, condition, ...) counts the rows that meet every condition. Example: =COUNTIFS(ev_cat,\"Software\") counts software recall events.",
    "SUMIFS(sum_range, range, condition, ...) adds up values in the rows that meet the conditions. Because class1 is 1 or 0, =SUMIFS(ev_class1, ev_cat, \"Software\") counts software events that were Class I.",
    "# Suggested reading order",
    "Settings -> Events -> 1 Trend -> 2 Pareto -> 3 Chi-square -> 4 Severity & Trend -> 5 Priority (RPN) -> 6 p-chart -> 7 Radiology -> 8 Impact -> 9 Logistic regression.",
], width=95)

explain("Settings", "E", 4, [
    "These are the only numbers you would normally change. Every other sheet reads them through their names (column C).",
    "# Rows 5-6: Base_Start and Base_End (2016, 2018)",
    "The baseline period: three normal years before COVID, used as 'before' when measuring change.",
    "# Rows 7-8: Recent_Start and Recent_End (2023, 2025)",
    "The recent period: the latest three full years, used as 'after'. 2026 is not used because many 2026 recalls are not yet classified by the FDA.",
    "# Row 9: Phase1_End (2018)",
    "The control chart (sheet 6) learns what 'normal' looks like from 2013 up to this year, then checks later quarters against it.",
    "# Row 10: Z (1.96)",
    "The number of standard errors that gives 95% confidence. Used for the confidence intervals in sheet 4.",
    "# Try it",
    "Change Recent_Start to 2022 and watch sheets 1, 4, 5 and 8 recalculate.",
], width=80)

explain("Events", "R", 4, [
    "This is the data. Each row is one recall event (one problem), made by step2_analysis.py from the two FDA files. Nothing on this sheet is a formula.",
    "# Columns",
    "event_id: the FDA's ID for the recall event.",
    "year, quarter, date: when the company started the recall.",
    "firm: the company that recalled the device.",
    "root_cause: the FDA root-cause label the company reported, e.g. 'Device Design'.",
    "category: the area the root cause was grouped into (8 areas, e.g. 'Design (hardware)', 'Production process'). 'Not determined' = no cause found.",
    "specialty: the medical field of the device, e.g. Radiology (imaging).",
    "products: how many recalled-product records belong to this event.",
    "class1: 1 if the FDA classified the event Class I (any of its products), otherwise 0.",
    "devclass: the FDA device class (1 = low risk, 2 = moderate, 3 = high risk; the highest among the event's products). Used by the regression.",
    "spec_model: specialty, with the 8 smallest specialties (under 200 known-cause events) pooled into 'Other specialties' so the regression has enough data for each group.",
    "period: 2013-2018, 2019-2022 or 2023-2025.",
    "# Columns N-P (formulas, sheet 9)",
    "N, logit: intercept + the coefficient for this event's root cause, device class, specialty and period, each found with VLOOKUP in the sheet 9 table. Blank for 'Not determined' events (not in the model).",
    "O, predicted P(Class I) = 1/(1+EXP(-logit)): the model's probability that this event is Class I.",
    "P, log-likelihood part = class1 x LN(p) + (1 - class1) x LN(1 - p): how well the model predicted this one event (closer to 0 = better). Their sum is the model's log-likelihood.",
    "# Why events, not products",
    "One problem can create hundreds of product records (the largest event has 233). Counting events stops a few large recalls from dominating the results.",
], width=80)

explain("1 Trend", "F", 4, [
    "Question: are the most serious recalls really increasing?",
    "# Columns B to D (rows 5-17), one row per year",
    "B, Recall events: =COUNTIFS(ev_year, A5) counts all events whose year equals the year in column A.",
    "C, Class I events: =SUMIFS(ev_class1, ev_year, A5) adds up the 1s, which counts the Class I events that year.",
    "D, Class I share: =C5/B5 is the fraction of that year's events that were Class I.",
    "# Rows 19-20: change from 2020 to 2025",
    "C19 = C17/C12 - 1 compares 2025 (row 17) with 2020 (row 12). 125 / 37 - 1 = 2.38, i.e. +238%. The FDA itself reported +232%.",
    "C20 does the same for all recalls: total recalls did not rise, only serious ones did.",
    "# Rows 22-28: two-proportion z-test",
    "It asks: is the jump in the Class I share from the baseline to the recent period bigger than chance could explain?",
    "Row 23, Events: COUNTIFS with two conditions, year >= start and year <= end, counts events in each period.",
    "Row 24, Class I events: the same with SUMIFS on ev_class1.",
    "Row 25, Class I share: Class I events / events. Baseline 3.5%, recent 11.5%.",
    "Row 26, Pooled share: both periods combined = (B24+C24)/(B23+C23). This is the share we would expect in both periods if nothing had changed.",
    "Row 27, z statistic: (recent share - baseline share) / SQRT(pooled x (1 - pooled) x (1/n1 + 1/n2)). It measures how many 'standard errors' apart the two shares are.",
    "Row 28, p-value: =2*(1-NORMSDIST(ABS(z))), the chance of seeing a gap this large if nothing had really changed. Rule of thumb: z above 1.96 (p below 0.05) means a real change. Here z = 12.16, so p is effectively 0.",
], width=95)

explain("2 Pareto", "G", 4, [
    "Question: which root-cause areas cause the most recalls? (Pareto analysis, the 80/20 rule.)",
    "# Columns (rows 5-11), areas listed from most to fewest events",
    "B, Recall events: =COUNTIFS(ev_cat, A5) counts events in that area.",
    "C, Share: =B5/$B$12 divides by the total of known-cause events (B12). The $ signs keep B12 fixed when the formula is copied down.",
    "D, Cumulative share: =SUM($C$5:C5) adds the shares from the top down to this row. Where it passes about 80% marks the 'vital few' causes.",
    "E, Class I events: =SUMIFS(ev_class1, ev_cat, A5) counts the serious recalls in each area.",
    "# Rows 12-14",
    "B12 = SUM(B5:B11): all events with a known cause (10,640).",
    "B14 counts 'Not determined' events (no root cause found). C14 = B14/(B12+B14) shows they are 20.7% of all events. They are kept out of the Pareto because you cannot prevent a cause you do not know.",
    "# What it shows",
    "Production process, hardware design, software and supplier & material together cover about 80% of known-cause recalls.",
], width=90)

explain("3 Chi-square", "J", 4, [
    "Question: does the chance of a recall being Class I depend on its root cause?",
    "Idea: compare the Class I count each area actually had (observed) with the count we would expect if danger had nothing to do with the cause (expected).",
    "# Columns (rows 5-11)",
    "B, Class I (observed): =SUMIFS(ev_class1, ev_cat, A5).",
    "C, Not Class I (observed): all events in the area minus the Class I ones.",
    "D, Row total: B + C.",
    "E, Class I (expected): =D5*$B$12/$D$12, i.e. the area's events x the overall Class I share. If design has 2,195 events and 6.7% of all events are Class I, we would expect about 148 Class I design recalls.",
    "F, Not Class I (expected): the same for the non-Class I column.",
    "G and H, Contribution: =(Observed - Expected)^2 / Expected. A big number means that area is far from what chance would give.",
    "# Rows 14-18",
    "B14, Chi-square statistic: the sum of all contributions (177.5). The bigger, the stronger the evidence of a link.",
    "B15, Degrees of freedom: (7 areas - 1) x (2 columns - 1) = 6.",
    "B16, p-value: =CHIDIST(chi-square, df), the chance of a chi-square this large if there were no link. Below 0.05 = a real link. Here it is about 1e-35.",
    "B17, Cramer's V: =SQRT(chi-square / total events), the strength of the link on a 0 to 1 scale. 0.13 = small but real.",
    "B18, Smallest expected count: the test is reliable when every expected count is at least 5. The smallest here is 9.0.",
    "# What it shows",
    "Hardware design has far more Class I recalls than expected (11.9% vs 6.7% average); labeling and packaging has far fewer (1.9%).",
], width=95)

explain("4 Severity & Trend", "L", 4, [
    "Question: how dangerous is each area, and is it getting worse?",
    "# Severity columns",
    "B, Class I rate 2013-2025: Class I events / all events in the area, taken from sheet 2.",
    "C, 95% CI +/-: =Z*SQRT(rate*(1-rate)/events). The margin of error on the rate; the true rate is likely within rate +/- this value. Small areas have wider margins.",
    "D, Relative to average: rate / the average rate in B13. 1.77 for design means design recalls are 1.77 times as likely as average to be Class I.",
    "# Trend columns",
    "E and F, Events per year: COUNTIFS for the area within the baseline or recent years, divided by the number of years (3). Per-year averages make the two periods comparable.",
    "G and H, Class I per year: the same with SUMIFS on ev_class1.",
    "I, Class I rate (recent): H / F, the share of recent events that were Class I.",
    "J, Growth in Class I per year: =H/G - 1, recent vs baseline. 3.25 for software means +325% (about 2.7 to 11.3 a year). 'n/a' means there were no baseline Class I events, so growth cannot be computed.",
    "# What it shows",
    "Every major area has become more severe. Design's Class I rate rose from 7.4% to 20.2%; software Class I recalls grew fastest.",
], width=110)

explain("5 Priority (RPN)", "M", 4, [
    "Question: putting frequency, danger and trend together, where should a manufacturer act first?",
    "Method: an FMEA-style risk priority number (RPN). In a classic FMEA, three 1-10 scores are multiplied. Here all three scores come from the data.",
    "# Raw values (columns B-D, recent period 2023-2025)",
    "B, Occurrence: recall events per year (from sheet 4, column F).",
    "C, Severity: Class I rate (sheet 4, column I).",
    "D, Trend: growth in Class I events per year vs 2016-2018 (sheet 4, column J).",
    "# Scores (columns E-G)",
    "Each raw value is rescaled to 1-10 with min-max scaling: =1+9*(value - MIN(column)) / (MAX(column) - MIN(column)). The lowest area gets 1, the highest gets 10, the rest fall in between. This puts events, percentages and growth on the same scale so they can be multiplied fairly.",
    "# Result columns",
    "H, RPN: =E*F*G. Higher = act sooner.",
    "I, Priority rank: =RANK(H, all RPNs), 1 = highest.",
    "J, ISO 13485 clause: the part of the quality standard that controls this area, which is where the fix belongs.",
    "K, Class I events per year (recent): used by the Impact sheet.",
    "# Notes",
    "Regulatory recalls (missing FDA paperwork) are left out because they are not product failures.",
    "Result: production process (232), hardware design (182) and software (152) are the top three.",
], width=110)

explain("6 p-chart", "K", 31, [
    "Question: is the rise in serious recalls a real change in the process, or normal up-and-down variation?",
    "Method: a p-chart, the control chart for percentages, used in Six Sigma and statistical process control.",
    "# Columns (one row per quarter)",
    "A, Quarter; B, Year: =VALUE(LEFT(A5,4)) takes the first 4 characters of the quarter (the year) as a number.",
    "C, Events (n): =COUNTIFS(ev_quarter, A5). D, Class I (d): =SUMIFS(ev_class1, ev_quarter, A5).",
    "E, Share (p): =D/C, the share of that quarter's events that were Class I.",
    "F, UCL (upper control limit): =p-bar + 3*SQRT(p-bar*(1-p-bar)/n). G, LCL: the same with minus, never below 0. Quarters with fewer events get wider limits.",
    "H, Above UCL?: 'Yes' when the share is above the upper limit, a signal that something has changed.",
    "# Cells L5-L8",
    "L5, p-bar (center line): Class I events / all events in 2013-2018 (Phase I, the 'normal' period) = 4.34%.",
    "L6 and L7: count of quarters above the UCL, in total and in 2023-2025. L8: number of quarters in 2023-2025.",
    "# What it shows",
    "No quarter from 2013 to 2018 is outside the limits, but 11 of the 12 quarters in 2023-2025 are above the UCL: a real shift, not chance.",
], width=95)

explain("7 Radiology", "H", 4, [
    "Question: is the picture different for imaging devices (CT, MRI, X-ray), the products of companies like Siemens Healthineers?",
    "# Columns",
    "B, Recall events: =COUNTIFS(ev_cat, A5, ev_spec, \"Radiology\"), events in that area AND in the Radiology specialty.",
    "C, Share and D, Cumulative share: the same Pareto logic as sheet 2.",
    "E, Class I events: SUMIFS with the same two conditions. F, Class I rate: E / B.",
    "# What it shows",
    "Software causes 46% of imaging recalls and hardware design 25%. Only 18 imaging recalls were Class I, so imaging recalls are mainly a cost and downtime problem, with production failures behind most of the serious ones.",
], width=90)

explain("8 Impact", "E", 4, [
    "Question: how much would focusing on the top two areas help?",
    "B4: Class I events per year (2023-2025) in the areas ranked 1 and 2. =SUMIFS(column K of sheet 5, rank column, \"<=2\") adds the rows whose rank is 2 or better.",
    "B5: Class I events per year in all known-cause areas (sum of sheet 4, column H).",
    "B6: B4 / B5, the share of serious recalls the top two areas cause (66%).",
    "Rows 9-11: B = B4 x the reduction in column A. If prevention cut recalls in the top two areas by 25%, about 13 Class I recalls a year would be avoided. C = that number as a share of all known-cause Class I recalls.",
    "These are scenarios showing the size of the opportunity, not predictions. Change the blue percentages in column A to try other values.",
], width=110)

explain("9 Logistic regression", "U", 4, [
    "Question: is a root cause still linked to Class I after we account for what kind of device it is (device class, specialty) and when it happened (period)?",
    "Why it matters: design recalls might look dangerous only because design problems are common in high-risk devices. The regression compares like with like.",
    "# The model",
    "logit = LN(p / (1 - p)) = b0 + b(root cause) + b(device class) + b(specialty) + b(period). p is the probability an event is Class I.",
    "Each variable has a reference level with b = 0 (Production process, Class 2, General Hospital, 2013-2018). Every other level's b says how it differs from that reference.",
    "# Table columns",
    "D, Coefficient b and E, Std error: estimated by maximum likelihood in Python, because Excel has no built-in logistic regression. They are typed in (blue).",
    "F, z = b / SE and G, p-value = 2*(1-NORMSDIST(ABS(z))): the same z-test idea as sheet 1. p below 0.05 = a real difference from the reference.",
    "H, Odds ratio = EXP(b). 2.01 for design: design recalls have about twice the odds of being Class I compared with production recalls on the same kind of device in the same period.",
    "I and J, 95% CI = EXP(b -/+ 1.96 x SE). If the interval does not include 1, the difference is significant.",
    "K and L: events and Class I events at that level (COUNTIFS / SUMIFS). M: the sum of predicted probabilities (column O of Events). For every level, M equals L: a property of a correctly fitted model.",
    "N: p-value with standard errors clustered by firm (Python). This allows several recalls from the same company to be related; the main findings stay significant.",
    "# Model check block",
    "SUM(ev_ll) recomputes the log-likelihood from the 10,640 events using only Excel formulas. Matching Python's value proves the coefficients are applied correctly.",
    "LR chi-square = 2 x (LL model - LL null) tests whether the predictors together help (p is essentially 0). McFadden R-squared of 0.2-0.4 is considered a good fit for this kind of model; here 0.22.",
    "AUC 0.84: given one Class I and one non-Class I event, the model gives the Class I event the higher probability 84% of the time (0.5 = guessing, 1 = perfect).",
    "# Unadjusted vs adjusted",
    "Unadjusted OR uses only the raw Class I rates (like sheet 3). Adjusted OR comes from the model. Design rises from 1.67 to 2.01: its higher severity is NOT explained by device type, it gets stronger. Software moves closer to 1 (0.59 to 0.78) and is no longer significant: its lower severity was partly due to the kinds of devices it affects.",
    "# What-if calculator",
    "Pick a root cause, device class, specialty and period from the drop-downs. The sheet adds the matching coefficients and converts the total to a probability. Example: a 2023-2025 Class 2 General Hospital recall: 23% if production, 38% if design.",
    "# Try it yourself (optional)",
    "You can re-estimate the coefficients in Excel: Data > Solver, maximize the 'Log-likelihood of the model' cell by changing the blue coefficients in column D (leave reference rows at 0). Solver should arrive at the same values.",
], width=110)

wb.save("outputs/Recall_Analysis_Calculations.xlsx")
print("Saved outputs/Recall_Analysis_Calculations.xlsx")
