"""
STEP 4 - Build the calculations workbook in the same formats used in ISE 250 class material.

Input : outputs/study_events.csv   (made by step2_analysis.py)
Output: outputs/Recall_Analysis_ISE250_Style.xlsx
Run   : python step4_excel_ise250.py   (about 30 seconds)

Sheets follow the DMAIC order and copy the layouts from class:
  - Test of two proportions  -> textbook Sec. 4.6.2 (KIND Karz airbag vendors, Table 4.13) + Minitab 2 Proportions output
  - p chart                  -> Module 3 "P chart_Excel Workbook" and Final Exam Q4 answer layout
  - Pareto                   -> textbook Sec. 4.5 / Minitab Pareto chart (Count, Percent, Cum %)
  - Chi-square               -> textbook Sec. 4.6.3 (KIND Karz brakes, Table 4.15), critical value with CHISQ.INV.RT
  - FMEA                     -> textbook Table 6.7 (ID, Action, Failure Mode, Failure Effects, SEV, OCC, DET, RPN)
  - Regression               -> Minitab regression handout style (Coef, SE Coef, Z/T-Value, P-Value, Regression Equation)
"""
import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from scipy import stats

ev = pd.read_csv("outputs/study_events.csv", dtype={"event_id": str})
ev["quarter"] = pd.PeriodIndex(pd.to_datetime(ev["date"], errors="coerce"), freq="Q").astype(str)

# ---------------------------------------------------------------- styles
F = "Calibri"
BLUE = Font(name=F, size=11, color="0000FF")
BLACK = Font(name=F, size=11)
BOLD = Font(name=F, size=11, bold=True)
WHITE = Font(name=F, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=F, size=15, bold=True, color="1F3864")
SUB = Font(name=F, size=10, italic=True, color="555555")
MONO = Font(name="Consolas", size=10)
MONOB = Font(name="Consolas", size=10, bold=True)
HEAD = PatternFill("solid", fgColor="1F3864")
KEY = PatternFill("solid", fgColor="FFFF00")
REF = PatternFill("solid", fgColor="E2EFDA")      # green box = "where this is in class"
HYP = PatternFill("solid", fgColor="FFF2CC")      # yellow box = hypotheses / decision
OUT = PatternFill("solid", fgColor="F2F2F2")      # grey box = Minitab-style output
side = Side(style="thin", color="A6A6A6")
BOX = Border(top=side, bottom=side, left=side, right=side)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFTWRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()


def put(ws, ref, value, font=BLACK, fmt=None, fill=None, box=True, align=None):
    c = ws[ref]
    c.value, c.font = value, font
    if box:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


def header(ws, row, labels, col=1, height=32):
    for j, t in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=t)
        c.font, c.fill, c.alignment, c.border = WHITE, HEAD, WRAP, BOX
    ws.row_dimensions[row].height = height


def title(ws, t, s):
    ws["A1"], ws["A2"] = t, s
    ws["A1"].font, ws["A2"].font = TITLE, SUB


def box(ws, row, col, ncols, lines, fill, head=None):
    """A shaded text box: one merged row per line. Returns next free row."""
    r = row
    if head:
        c = ws.cell(row=r, column=col, value=head)
        c.font, c.fill = BOLD, fill
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + ncols - 1)
        r += 1
    for line in lines:
        c = ws.cell(row=r, column=col, value=line)
        c.font, c.fill, c.alignment = BLACK, fill, LEFTWRAP
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + ncols - 1)
        ws.row_dimensions[r].height = 15 * max(1, len(str(line or "")) // 105 + 1)
        r += 1
    return r


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# ================================================================ Read Me (DMAIC map)
ws = wb.active
ws.title = "Read Me"
title(ws, "Preventing Serious Medical Device Recalls: ISE 250-style calculations",
      "13,410 FDA recall events, 2013-2025. Organised by DMAIC, with each sheet in the same format as the class material.")
header(ws, 4, ["DMAIC phase", "Question (in the words of the Six Sigma slides)", "Tool used", "Sheet", "Where it is in your ISE 250 material"])
rows = [
    ("Define", "Identify the problem and project goals: are serious (Class I) recalls rising?", "Trend table and chart", "D1 Problem",
     "Six Sigma.pptx: DMAIC - Define"),
    ("Measure", "Measure the defect rate: is the Class I share really higher now than before?", "Test of two proportions", "M1 Two Proportions",
     "Textbook 4.6.2 Test of Two Proportions (KIND Karz airbags, Table 4.13); Minitab Stat > Basic Statistics > 2 Proportions"),
    ("Measure", "Is the process in statistical control?", "p chart (attribute control chart)", "M2 p Chart",
     "Module 3 P chart_Excel Workbook; Final Exam Q4; HW2 Q1; Statistical Process Control.pdf"),
    ("Analyze", "Which root causes create most of the recalls?", "Pareto chart", "A1 Pareto",
     "Textbook 4.5 Pareto Analysis (KIND Karz defect categories, Fig. 4.12)"),
    ("Analyze", "Is one root cause worse (more Class I) than the others?", "Chi-square test of multiple proportions", "A2 Chi-Square",
     "Textbook 4.6.3 Chi-Square Test of Multiple Proportions (KIND Karz brakes, Table 4.15)"),
    ("Analyze", "How sure are we about each Class I rate?", "95% confidence interval for a proportion", "A3 Confidence Intervals",
     "Minitab handout: confidence level = 1 - alpha; textbook: smaller samples give wider intervals"),
    ("Analyze", "Is the root-cause effect still there after accounting for device type and time?", "Binary logistic regression", "A4 Logistic Regression",
     "Module 4 Correlation and Regression + Minitab regression handout (Coef, SE Coef, P-Value, Regression Equation). Logistic is the yes/no version"),
    ("Analyze", "Is the picture different for imaging devices?", "Pareto for one specialty", "A5 Radiology", "Same as A1"),
    ("Improve", "Where should improvement go first?", "FMEA, Risk Priority Number", "I1 FMEA (RPN)",
     "Textbook 6.5 Risk Analysis (KIND Karz airbag FMEA, Table 6.7); FMEA for Wind Turbine.pdf"),
    ("Improve", "What would it be worth?", "Impact scenario", "I2 Impact", "Textbook 6.4: pro forma estimate of the improvement"),
    ("Control", "How do we keep the gains?", "Control plan using the p chart", "C1 Control Plan",
     "Six Sigma.pptx: DMAIC - Control ('put tools in place to ensure key variables remain within acceptable ranges')"),
]
for i, r in enumerate(rows, start=5):
    for j, v in enumerate(r):
        put(ws, f"{L(j + 1)}{i}", v, BOLD if j in (0, 3) else BLACK, align=LEFTWRAP)
    ws.row_dimensions[i].height = 32
r = box(ws, 18, 1, 5, [
    "Blue text = data or a value you can change.  Black = Excel formula (click a cell to see it).  Yellow cell = the headline answer.",
    "Green box = where this method is in your class material.  Yellow box = hypotheses and decision.  Grey box = Minitab-style output.",
    "Decision rule used on every test (as in the Minitab handout): if P-value < alpha (0.05), reject H0.",
    "Data sheet: one row per recall event. class1 = 1 if the FDA rated the recall Class I (most dangerous), else 0 - this is the 'defective' in every test.",
    "Settings sheet: study periods and alpha. Change a blue value and every sheet updates.",
], REF, head="How to read this workbook")
widths(ws, {"A": 12, "B": 60, "C": 32, "D": 24, "E": 70})

# ================================================================ Settings
ws = wb.create_sheet("Settings")
title(ws, "Settings", "Every sheet reads these cells by name.")
header(ws, 4, ["Setting", "Value", "Name used in formulas"])
for i, (a, v, n) in enumerate([
    ("Baseline period: first year", 2016, "Base_Start"), ("Baseline period: last year", 2018, "Base_End"),
    ("Recent period: first year", 2023, "Recent_Start"), ("Recent period: last year", 2025, "Recent_End"),
    ("p chart Phase I (limits set from 2013 to this year)", 2018, "Phase1_End"),
    ("Significance level alpha", 0.05, "Alpha"), ("z for 95% confidence", 1.96, "Z"),
], start=5):
    put(ws, f"A{i}", a)
    put(ws, f"B{i}", v, BLUE)
    put(ws, f"C{i}", n, BOLD)
    wb.defined_names[n] = DefinedName(n, attr_text=f"Settings!$B${i}")
widths(ws, {"A": 48, "B": 10, "C": 22})

# ================================================================ Data
ws = wb.create_sheet("Data")
title(ws, "Data: 13,410 recall events, 2013-2025", "From step2_analysis.py. One row = one recall event. class1 = 1 if Class I.")
cols = ["event_id", "year", "quarter", "date", "firm", "root_cause", "category", "specialty", "products", "class1",
        "devclass", "spec_model", "period"]
header(ws, 4, cols + ["Y' (logit)", "P(Class I)", "log-likelihood part"])
E0, E1 = 5, 5 + len(ev) - 1
LKP = "VLOOKUP(\"{v}|\"&{c}{r},lr_tab,3,0)"
for i, rec in enumerate(ev[cols].itertuples(index=False), start=E0):
    for j, v in enumerate(rec, start=1):
        ws.cell(row=i, column=j, value=v).font = BLUE
    terms = "+".join(LKP.format(v=v, c=c, r=i) for v, c in
                     (("Root cause", "G"), ("Device class", "K"), ("Specialty", "L"), ("Period", "M")))
    ws.cell(row=i, column=14, value=f'=IF(G{i}="Not determined","",lr_b0+{terms})').number_format = "0.000"
    ws.cell(row=i, column=15, value=f'=IF(N{i}="","",EXP(N{i})/(1+EXP(N{i})))').number_format = "0.000"
    ws.cell(row=i, column=16, value=f'=IF(N{i}="","",J{i}*LN(O{i})+(1-J{i})*LN(1-O{i}))').number_format = "0.000"
ws.freeze_panes = "A5"
widths(ws, {"A": 9, "B": 7, "C": 9, "D": 11, "E": 28, "F": 30, "G": 20, "H": 22, "I": 9, "J": 8,
            "K": 10, "L": 22, "M": 11, "N": 11, "O": 11, "P": 12})
for nm, col in (("ev_year", "B"), ("ev_quarter", "C"), ("ev_cat", "G"), ("ev_spec", "H"), ("ev_class1", "J"),
                ("ev_devclass", "K"), ("ev_specm", "L"), ("ev_period", "M"), ("ev_phat", "O"), ("ev_ll", "P")):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"Data!${col}${E0}:${col}${E1}")

CATS = ["Production process", "Design (hardware)", "Software", "Supplier & material",
        "Labeling & packaging", "Regulatory", "Use error"]
KNOWN = 'ev_cat,"<>Not determined"'

# ================================================================ D1 Problem
ws = wb.create_sheet("D1 Problem")
title(ws, "D1. Define: serious recalls are rising",
      "Recall events and Class I events per year. Class I = the FDA's most serious rating (risk of serious injury or death).")
header(ws, 4, ["Year", "Recall events", "Class I events", "Class I share"])
for k, yr in enumerate(range(2013, 2026), start=5):
    put(ws, f"A{k}", yr, BLUE)
    put(ws, f"B{k}", f"=COUNTIFS(ev_year,A{k})", fmt="#,##0")
    put(ws, f"C{k}", f"=SUMIFS(ev_class1,ev_year,A{k})", fmt="#,##0")
    put(ws, f"D{k}", f"=C{k}/B{k}", fmt="0.0%")
put(ws, "A19", "Change 2020 to 2025, Class I events", BOLD)
put(ws, "D19", "=C17/C12-1", fmt="0%", fill=KEY)
put(ws, "A20", "Change 2020 to 2025, all recall events", BOLD)
put(ws, "D20", "=B17/B12-1", fmt="0%")
box(ws, 22, 1, 4, [
    "Problem statement: Class I recalls grew from 37 (2020) to 125 (2025) while total recalls barely changed.",
    "The FDA itself reported +232%; our event count gives +238%, so the data agrees with the FDA.",
    "Project goal: find which quality-system areas cause the serious recalls, and where to act first.",
], HYP, head="Define - problem statement and goal")
ch = BarChart()
ch.title, ch.y_axis.title = "Class I recall events per year", "Class I events"
ch.add_data(Reference(ws, min_col=3, min_row=4, max_row=17), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=17))
ch.height, ch.width, ch.legend = 8, 16, None
ws.add_chart(ch, "F4")
widths(ws, {"A": 36, "B": 14, "C": 14, "D": 14})

# ================================================================ M1 Two Proportions
ws = wb.create_sheet("M1 Two Proportions")
title(ws, "M1. Test of two proportions: is the Class I share higher now?",
      "Same layout as the textbook's KIND Karz airbag example (Table 4.13): number of events and number of trials for two samples.")
r = box(ws, 4, 1, 6, [
    "Textbook Sec. 4.6.2 'Test of Two Proportions in Minitab': airbags with issues 200 of 50,000 (vendor 1) vs 180 of 40,000 (vendor 2).",
    "Here: Class I recalls out of all recalls, recent period (sample 1) vs baseline period (sample 2).",
    "Minitab path: Stat > Basic Statistics > 2 Proportions > Summarized data. Number of events = Class I recalls; number of trials = all recalls.",
], REF, head="Where this is in class")
header(ws, 9, ["Category", "Sample 1: Recent", "Sample 2: Baseline"])
put(ws, "A10", "Years", BOLD)
put(ws, "B10", '=Recent_Start&"-"&Recent_End')
put(ws, "C10", '=Base_Start&"-"&Base_End')
put(ws, "A11", "Class I recalls (number of events, X)", BOLD)
put(ws, "B11", '=SUMIFS(ev_class1,ev_year,">="&Recent_Start,ev_year,"<="&Recent_End)', fmt="#,##0")
put(ws, "C11", '=SUMIFS(ev_class1,ev_year,">="&Base_Start,ev_year,"<="&Base_End)', fmt="#,##0")
put(ws, "A12", "All recalls (number of trials, N)", BOLD)
put(ws, "B12", '=COUNTIFS(ev_year,">="&Recent_Start,ev_year,"<="&Recent_End)', fmt="#,##0")
put(ws, "C12", '=COUNTIFS(ev_year,">="&Base_Start,ev_year,"<="&Base_End)', fmt="#,##0")
put(ws, "A13", "Sample p = X / N (percent Class I)", BOLD)
put(ws, "B13", "=B11/B12", fmt="0.00%", fill=KEY)
put(ws, "C13", "=C11/C12", fmt="0.00%")

box(ws, 15, 1, 6, [
    "H0: p1 = p2  (the Class I share is the same in both periods)",
    "H1: p1 ≠ p2  (at least one proportion is different)",
    "Significance level: alpha = 0.05.  Decision rule: if P-value < 0.05, reject H0.",
], HYP, head="Hypotheses")

header(ws, 20, ["Step-by-step calculation", "Value", "Formula in words"])
steps = [
    ("Difference p1 - p2", "=B13-C13", "0.0000", "recent share minus baseline share"),
    ("Pooled proportion p-hat", "=(B11+C11)/(B12+C12)", "0.0000", "(X1 + X2) / (N1 + N2): the share if nothing had changed"),
    ("Standard error (pooled)", "=SQRT(B22*(1-B22)*(1/B12+1/C12))", "0.00000", "SQRT( p-hat x (1 - p-hat) x (1/N1 + 1/N2) )"),
    ("Z-Value", "=B21/B23", "0.00", "difference / standard error"),
    ("P-Value (two-sided)", "=2*(1-NORMSDIST(ABS(B24)))", "0.0E+00", "2 x (1 - NORMSDIST(|Z|))"),
    ("Standard error (unpooled, for the CI)", "=SQRT(B13*(1-B13)/B12+C13*(1-C13)/C12)", "0.00000", "SQRT( p1(1-p1)/N1 + p2(1-p2)/N2 )"),
    ("95% CI for difference: lower", "=B21-Z*B26", "0.0000", "difference - 1.96 x SE"),
    ("95% CI for difference: upper", "=B21+Z*B26", "0.0000", "difference + 1.96 x SE"),
]
for i, (a, f, fmt, w) in enumerate(steps, start=21):
    put(ws, f"A{i}", a, BOLD)
    put(ws, f"B{i}", f, fmt=fmt, fill=KEY if i in (24, 25) else None)
    put(ws, f"C{i}", w, SUB)

# Minitab-style output
o = 31
put(ws, f"A{o}", "Test and CI for Two Proportions  (Minitab-style output)", MONOB, fill=OUT, box=False)
put(ws, f"A{o+1}", "Method:  p1 = proportion Class I, Recent;  p2 = proportion Class I, Baseline;  Difference: p1 - p2", MONO, fill=OUT, box=False)
put(ws, f"A{o+2}", "Descriptive Statistics", MONOB, fill=OUT, box=False)
for j, h in enumerate(["Sample", "N", "Event", "Sample p"]):
    put(ws, f"{L(j+1)}{o+3}", h, MONOB, fill=OUT)
for k, (lab, col) in enumerate((("Sample 1 (Recent)", "B"), ("Sample 2 (Baseline)", "C")), start=o + 4):
    put(ws, f"A{k}", lab, MONO, fill=OUT)
    put(ws, f"B{k}", f"={col}12", MONO, "#,##0", OUT)
    put(ws, f"C{k}", f"={col}11", MONO, "#,##0", OUT)
    put(ws, f"D{k}", f"={col}13", MONO, "0.000000", OUT)
put(ws, f"A{o+7}", "Estimation for Difference", MONOB, fill=OUT, box=False)
put(ws, f"A{o+8}", "Difference", MONOB, fill=OUT)
put(ws, f"B{o+8}", "95% CI for Difference", MONOB, fill=OUT)
put(ws, f"A{o+9}", "=B21", MONO, "0.000000", OUT)
put(ws, f"B{o+9}", '="("&TEXT(B27,"0.000000")&", "&TEXT(B28,"0.000000")&")"', MONO, fill=OUT)
put(ws, f"A{o+10}", "Test", MONOB, fill=OUT, box=False)
put(ws, f"A{o+11}", "Null hypothesis  H0: p1 - p2 = 0;  Alternative  H1: p1 - p2 ≠ 0", MONO, fill=OUT, box=False)
for j, h in enumerate(["Method", "Z-Value", "P-Value"]):
    put(ws, f"{L(j+1)}{o+12}", h, MONOB, fill=OUT)
put(ws, f"A{o+13}", "Normal approximation (pooled)", MONO, fill=OUT)
put(ws, f"B{o+13}", "=B24", MONO, "0.00", OUT)
put(ws, f"C{o+13}", '=IF(B25<0.0005,"0.000",TEXT(B25,"0.000"))', MONO, fill=OUT)

d = o + 15
box(ws, d, 1, 6, [None], HYP, head="Decision and conclusion")
ws[f"A{d+1}"] = ('=IF(B25<Alpha,"P-value ("&TEXT(B25,"0.0E+00")&") < 0.05, so we reject H0. The Class I share is significantly higher in "&B10&" ("&TEXT(B13,"0.0%")&") than in "&C10&" ("&TEXT(C13,"0.0%")&"). The rise in serious recalls is real, not chance.",'
                 '"P-value >= 0.05, so we do not reject H0: we cannot conclude the two periods differ.")')
ws.row_dimensions[d + 1].height = 32
box(ws, d + 3, 1, 6, [
    "Low cell counts: use Fisher's exact test if the number of events is under 5. Here the smallest is 120, so the normal approximation is fine.",
    "Independence: several recalls can come from one company. The largest firm has only 2.2% of events, so this barely matters (sheet A4 checks it).",
], REF, head="Watchouts (textbook Step 3)")
widths(ws, {"A": 44, "B": 22, "C": 52, "D": 14, "E": 12, "F": 12})

# ================================================================ M2 p Chart
ws = wb.create_sheet("M2 p Chart")
title(ws, "M2. p chart: share of recalls that are Class I, by quarter",
      "Same layout as the Module 3 'P chart_Excel Workbook' and your Final Exam Q4 answer.")
box(ws, 4, 1, 9, [
    "Class: subgroups (samples) = shifts or hours; # of defective parts; sample size n; P = defectives / n; P bar; UCL; LCL ('LCL cannot be negative').",
    "Here: subgroup = one quarter; 'defective' = a Class I recall; n = all recalls that quarter. 52 quarters, 2013 Q1 to 2025 Q4.",
    "One difference from class: in class n was the same every sample (50 or 100), so there was one UCL. Here n changes every quarter (178 to 331), so UCL and LCL are calculated on each row with the same formula.",
    "Phase I: P bar is set from 2013-2018 (the 'normal' period). Later quarters are then checked against those limits (Phase II).",
], REF, head="Where this is in class")
P0 = 10
header(ws, P0, ["Subgroup (Quarter)", "Year", "# of Class I recalls (defectives)", "Sample Size (n)",
                "P (fraction Class I)", "P bar", "UCL", "LCL", "Above UCL?", "Run on same side of P bar"])
quarters = sorted(ev["quarter"].unique())
Q0, Q1 = P0 + 1, P0 + len(quarters)
for k, qq in enumerate(quarters, start=Q0):
    put(ws, f"A{k}", qq, BLUE)
    put(ws, f"B{k}", f"=VALUE(LEFT(A{k},4))")
    put(ws, f"C{k}", f"=SUMIFS(ev_class1,ev_quarter,A{k})")
    put(ws, f"D{k}", f"=COUNTIFS(ev_quarter,A{k})")
    put(ws, f"E{k}", f"=C{k}/D{k}", fmt="0.0000")
    put(ws, f"F{k}", "=$N$13", fmt="0.0000")
    put(ws, f"G{k}", f"=F{k}+3*SQRT(F{k}*(1-F{k})/D{k})", fmt="0.0000")
    put(ws, f"H{k}", f"=MAX(0,F{k}-3*SQRT(F{k}*(1-F{k})/D{k}))", fmt="0.0000")
    put(ws, f"I{k}", f'=IF(E{k}>G{k},"Yes",IF(E{k}<H{k},"Below LCL",""))')
    prev = f"J{k-1}" if k > Q0 else "0"
    put(ws, f"J{k}", f"=IF(E{k}>F{k},IF({prev}>0,{prev}+1,1),IF(E{k}<F{k},IF({prev}<0,{prev}-1,-1),0))")
# summary block (class layout: n=, k=, P bar, UCL, LCL)
put(ws, "L11", "k = number of subgroups", BOLD)
put(ws, "N11", f"=COUNTA(A{Q0}:A{Q1})")
put(ws, "L12", "Phase I subgroups (2013 to Phase1_End)", BOLD)
put(ws, "N12", f'=COUNTIFS(B{Q0}:B{Q1},"<="&Phase1_End)')
put(ws, "L13", "P bar = total defectives / total inspected (Phase I)", BOLD)
put(ws, "N13", f'=SUMIFS(C{Q0}:C{Q1},B{Q0}:B{Q1},"<="&Phase1_End)/SUMIFS(D{Q0}:D{Q1},B{Q0}:B{Q1},"<="&Phase1_End)', fmt="0.0000", fill=KEY)
put(ws, "L14", "Average n (for reference)", BOLD)
put(ws, "N14", f"=AVERAGE(D{Q0}:D{Q1})", fmt="0.0")
put(ws, "L15", "UCL at the average n = P bar + 3*SQRT(P bar(1-P bar)/n)", BOLD)
put(ws, "N15", "=N13+3*SQRT(N13*(1-N13)/N14)", fmt="0.0000")
put(ws, "L16", "LCL at the average n (LCL cannot be negative)", BOLD)
put(ws, "N16", "=MAX(0,N13-3*SQRT(N13*(1-N13)/N14))", fmt="0.0000")

header(ws, 18, ["A process is in control if ... (class slide)", "", "Result"], col=12, height=18)
ws.merge_cells("L18:M18")
checks = [
    ("1. ... no sample points outside limits", f'=COUNTIF(I{Q0}:I{Q1},"Yes")+COUNTIF(I{Q0}:I{Q1},"Below LCL")&" points outside the limits"'),
    ("   of which in Phase I (2013-2018)", f'=COUNTIFS(I{Q0}:I{Q1},"Yes",B{Q0}:B{Q1},"<="&Phase1_End)&" points"'),
    ("   of which in 2023-2025", f'=COUNTIFS(I{Q0}:I{Q1},"Yes",B{Q0}:B{Q1},">=2023")&" of "&COUNTIFS(B{Q0}:B{Q1},">=2023")&" quarters above UCL"'),
    ("2. ... most points near process average", "See the chart: recent points sit well above P bar"),
    ("3. ... about equal number above and below the centerline", f'=COUNTIF(J{Q0}:J{Q1},">0")&" above vs "&COUNTIF(J{Q0}:J{Q1},"<0")&" below"'),
    ("4. ... points appear randomly distributed (8 in a row on one side = shift)", f'="Longest run above P bar: "&MAX(J{Q0}:J{Q1})&" quarters in a row"'),
]
for i, (a, f) in enumerate(checks, start=19):
    put(ws, f"L{i}", a, BOLD)
    ws.merge_cells(f"L{i}:M{i}")
    put(ws, f"N{i}", f)
box(ws, 4, 12, 3, [None], HYP, head="Conclusion (same wording style as Final Exam Q4)")
ws["L5"] = (f'=IF(COUNTIF(I{Q0}:I{Q1},"Yes")=0,"There are no points above UCL and below LCL, therefore we can say the process is in control.",'
             f'"There are "&COUNTIF(I{Q0}:I{Q1},"Yes")&" points above the UCL ("&COUNTIFS(I{Q0}:I{Q1},"Yes",B{Q0}:B{Q1},">=2023")&" of them in 2023-2025) and a run of "&MAX(J{Q0}:J{Q1})&" quarters above P bar, therefore the process is NOT in control: there is special-cause variation and the Class I share has shifted upward.")')
ws.row_dimensions[5].height = 62
ch = LineChart()
ch.title, ch.y_axis.title, ch.x_axis.title = "P chart: fraction of recalls that are Class I", "P", "Quarter"
ch.height, ch.width = 9, 26
for col in (5, 6, 7, 8):
    ch.add_data(Reference(ws, min_col=col, min_row=P0, max_row=Q1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=Q0, max_row=Q1))
colors = ["2A6FDB", "000000", "C00000", "C00000"]
for s_, c_ in zip(ch.series, colors):
    s_.graphicalProperties.line.solidFill = c_
    s_.smooth = False
ch.series[2].graphicalProperties.line.dashStyle = "dash"
ch.series[3].graphicalProperties.line.dashStyle = "dash"
ws.add_chart(ch, "L27")
widths(ws, {"A": 12, "B": 7, "C": 14, "D": 11, "E": 11, "F": 9, "G": 9, "H": 9, "I": 11, "J": 12, "K": 3,
            "L": 52, "M": 6, "N": 46})

# ================================================================ A1 Pareto
ws = wb.create_sheet("A1 Pareto")
title(ws, "A1. Pareto chart of recall root causes",
      "Events with a known root cause, 2013-2025. Bars from highest to lowest count; line = cumulative %.")
box(ws, 4, 1, 8, [
    "Textbook Sec. 4.1.5 / 4.5: 'A Pareto chart is a bar chart where the X axis represents categories, arranged from highest to smallest counts. Look for ~20% of the categories representing ~80% of the problem.'",
    "KIND Karz example: brakes and airbags together were 79.1% of defects. Minitab path: Stat > Quality Tools > Pareto Chart.",
], REF, head="Where this is in class")
put(ws, "A8", "Minitab-style Pareto table", MONOB, fill=OUT, box=False)
put(ws, "A9", "Root cause area", MONOB, fill=OUT)
for j, c in enumerate(CATS):
    col = L(j + 2)
    put(ws, f"{col}9", c, MONOB, fill=OUT, align=WRAP)
    put(ws, f"{col}10", f'=COUNTIFS(ev_cat,{col}9)', MONO, "#,##0", OUT)
    put(ws, f"{col}11", f"={col}10/$I$10", MONO, "0.0%", OUT)
    put(ws, f"{col}12", f"=SUM($B$11:{col}11)", MONO, "0.0%", OUT)
put(ws, "A10", "Count", MONOB, fill=OUT)
put(ws, "A11", "Percent", MONOB, fill=OUT)
put(ws, "A12", "Cum %", MONOB, fill=OUT)
put(ws, "I9", "Total", MONOB, fill=OUT)
put(ws, "I10", "=SUM(B10:H10)", MONO, "#,##0", OUT)
ws.row_dimensions[9].height = 32
put(ws, "A14", "Not determined (no root cause found), kept out of the Pareto", BOLD)
put(ws, "E14", '=COUNTIFS(ev_cat,"Not determined")', fmt="#,##0")
put(ws, "F14", "=E14/(E14+I10)", fmt="0.0%", fill=KEY)
put(ws, "G14", "of all events", SUB, box=False)
box(ws, 16, 1, 8, [
    "Top 2 areas (production process, hardware design) = 47.4%; top 4 = 79.8% of recalls with a known cause.",
    "So recalls here follow roughly 80/50 rather than 80/20: the problem is spread over several areas, and no single fix solves it.",
    "A Pareto counts every recall the same, safe or dangerous. Sheet A2 asks which areas produce the dangerous (Class I) ones.",
    "Not determined: 20.7% of all events had no root cause found, and 28% recently. That is a gap in corrective action (ISO 13485 clause 8.5.2).",
], HYP, head="Interpretation")
ch = BarChart()
ch.title, ch.y_axis.title = "Pareto chart of root causes", "Count"
ch.add_data(Reference(ws, min_col=1, max_col=8, min_row=10), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, max_col=8, min_row=9))
ch.series[0].graphicalProperties.solidFill = "4472C4"
ln = LineChart()
ln.add_data(Reference(ws, min_col=1, max_col=8, min_row=12), from_rows=True, titles_from_data=True)
ln.y_axis.axId, ln.y_axis.title, ln.y_axis.number_format = 200, "Cum %", "0%"
ln.y_axis.scaling.min, ln.y_axis.scaling.max = 0, 1
ln.y_axis.crosses = "max"
ln.series[0].graphicalProperties.line.solidFill = "C00000"
ch += ln
ch.height, ch.width = 9, 22
ws.add_chart(ch, "A22")
widths(ws, {"A": 30, **{L(c): 15 for c in range(2, 10)}})

# ================================================================ A2 Chi-Square
ws = wb.create_sheet("A2 Chi-Square")
title(ws, "A2. Chi-square test of multiple proportions: is one root cause more dangerous?",
      "Same layout as the textbook's KIND Karz brake example (Table 4.15): categories across, issues / no issues down.")
box(ws, 4, 1, 10, [
    "Textbook Sec. 4.6.3: brake warranty issues by vehicle 1-5; chi-square = sum of (Observed - Expected)^2 / Expected; critical value = CHISQ.INV.RT(0.05, DF), DF = number of proportions - 1.",
    "Here: Class I recalls by root-cause area (7 areas). 'Defect' = a Class I recall. Minitab path: Stat > Tables > Chi-Square Test (Two-Way Table in Worksheet).",
    "In the textbook table only the 'issues' row contributions were shown because the 'no issues' row added almost nothing. Here both rows are added, as Minitab does.",
], REF, head="Where this is in class")
put(ws, "A9", "Category", WHITE, fill=HEAD, align=WRAP)
for j, c in enumerate(CATS):
    put(ws, f"{L(j+2)}9", c, WHITE, fill=HEAD, align=WRAP)
put(ws, "I9", "Total", WHITE, fill=HEAD, align=WRAP)
put(ws, "J9", "Expected proportion", WHITE, fill=HEAD, align=WRAP)
ws.row_dimensions[9].height = 32
lab = {10: "Class I recalls (observed)", 11: "Not Class I recalls (observed)", 12: "Total recalls",
       13: "Percent Class I", 14: "Expected Class I", 15: "Chi-Sq value, Class I",
       16: "Expected not Class I", 17: "Chi-Sq value, not Class I"}
for rr, t in lab.items():
    put(ws, f"A{rr}", t, BOLD)
for j in range(len(CATS)):
    c = L(j + 2)
    put(ws, f"{c}10", f"=SUMIFS(ev_class1,ev_cat,{c}9)", fmt="#,##0")
    put(ws, f"{c}11", f"=COUNTIFS(ev_cat,{c}9)-{c}10", fmt="#,##0")
    put(ws, f"{c}12", f"={c}10+{c}11", fmt="#,##0")
    put(ws, f"{c}13", f"={c}10/{c}12", fmt="0.00%")
    put(ws, f"{c}14", f"={c}12*$J$10", fmt="0.0")
    put(ws, f"{c}15", f"=({c}10-{c}14)^2/{c}14", fmt="0.0")
    put(ws, f"{c}16", f"={c}12*$J$11", fmt="0.0")
    put(ws, f"{c}17", f"=({c}11-{c}16)^2/{c}16", fmt="0.0")
for rr in (10, 11, 12, 15, 17):
    put(ws, f"I{rr}", f"=SUM(B{rr}:H{rr})", BOLD, "#,##0.0" if rr in (15, 17) else "#,##0")
put(ws, "J10", "=I10/I12", fmt="0.00%")
put(ws, "J11", "=I11/I12", fmt="0.00%")
put(ws, "I13", "=I10/I12", fmt="0.00%")

header(ws, 19, ["Test result", "Value", "How it is calculated"])
res = [
    ("Total Chi-Sq", "=I15+I17", "0.0", "sum of every (O - E)^2 / E"),
    ("DF = number of proportions - 1", f"=COUNTA(B9:H9)-1", "0", "7 areas - 1 (same as (rows-1) x (columns-1))"),
    ("Critical Chi-Sq statistic (alpha = 0.05)", "=_xlfn.CHISQ.INV.RT(Alpha,B21)", "0.00", "Excel: CHISQ.INV.RT(0.05, DF), as in the textbook"),
    ("P-Value", "=_xlfn.CHISQ.DIST.RT(B20,B21)", "0.0E+00", "Excel: CHISQ.DIST.RT(Chi-Sq, DF)"),
    ("Largest contribution to Chi-Sq", '=INDEX(B9:H9,MATCH(MAX(B15:H15),B15:H15,0))', None, "the most different area (like Vehicle 4 in the textbook)"),
    ("   its contribution", "=MAX(B15:H15)", "0.0", ""),
    ("Smallest expected count (must be 5 or more)", "=MIN(B14:H14,B16:H16)", "0.0", "textbook watchout: low expected counts (<5) make the test unreliable"),
    ("Cramer's V (strength of the link, 0 to 1)", "=SQRT(B20/I12)", "0.000", "extra, not in the class notes: SQRT(Chi-Sq / total). 0.1-0.3 = small to moderate"),
]
for i, (a, f, fmt, w) in enumerate(res, start=20):
    put(ws, f"A{i}", a, BOLD)
    put(ws, f"B{i}", f, fmt=fmt, fill=KEY if i in (20, 23) else None)
    put(ws, f"C{i}", w, SUB)
box(ws, 29, 1, 10, [
    "H0: the proportion of Class I recalls is the same for every root-cause area.",
    "H1: at least one proportion is different.",
], HYP, head="Hypotheses (alpha = 0.05)")
box(ws, 33, 1, 10, [None], HYP, head="Decision and conclusion")
ws["A34"] = ('=IF(B20>B22,"Chi-Sq ("&TEXT(B20,"0.0")&") > critical value ("&TEXT(B22,"0.00")&") and P-value < 0.05, so we reject H0: at least one root cause has a different Class I proportion. The greatest contribution is "&B24&" ("&TEXT(B25,"0.0")&"), which has "&TEXT(C13,"0.0%")&" Class I vs "&TEXT(I13,"0.0%")&" overall. Practical application: focus root-cause analysis and corrective action on design controls.",'
             '"Chi-Sq is below the critical value: do not reject H0.")')
ws.row_dimensions[34].height = 48
o = 37
put(ws, f"A{o}", "Chi-Square Test  (Minitab-style output)", MONOB, fill=OUT, box=False)
for j, h in enumerate(["", "Chi-Square", "DF", "P-Value"]):
    put(ws, f"{L(j+1)}{o+1}", h, MONOB, fill=OUT)
put(ws, f"A{o+2}", "Pearson", MONO, fill=OUT)
put(ws, f"B{o+2}", "=B20", MONO, "0.000", OUT)
put(ws, f"C{o+2}", "=B21", MONO, "0", OUT)
put(ws, f"D{o+2}", '=IF(B23<0.0005,"0.000",TEXT(B23,"0.000"))', MONO, fill=OUT)
widths(ws, {"A": 42, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 13, "J": 13})

# ================================================================ A3 Confidence Intervals
ws = wb.create_sheet("A3 Confidence Intervals")
title(ws, "A3. 95% confidence interval for each Class I proportion",
      "Minitab-style 'Test and CI for One Proportion' (normal approximation), one row per root-cause area.")
box(ws, 4, 1, 8, [
    "Minitab handout Step 11: the confidence level is set by alpha (95% = 1 - 0.05). Textbook: 'smaller sample sizes result in larger confidence intervals'.",
    "Formula: p ± Z x SQRT( p(1 - p) / N ), with Z = 1.96 for 95%. Minitab path: Stat > Basic Statistics > 1 Proportion > Summarized data (Method: normal approximation).",
], REF, head="Where this is in class")
header(ws, 8, ["Sample (root-cause area)", "X (Class I)", "N (recalls)", "Sample p", "Margin ±", "95% CI lower",
               "95% CI upper", "Compared with the overall rate"])
for k, c in enumerate(CATS, start=9):
    put(ws, f"A{k}", c, BLUE)
    put(ws, f"B{k}", f"=SUMIFS(ev_class1,ev_cat,A{k})", fmt="#,##0")
    put(ws, f"C{k}", f"=COUNTIFS(ev_cat,A{k})", fmt="#,##0")
    put(ws, f"D{k}", f"=B{k}/C{k}", fmt="0.00%")
    put(ws, f"E{k}", f"=Z*SQRT(D{k}*(1-D{k})/C{k})", fmt="0.00%")
    put(ws, f"F{k}", f"=MAX(0,D{k}-E{k})", fmt="0.00%")
    put(ws, f"G{k}", f"=D{k}+E{k}", fmt="0.00%")
    put(ws, f"H{k}", f'=IF(F{k}>$D$17,"Higher: whole CI above the overall rate",IF(G{k}<$D$17,"Lower: whole CI below the overall rate","Not clearly different"))')
put(ws, "A17", "All known-cause recalls", BOLD)
put(ws, "B17", "=SUM(B9:B15)", BOLD, "#,##0")
put(ws, "C17", "=SUM(C9:C15)", BOLD, "#,##0")
put(ws, "D17", "=B17/C17", BOLD, "0.00%", KEY)
box(ws, 19, 1, 8, [
    "Design (hardware): 11.9% ± 1.4 points, so the true rate is very likely between 10.5% and 13.3%. The whole interval is above the 6.7% average: clearly more dangerous.",
    "Use error has the widest interval (± 4.4 points) because it has only 134 recalls: smaller samples give wider intervals, as the textbook says.",
], HYP, head="Interpretation")
widths(ws, {"A": 28, "B": 12, "C": 12, "D": 11, "E": 11, "F": 13, "G": 13, "H": 40})

# ================================================================ A4 Logistic Regression
k_ = ev[ev["category"] != "Not determined"].copy()
REFL = {"category": "Production process", "devclass": "Class 2", "spec_model": "General Hospital", "period": "2013-2018"}
TERMS = [("category", "Root cause"), ("devclass", "Device class"), ("spec_model", "Specialty"), ("period", "Period")]
term = {v: f"C({v}, Treatment('{REFL[v]}'))" for v, _ in TERMS}
full = smf.logit("class1 ~ " + " + ".join(term.values()), data=k_).fit(disp=0, maxiter=200)
null = smf.logit("class1 ~ 1", data=k_).fit(disp=0)
drop = {}
for v, lab_ in TERMS:
    m = smf.logit("class1 ~ " + " + ".join(t for vv, t in term.items() if vv != v), data=k_).fit(disp=0, maxiter=200)
    drop[lab_] = (int(full.df_model - m.df_model), 2 * (full.llf - m.llf))
p_hat = full.predict(k_)
ranks = stats.rankdata(p_hat)
n1, n0 = int(k_["class1"].sum()), int((1 - k_["class1"]).sum())
auc = (ranks[k_["class1"].values == 1].sum() - n1 * (n1 + 1) / 2) / (n1 * n0)

ws = wb.create_sheet("A4 Logistic Regression")
title(ws, "A4. Binary logistic regression: Class I versus root cause, device class, specialty and period",
      "Laid out like Minitab output (Stat > Regression > Binary Logistic Regression): coefficients, odds ratios, model summary, deviance table, regression equation.")
box(ws, 4, 1, 9, [
    "Module 4 Correlation and Regression + Minitab handout: Coef, SE Coef, T-Value, P-Value and 'Regression Equation'. Same idea here, but the response Y is yes/no (Class I or not), so Minitab uses binary logistic regression.",
    "Your midterm Q5 (Challenger O-rings) asked for a model of the PROBABILITY of failure. A straight line can predict probabilities below 0 or above 1; logistic regression keeps them between 0 and 1.",
    "Coefficient test (as in the handout): H0: coefficient = 0 (no effect); H1: coefficient ≠ 0. Z-Value = Coef / SE Coef. If P-Value < 0.05, reject H0.",
    "Each categorical predictor has a reference level (coefficient 0). Odds ratio = EXP(Coef): how many times the odds of Class I are multiplied compared with the reference.",
], REF, head="Where this is in class")
put(ws, "A10", "Response Information:  Variable = class1 (Event = 1 = Class I).  Events used: 10,640 known-cause recalls; Class I = 717.", MONO, fill=OUT, box=False)

# Coefficients table
C0 = 12
put(ws, f"A{C0}", "Coefficients", MONOB, fill=OUT, box=False)
header(ws, C0 + 1, ["Lookup key", "Term", "Coef", "SE Coef", "95% CI low", "95% CI high", "Z-Value", "P-Value",
                    "Reject H0? (P < 0.05)"])
rows = [("Intercept", "Constant", full.params["Intercept"], full.bse["Intercept"])]
level_rows = {}
for v, lab_ in TERMS:
    levels = sorted(l for l in k_[v].unique() if l != REFL[v])
    rows.append((f"{lab_}|{REFL[v]}", f"{lab_}: {REFL[v]}  (reference)", 0.0, None))
    for lev in levels:
        nm = f"{term[v]}[T.{lev}]"
        rows.append((f"{lab_}|{lev}", f"{lab_}: {lev}", full.params[nm], full.bse[nm]))
r0 = C0 + 2
for i, (key, t, b, se) in enumerate(rows, start=r0):
    put(ws, f"A{i}", key, SUB)
    put(ws, f"B{i}", t, BOLD if se is None else BLACK)
    put(ws, f"C{i}", round(float(b), 6), BLUE, "0.0000")
    if se is not None:
        put(ws, f"D{i}", round(float(se), 6), BLUE, "0.0000")
        put(ws, f"E{i}", f"=C{i}-Z*D{i}", fmt="0.0000")
        put(ws, f"F{i}", f"=C{i}+Z*D{i}", fmt="0.0000")
        put(ws, f"G{i}", f"=C{i}/D{i}", fmt="0.00")
        put(ws, f"H{i}", f"=2*(1-NORMSDIST(ABS(G{i})))", fmt="0.000")
        put(ws, f"I{i}", f'=IF(H{i}<Alpha,"Yes - significant","No")')
    lab_ = t.split(":")[0]
    level_rows.setdefault(lab_, []).append((i, t.split(": ", 1)[-1].replace("  (reference)", ""), se is None))
r1 = r0 + len(rows) - 1
wb.defined_names["lr_tab"] = DefinedName("lr_tab", attr_text=f"'A4 Logistic Regression'!$A${r0}:$C${r1}")
wb.defined_names["lr_b0"] = DefinedName("lr_b0", attr_text=f"'A4 Logistic Regression'!$C${r0}")
for i in range(r0, r1 + 1):
    if "Design (hardware)" in str(ws[f"B{i}"].value) or "2023-2025" in str(ws[f"B{i}"].value):
        for col in "BCGH":
            ws[f"{col}{i}"].fill = KEY

# Odds ratios table
O0 = r1 + 3
put(ws, f"A{O0}", "Odds Ratios for Categorical Predictors", MONOB, fill=OUT, box=False)
header(ws, O0 + 1, ["", "Level A", "Level B (reference)", "Odds Ratio = EXP(Coef)", "95% CI low", "95% CI high",
                    "Events at Level A", "Class I observed", "Class I predicted (sum of P)"])
RNG = {"Root cause": "ev_cat", "Device class": "ev_devclass", "Specialty": "ev_specm", "Period": "ev_period"}
rr = O0 + 2
or_rows = {}
for v, lab_ in TERMS:
    put(ws, f"A{rr}", lab_, BOLD, box=False)
    rr += 1
    for (ci, lev, is_ref) in level_rows[lab_]:
        put(ws, f"B{rr}", lev)
        put(ws, f"C{rr}", REFL[v])
        if is_ref:
            put(ws, f"D{rr}", "1 (reference)", SUB)
        else:
            put(ws, f"D{rr}", f"=EXP(C{ci})", fmt="0.000",
                fill=KEY if lev in ("Design (hardware)", "2023-2025") else None)
            put(ws, f"E{rr}", f"=EXP(E{ci})", fmt="0.000")
            put(ws, f"F{rr}", f"=EXP(F{ci})", fmt="0.000")
        put(ws, f"G{rr}", f'=COUNTIFS({RNG[lab_]},B{rr},{KNOWN})', fmt="#,##0")
        put(ws, f"H{rr}", f'=SUMIFS(ev_class1,{RNG[lab_]},B{rr},{KNOWN})', fmt="#,##0")
        put(ws, f"I{rr}", f'=SUMIFS(ev_phat,{RNG[lab_]},B{rr})', fmt="0.0")
        or_rows[(lab_, lev)] = rr
        rr += 1

# Model summary + deviance table
M0 = rr + 2
put(ws, f"A{M0}", "Model Summary", MONOB, fill=OUT, box=False)
header(ws, M0 + 1, ["", "Log-likelihood (Excel, from Data col P)", "Log-likelihood (Python)", "Deviance R-Sq",
                    "AIC", "AUC (Python)"])
put(ws, f"B{M0+2}", "=SUM(ev_ll)", fmt="0.000", fill=KEY)
put(ws, f"C{M0+2}", round(float(full.llf), 3), BLUE, "0.000")
put(ws, f"D{M0+2}", f"=1-B{M0+2}/C{M0+6}", fmt="0.00%")
put(ws, f"E{M0+2}", f"=-2*B{M0+2}+2*{len([r for r in rows if r[3] is not None])}", fmt="0.0")
put(ws, f"F{M0+2}", round(float(auc), 3), BLUE, "0.000")
put(ws, f"B{M0+3}", "Matching Excel and Python log-likelihoods proves the coefficients are applied correctly to all 10,640 events.", SUB, box=False)
put(ws, f"B{M0+5}", "Null model (no predictors):", BOLD, box=False)
put(ws, f"B{M0+6}", "Log-likelihood of the null model", BOLD)
put(ws, f"C{M0+6}", f"=SUMIFS(ev_class1,{KNOWN})*LN(SUMIFS(ev_class1,{KNOWN})/COUNTIFS({KNOWN}))+(COUNTIFS({KNOWN})-SUMIFS(ev_class1,{KNOWN}))*LN(1-SUMIFS(ev_class1,{KNOWN})/COUNTIFS({KNOWN}))", fmt="0.000")
put(ws, f"D{M0+6}", "every recall gets the same chance, p-bar = 717 / 10,640", SUB, box=False)

D0 = M0 + 9
put(ws, f"A{D0}", "Deviance Table  (Deviance = -2 x log-likelihood; Adj Dev of a term = how much worse the model fits without it)", MONOB, fill=OUT, box=False)
header(ws, D0 + 1, ["", "Source", "DF", "Adj Dev", "Adj Mean", "Chi-Square", "P-Value", "Term matters? (P < 0.05)"])
dev_rows = [("Regression", int(full.df_model), 2 * (full.llf - null.llf))] + \
           [(lab_, drop[lab_][0], drop[lab_][1]) for _, lab_ in TERMS]
for i, (src, df_, dv) in enumerate(dev_rows, start=D0 + 2):
    put(ws, f"B{i}", ("  " if src != "Regression" else "") + src, BOLD)
    put(ws, f"C{i}", df_, BLUE, "0")
    put(ws, f"D{i}", round(float(dv), 3), BLUE, "0.00")
    put(ws, f"E{i}", f"=D{i}/C{i}", fmt="0.00")
    put(ws, f"F{i}", f"=D{i}", fmt="0.00")
    put(ws, f"G{i}", f"=_xlfn.CHISQ.DIST.RT(F{i},C{i})", fmt="0.000")
    put(ws, f"I{i}" if False else f"H{i}", f'=IF(G{i}<Alpha,"Yes","No")')
ws[f"D{D0+2}"].fill = KEY
e_ = D0 + 2 + len(dev_rows)
put(ws, f"B{e_}", "Error", BOLD)
put(ws, f"C{e_}", f"=COUNTIFS({KNOWN})-C{D0+2}-1", fmt="0")
put(ws, f"D{e_}", f"=-2*B{M0+2}", fmt="0.00")
put(ws, f"E{e_}", f"=D{e_}/C{e_}", fmt="0.000")
put(ws, f"B{e_+1}", "Total", BOLD)
put(ws, f"C{e_+1}", f"=COUNTIFS({KNOWN})-1", fmt="0")
put(ws, f"D{e_+1}", f"=-2*C{M0+6}", fmt="0.00")
put(ws, f"B{e_+2}", "Check: Regression Adj Dev = Total - Error", SUB, box=False)
put(ws, f"D{e_+2}", f"=D{e_+1}-D{e_}", fmt="0.00")

# Regression equation
Q_ = e_ + 4
put(ws, f"A{Q_}", "Regression Equation", MONOB, fill=OUT, box=False)
eq = [f"Y' = {full.params['Intercept']:.4f}"]
for v, lab_ in TERMS:
    for lev in sorted(l for l in k_[v].unique() if l != REFL[v]):
        b = full.params[f"{term[v]}[T.{lev}]"]
        eq.append(f"{'+' if b >= 0 else '-'} {abs(b):.4f} {lab_}_{lev}")
lines, cur = [], ""
for piece in eq:
    if len(cur) + len(piece) > 120:
        lines.append(cur)
        cur = "      "
    cur += piece + " "
lines.append(cur)
put(ws, f"A{Q_+1}", "P(Class I) = EXP(Y') / (1 + EXP(Y'))", MONO, fill=OUT, box=False)
for i, line in enumerate(lines, start=Q_ + 2):
    put(ws, f"A{i}", line, MONO, fill=OUT, box=False)
put(ws, f"A{Q_+2+len(lines)}", "Each name_level is 1 if the recall has that level, else 0. Reference levels (Production process, Class 2, General Hospital, 2013-2018) have no term.", SUB, box=False)

# What-if calculator
W0 = Q_ + 4 + len(lines)
put(ws, f"A{W0}", "Prediction: pick values from the drop-downs (like 'estimate Y when X = ...' in the class activities)", MONOB, fill=OUT, box=False)
header(ws, W0 + 1, ["", "Predictor", "Value (drop-down)", "Coef used"])
choices = [("Root cause", "Design (hardware)"), ("Device class", "Class 2"), ("Specialty", "General Hospital"), ("Period", "2023-2025")]
LIST_COL = 12   # lists of levels for the drop-downs live in columns L-O
for j, (lab_, _) in enumerate(choices):
    c_ = ws.cell(row=C0 + 1, column=LIST_COL + j, value=f"{lab_} levels")
    c_.font, c_.fill, c_.alignment = WHITE, HEAD, WRAP
    for t, (ci, lev, _) in enumerate(level_rows[lab_]):
        ws.cell(row=C0 + 2 + t, column=LIST_COL + j, value=lev)
for i, (lab_, dflt) in enumerate(choices, start=W0 + 2):
    put(ws, f"B{i}", lab_, BOLD)
    put(ws, f"C{i}", dflt, BLUE)
    put(ws, f"D{i}", f'=VLOOKUP(B{i}&"|"&C{i},lr_tab,3,0)', fmt="0.0000")
    j = i - W0 - 2
    col = L(LIST_COL + j)
    dv = DataValidation(type="list", formula1=f"=${col}${C0+2}:${col}${C0+1+len(level_rows[lab_])}")
    ws.add_data_validation(dv)
    dv.add(f"C{i}")
put(ws, f"B{W0+6}", "Constant", BOLD)
put(ws, f"D{W0+6}", "=lr_b0", fmt="0.0000")
put(ws, f"B{W0+7}", "Y' = sum of the coefficients", BOLD)
put(ws, f"D{W0+7}", f"=SUM(D{W0+2}:D{W0+6})", fmt="0.0000")
put(ws, f"B{W0+8}", "Predicted P(Class I) = EXP(Y')/(1+EXP(Y'))", BOLD)
put(ws, f"D{W0+8}", f"=EXP(D{W0+7})/(1+EXP(D{W0+7}))", fmt="0.0%", fill=KEY)

# Interpretation
I0 = W0 + 11
dz = or_rows[("Root cause", "Design (hardware)")]
pz = or_rows[("Period", "2023-2025")]
box(ws, I0, 1, 9, [
    "1. Root cause still matters after the controls: in the deviance table its P-Value is far below 0.05, so we reject H0 (no root-cause effect).",
    "2. Hardware design: odds ratio 2.01 (95% CI 1.63-2.48). On the same kind of device in the same period, a design recall has about twice the odds of being Class I compared with a production recall.",
    "3. Unadjusted (sheet A2 rates) the design odds ratio is 1.67; adjusted it is 2.01. Device type was hiding part of design's effect, so the finding is not explained by device type.",
    "4. Period 2023-2025: odds ratio 2.54. The rise in Class I is still there with root cause and device type held fixed - it is not just a change in the mix of recalls.",
    "5. Software (0.78) and supplier & material (0.98) are not significantly different from production (P > 0.05). Labeling & packaging is much less severe (0.27).",
    "6. Fit: Deviance R-Sq 21.6% and AUC 0.84 (the model ranks a Class I recall above a non-Class I recall 84% of the time). With firm-clustered standard errors (Python) the same terms stay significant.",
], HYP, head="Interpretation and conclusion")
widths(ws, {"A": 16, "B": 40, "C": 20, "D": 22, "E": 14, "F": 14, "G": 16, "H": 16, "I": 22,
            "L": 20, "M": 14, "N": 24, "O": 12})

# ================================================================ A5 Radiology
ws = wb.create_sheet("A5 Radiology")
title(ws, "A5. Pareto for imaging devices (FDA specialty = Radiology)", "Same Pareto as A1, filtered to one specialty.")
put(ws, "A5", "Root cause area", MONOB, fill=OUT)
RAD = ["Software", "Design (hardware)", "Production process", "Supplier & material", "Regulatory",
       "Labeling & packaging", "Use error"]
for j, c in enumerate(RAD):
    col = L(j + 2)
    put(ws, f"{col}5", c, MONOB, fill=OUT, align=WRAP)
    put(ws, f"{col}6", f'=COUNTIFS(ev_cat,{col}5,ev_spec,"Radiology")', MONO, "#,##0", OUT)
    put(ws, f"{col}7", f"={col}6/$I$6", MONO, "0.0%", OUT)
    put(ws, f"{col}8", f"=SUM($B$7:{col}7)", MONO, "0.0%", OUT)
    put(ws, f"{col}9", f'=SUMIFS(ev_class1,ev_cat,{col}5,ev_spec,"Radiology")', MONO, "0", OUT)
for rr_, t in ((6, "Count"), (7, "Percent"), (8, "Cum %"), (9, "Class I recalls")):
    put(ws, f"A{rr_}", t, MONOB, fill=OUT)
put(ws, "I5", "Total", MONOB, fill=OUT)
put(ws, "I6", "=SUM(B6:H6)", MONO, "#,##0", OUT)
put(ws, "I9", "=SUM(B9:H9)", MONO, "0", OUT)
ws.row_dimensions[5].height = 32
box(ws, 11, 1, 9, [
    "Software causes 46% of imaging recalls, but none of them were Class I. Production process causes most of the serious ones (13 of 18).",
    "For imaging makers, recalls are mainly a cost and downtime problem. Only 18 Class I recalls is too few to build an RPN for imaging.",
], HYP, head="Interpretation")
widths(ws, {"A": 18, **{L(c): 14 for c in range(2, 10)}})

# ================================================================ I1 FMEA (RPN)
ws = wb.create_sheet("I1 FMEA (RPN)")
title(ws, "I1. FMEA: risk priority number by root-cause area",
      "Same columns as the textbook's KIND Karz airbag FMEA (Table 6.7). Scores 1-10, RPN = SEV x OCC x DET.")
box(ws, 4, 1, 11, [
    "Textbook Sec. 6.5: 'The FMEA has three scores - Severity, Likelihood of Occurrence and Likelihood of Detection, scored 1-10. The product of all three is the Risk Priority Number.' RPNs above 100 need risk reduction.",
    "Difference from class: in class the team scored 1-10 by judgement (often 1, 5, 9). Here every score is calculated from the data, so anyone gets the same answer.",
    "The FDA data has no 'detection' information, so DET is replaced by TREND: how fast Class I recalls in that area are growing (a manager needs to know what is getting worse).",
], REF, head="Where this is in class")
header(ws, 9, ["ID", "Root-cause area (where the process fails)", "Failure Mode (examples from FDA labels)",
               "Failure Effects", "SEV", "OCC", "DET → TREND", "RPN", "Rank", "RPN > 100?",
               "Action: ISO 13485 clause to strengthen"], height=44)
FM = {
    "Production process": ("Process control, process design, equipment maintenance failures", "Out-of-spec devices shipped; recall",
                           "7.5 Production: process validation and control"),
    "Design (hardware)": ("Device design, component design or selection errors", "Device fails in use; harm to patient",
                          "7.3 Design: verification, validation, design reviews"),
    "Software": ("Software design or change-control errors", "Wrong output or device stops; recall",
                 "7.3 + IEC 62304 software life cycle"),
    "Supplier & material": ("Nonconforming or contaminated material, vendor changes", "Defective parts in finished devices",
                            "7.4 Purchasing: supplier control"),
    "Labeling & packaging": ("Labeling mix-ups, packaging process failures", "Wrong use or sterility loss",
                             "7.5.1 Labeling and packaging control"),
    "Use error": ("User mistakes the design did not prevent", "Misuse leading to harm", "7.3 + IEC 62366 usability"),
}
RC = list(FM)
# Calculation block (raw values) under the table
B0 = 20
header(ws, B0, ["", "Root-cause area", "Recalls / year, recent", "Class I / year, recent", "Class I / year, baseline",
                "SEV raw: Class I rate, recent", "OCC raw: recalls / year, recent", "TREND raw: growth in Class I / year"], height=44)
NB, NR = "(Base_End-Base_Start+1)", "(Recent_End-Recent_Start+1)"
for i, c in enumerate(RC, start=B0 + 1):
    put(ws, f"B{i}", c, BLUE)
    put(ws, f"C{i}", f'=COUNTIFS(ev_cat,B{i},ev_year,">="&Recent_Start,ev_year,"<="&Recent_End)/{NR}', fmt="0.0")
    put(ws, f"D{i}", f'=SUMIFS(ev_class1,ev_cat,B{i},ev_year,">="&Recent_Start,ev_year,"<="&Recent_End)/{NR}', fmt="0.0")
    put(ws, f"E{i}", f'=SUMIFS(ev_class1,ev_cat,B{i},ev_year,">="&Base_Start,ev_year,"<="&Base_End)/{NB}', fmt="0.0")
    put(ws, f"F{i}", f"=D{i}/C{i}", fmt="0.0%")
    put(ws, f"G{i}", f"=C{i}", fmt="0.0")
    put(ws, f"H{i}", f"=D{i}/E{i}-1", fmt="0%")
bl, bh = B0 + 1, B0 + len(RC)
put(ws, f"B{bh+1}", "Score = 1 + 9 x (value - MIN) / (MAX - MIN): lowest area scores 1, highest scores 10.", SUB, box=False)
for i, c in enumerate(RC, start=10):
    s = B0 + 1 + (i - 10)
    put(ws, f"A{i}", f"{i-9}")
    put(ws, f"B{i}", c, BOLD, align=LEFTWRAP)
    put(ws, f"C{i}", FM[c][0], align=LEFTWRAP)
    put(ws, f"D{i}", FM[c][1], align=LEFTWRAP)
    for col, raw in (("E", "F"), ("F", "G"), ("G", "H")):
        put(ws, f"{col}{i}", f"=1+9*({raw}{s}-MIN({raw}${bl}:{raw}${bh}))/(MAX({raw}${bl}:{raw}${bh})-MIN({raw}${bl}:{raw}${bh}))", fmt="0.0")
    put(ws, f"H{i}", f"=E{i}*F{i}*G{i}", BOLD, "0", KEY)
    put(ws, f"I{i}", f"=RANK(H{i},$H$10:$H$15)")
    put(ws, f"J{i}", f'=IF(H{i}>100,"Yes - act","No")')
    put(ws, f"K{i}", FM[c][2], align=LEFTWRAP)
    ws.row_dimensions[i].height = 32
put(ws, "B16", "Regulatory recalls (missing FDA paperwork) are left out: they are not product failures and had no recent Class I recalls.", SUB, box=False)
box(ws, B0 + len(RC) + 3, 1, 11, [
    "Production process (RPN 232), hardware design (182) and software (152) are above 100: these need action first.",
    "Production ranks first because it is by far the most frequent; design is the most severe (1 in 5 recent design recalls is Class I).",
    "Labeling's TREND is 10 because Class I labeling recalls went from 0.7 to 4 a year: a big percentage on a small base. Its low SEV keeps it at rank 5.",
    "Textbook tip: SEV should rarely be reduced when rescoring - improvement lowers OCC (fewer recalls) or DET/TREND, not the severity of the failure itself.",
], HYP, head="Interpretation")
widths(ws, {"A": 5, "B": 24, "C": 34, "D": 28, "E": 14, "F": 14, "G": 16, "H": 12, "I": 7, "J": 11, "K": 36})

# ================================================================ I2 Impact
ws = wb.create_sheet("I2 Impact")
title(ws, "I2. Impact: what improving the top two areas would be worth",
      "A scenario (pro forma estimate, textbook Sec. 6.4), not a forecast. Change the blue percentages.")
put(ws, "A4", "Class I recalls per year, recent, top 2 RPN areas", BOLD)
put(ws, "B4", "=SUMPRODUCT(('I1 FMEA (RPN)'!I10:I15<=2)*'I1 FMEA (RPN)'!D21:D26)", fmt="0.0")
put(ws, "A5", "Class I recalls per year, recent, all known causes", BOLD)
put(ws, "B5", f'=SUMIFS(ev_class1,{KNOWN},ev_year,">="&Recent_Start,ev_year,"<="&Recent_End)/(Recent_End-Recent_Start+1)', fmt="0.0")
put(ws, "A6", "Top 2 share of Class I recalls", BOLD)
put(ws, "B6", "=B4/B5", fmt="0%", fill=KEY)
header(ws, 8, ["Reduction in the top 2 areas", "Class I recalls avoided per year", "Share of all known-cause Class I"])
for k, cut in enumerate((0.10, 0.25, 0.50), start=9):
    put(ws, f"A{k}", cut, BLUE, "0%")
    put(ws, f"B{k}", f"=$B$4*A{k}", fmt="0.0", fill=KEY if cut == 0.25 else None)
    put(ws, f"C{k}", f"=B{k}/$B$5", fmt="0.0%")
widths(ws, {"A": 50, "B": 30, "C": 32})

# ================================================================ C1 Control Plan
ws = wb.create_sheet("C1 Control Plan")
title(ws, "C1. Control: keeping the gains", "Six Sigma slides, DMAIC - Control: 'Put tools in place to ensure the key variables remain within the maximum acceptable ranges.'")
header(ws, 4, ["What to control", "Tool", "How often", "Reaction plan if it signals"])
plan = [
    ("Share of recalls that are Class I", "p chart (sheet M2), limits from Phase I", "Every quarter, from new openFDA data",
     "Point above UCL or 8 quarters in a row above P bar: investigate special cause, review root causes of that quarter's Class I recalls"),
    ("Production process recalls", "Count per quarter + RPN refresh (sheet I1)", "Every quarter",
     "If OCC rises: audit process validation and change control (ISO 13485 clause 7.5)"),
    ("Design recalls", "Class I rate for design (sheet A3)", "Every quarter",
     "If the CI moves above the current rate: review design verification and validation (clause 7.3)"),
    ("Recalls closed with no root cause", "Not-determined share (sheet A1)", "Every quarter",
     "If above 20%: require root-cause investigation before a recall is closed (clause 8.5.2)"),
]
for i, rw in enumerate(plan, start=5):
    for j, v in enumerate(rw):
        put(ws, f"{L(j+1)}{i}", v, BOLD if j == 0 else BLACK, align=LEFTWRAP)
    ws.row_dimensions[i].height = 48
box(ws, 11, 1, 4, [
    "To refresh: download the two new FDA files, run step1, step2 and step4. Every sheet recalculates from the new Data sheet.",
], REF, head="How to update")
widths(ws, {"A": 32, "B": 36, "C": 28, "D": 70})

for sh in wb.worksheets:
    for chrt in sh._charts:
        for ax in (chrt.x_axis, chrt.y_axis):
            ax.delete = False
        for sub in getattr(chrt, "_charts", []):
            sub.y_axis.delete = False
wb.save("outputs/Recall_Analysis_ISE250_Style.xlsx")
print("Saved outputs/Recall_Analysis_ISE250_Style.xlsx")
