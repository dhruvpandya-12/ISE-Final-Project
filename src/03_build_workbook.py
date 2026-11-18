from pathlib import Path

import pandas as pd
import statsmodels.formula.api as smf
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from scipy import stats

ROOT = Path(__file__).resolve().parents[1]
ev = pd.read_csv(ROOT / "outputs" / "study_events.csv", dtype={"event_id": str})
ev["quarter"] = pd.PeriodIndex(pd.to_datetime(ev["date"], errors="coerce"), freq="Q").astype(str)


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
REF = PatternFill("solid", fgColor="E2EFDA")
HYP = PatternFill("solid", fgColor="FFF2CC")
OUT = PatternFill("solid", fgColor="F2F2F2")
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


ws = wb.active
ws.title = "Settings"
title(ws, "Settings", "Study periods and test parameters used by every sheet.")
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


ws = wb.create_sheet("Data")
title(ws, "Data: 13,410 recall events, 2013-2025", "One row per recall event. class1 = 1 if the event is FDA Class I.")
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
ch = BarChart()
ch.title, ch.y_axis.title = "Class I recall events per year", "Class I events"
ch.add_data(Reference(ws, min_col=3, min_row=4, max_row=17), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=17))
ch.height, ch.width, ch.legend = 8, 16, None
ws.add_chart(ch, "F4")
widths(ws, {"A": 36, "B": 14, "C": 14, "D": 14})


ws = wb.create_sheet("M1 Two Proportions")
title(ws, "M1. Test of two proportions: is the Class I share higher now?",
      "Recent period versus baseline period.")
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

header(ws, 20, ["Calculation", "Value"])
steps = [
    ("Difference p1 - p2", "=B13-C13", "0.0000"),
    ("Pooled proportion p-hat", "=(B11+C11)/(B12+C12)", "0.0000"),
    ("Standard error (pooled)", "=SQRT(B22*(1-B22)*(1/B12+1/C12))", "0.00000"),
    ("Z-Value", "=B21/B23", "0.00"),
    ("P-Value (two-sided)", "=2*(1-NORMSDIST(ABS(B24)))", "0.0E+00"),
    ("Standard error (unpooled, for the CI)", "=SQRT(B13*(1-B13)/B12+C13*(1-C13)/C12)", "0.00000"),
    ("95% CI for difference: lower", "=B21-Z*B26", "0.0000"),
    ("95% CI for difference: upper", "=B21+Z*B26", "0.0000"),
]
for i, (a, f, fmt) in enumerate(steps, start=21):
    put(ws, f"A{i}", a, BOLD)
    put(ws, f"B{i}", f, fmt=fmt, fill=KEY if i in (24, 25) else None)


o = 31
put(ws, f"A{o}", "Test and CI for Two Proportions", MONOB, fill=OUT, box=False)
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
widths(ws, {"A": 44, "B": 22, "C": 52, "D": 14, "E": 12, "F": 12})


ws = wb.create_sheet("M2 p Chart")
title(ws, "M2. p chart: share of recalls that are Class I, by quarter",
      "Phase I limits from 2013-2018; variable limits for each quarter's sample size.")
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

header(ws, 18, ["Control check", "", "Result"], col=12, height=18)
ws.merge_cells("L18:M18")
checks = [
    ("1. Points outside the control limits", f'=COUNTIF(I{Q0}:I{Q1},"Yes")+COUNTIF(I{Q0}:I{Q1},"Below LCL")&" points outside the limits"'),
    ("   in Phase I (2013-2018)", f'=COUNTIFS(I{Q0}:I{Q1},"Yes",B{Q0}:B{Q1},"<="&Phase1_End)&" points"'),
    ("   in 2023-2025", f'=COUNTIFS(I{Q0}:I{Q1},"Yes",B{Q0}:B{Q1},">=2023")&" of "&COUNTIFS(B{Q0}:B{Q1},">=2023")&" quarters above UCL"'),
        ("2. Points above vs below the centerline", f'=COUNTIF(J{Q0}:J{Q1},">0")&" above vs "&COUNTIF(J{Q0}:J{Q1},"<0")&" below"'),
    ("3. ... no run of 8 or more on one side of P bar", f'="Longest run above P bar: "&MAX(J{Q0}:J{Q1})&" quarters in a row"'),
]
for i, (a, f) in enumerate(checks, start=19):
    put(ws, f"L{i}", a, BOLD)
    ws.merge_cells(f"L{i}:M{i}")
    put(ws, f"N{i}", f)
box(ws, 4, 12, 3, [None], HYP, head="Conclusion")
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


ws = wb.create_sheet("A1 Pareto")
title(ws, "A1. Pareto chart of recall root causes",
      "Events with a known root cause, 2013-2025. Bars from highest to lowest count; line = cumulative %.")
put(ws, "A8", "Pareto table", MONOB, fill=OUT, box=False)
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


ws = wb.create_sheet("A2 Chi-Square")
title(ws, "A2. Chi-square test of multiple proportions: is one root cause more dangerous?",
      "Class I versus not Class I, by root-cause area.")
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

header(ws, 19, ["Test result", "Value"])
res = [
    ("Total Chi-Sq", "=I15+I17", "0.0"),
    ("DF = number of proportions - 1", "=COUNTA(B9:H9)-1", "0"),
    ("Critical Chi-Sq statistic (alpha = 0.05)", "=_xlfn.CHISQ.INV.RT(Alpha,B21)", "0.00"),
    ("P-Value", "=_xlfn.CHISQ.DIST.RT(B20,B21)", "0.0E+00"),
    ("Largest contribution to Chi-Sq", '=INDEX(B9:H9,MATCH(MAX(B15:H15),B15:H15,0))', None),
    ("   its contribution", "=MAX(B15:H15)", "0.0"),
    ("Smallest expected count (must be 5 or more)", "=MIN(B14:H14,B16:H16)", "0.0"),
    ("Cramer's V ", "=SQRT(B20/I12)", "0.000"),
]
for i, (a, f, fmt) in enumerate(res, start=20):
    put(ws, f"A{i}", a, BOLD)
    put(ws, f"B{i}", f, fmt=fmt, fill=KEY if i in (20, 23) else None)
box(ws, 29, 1, 10, [
    "H0: the proportion of Class I recalls is the same for every root-cause area.",
    "H1: at least one proportion is different.",
], HYP, head="Hypotheses (alpha = 0.05)")
box(ws, 33, 1, 10, [None], HYP, head="Decision and conclusion")
ws["A34"] = ('=IF(B20>B22,"Chi-Sq ("&TEXT(B20,"0.0")&") > critical value ("&TEXT(B22,"0.00")&") and P-value < 0.05, so we reject H0: at least one root cause has a different Class I proportion. The greatest contribution is "&B24&" ("&TEXT(B25,"0.0")&"), which has "&TEXT(C13,"0.0%")&" Class I vs "&TEXT(I13,"0.0%")&" overall. Practical application: focus root-cause analysis and corrective action on design controls.",'
             '"Chi-Sq is below the critical value: do not reject H0.")')
ws.row_dimensions[34].height = 48
o = 37
put(ws, f"A{o}", "Chi-Square Test", MONOB, fill=OUT, box=False)
for j, h in enumerate(["", "Chi-Square", "DF", "P-Value"]):
    put(ws, f"{L(j+1)}{o+1}", h, MONOB, fill=OUT)
put(ws, f"A{o+2}", "Pearson", MONO, fill=OUT)
put(ws, f"B{o+2}", "=B20", MONO, "0.000", OUT)
put(ws, f"C{o+2}", "=B21", MONO, "0", OUT)
put(ws, f"D{o+2}", '=IF(B23<0.0005,"0.000",TEXT(B23,"0.000"))', MONO, fill=OUT)
widths(ws, {"A": 42, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 13, "J": 13})


ws = wb.create_sheet("A3 Confidence Intervals")
title(ws, "A3. 95% confidence interval for each Class I proportion",
      "Normal approximation, one row per root-cause area.")
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
widths(ws, {"A": 28, "B": 12, "C": 12, "D": 11, "E": 11, "F": 13, "G": 13, "H": 40})
ch = BarChart()
ch.type, ch.title, ch.x_axis.title = "bar", "Class I rate by root-cause area", "Class I rate"
ch.add_data(Reference(ws, min_col=4, min_row=8, max_row=15), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=9, max_row=15))
ch.series[0].graphicalProperties.solidFill = "C00000"
ch.x_axis.number_format, ch.height, ch.width, ch.legend = "0%", 8, 18, None
ws.add_chart(ch, "J4")


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
      "Known-cause events. Coefficients estimated by maximum likelihood (Python statsmodels).")
put(ws, "A10", f'="Response: class1 (1 = Class I).  Events used: "&TEXT(COUNTIFS({KNOWN}),"#,##0")&";  Class I events: "&TEXT(SUMIFS(ev_class1,{KNOWN}),"#,##0")', MONO, fill=OUT, box=False)


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


M0 = rr + 2
put(ws, f"A{M0}", "Model Summary", MONOB, fill=OUT, box=False)
header(ws, M0 + 1, ["", "Log-likelihood (Excel)", "Log-likelihood (Python)", "Deviance R-Sq",
                    "AIC", "AUC (Python)"])
put(ws, f"B{M0+2}", "=SUM(ev_ll)", fmt="0.000", fill=KEY)
put(ws, f"C{M0+2}", round(float(full.llf), 3), BLUE, "0.000")
put(ws, f"D{M0+2}", f"=1-B{M0+2}/C{M0+6}", fmt="0.00%")
put(ws, f"E{M0+2}", f"=-2*B{M0+2}+2*{len([r for r in rows if r[3] is not None])}", fmt="0.0")
put(ws, f"F{M0+2}", round(float(auc), 3), BLUE, "0.000")
put(ws, f"B{M0+5}", "Null model (no predictors):", BOLD, box=False)
put(ws, f"B{M0+6}", "Log-likelihood of the null model", BOLD)
put(ws, f"C{M0+6}", f"=SUMIFS(ev_class1,{KNOWN})*LN(SUMIFS(ev_class1,{KNOWN})/COUNTIFS({KNOWN}))+(COUNTIFS({KNOWN})-SUMIFS(ev_class1,{KNOWN}))*LN(1-SUMIFS(ev_class1,{KNOWN})/COUNTIFS({KNOWN}))", fmt="0.000")

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
    put(ws, f"H{i}", f'=IF(G{i}<Alpha,"Yes","No")')
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


W0 = Q_ + 4 + len(lines)
put(ws, f"A{W0}", "Prediction (select values from the drop-downs)", MONOB, fill=OUT, box=False)
header(ws, W0 + 1, ["", "Predictor", "Value (drop-down)", "Coef used"])
choices = [("Root cause", "Design (hardware)"), ("Device class", "Class 2"), ("Specialty", "General Hospital"), ("Period", "2023-2025")]
LIST_COL = 12
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


rc = [r_ for (lab_, lev), r_ in or_rows.items() if lab_ == "Root cause" and lev != REFL["category"]]
ch = BarChart()
ch.type, ch.title = "bar", "Adjusted odds ratio of Class I vs production process"
ch.add_data(Reference(ws, min_col=4, min_row=min(rc), max_row=max(rc)), titles_from_data=False)
ch.set_categories(Reference(ws, min_col=2, min_row=min(rc), max_row=max(rc)))
ch.series[0].graphicalProperties.solidFill = "2A6FDB"
ch.height, ch.width, ch.legend = 8, 18, None
ws.add_chart(ch, f"K{O0}")
widths(ws, {"A": 16, "B": 40, "C": 20, "D": 22, "E": 14, "F": 14, "G": 16, "H": 16, "I": 22,
            "L": 20, "M": 14, "N": 24, "O": 12})


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
widths(ws, {"A": 18, **{L(c): 14 for c in range(2, 10)}})


ws = wb.create_sheet("I1 FMEA (RPN)")
title(ws, "I1. FMEA: risk priority number by root-cause area",
      "Scores 1-10 by min-max scaling. RPN = SEV x OCC x TREND.")
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
for i, c in enumerate(RC, start=10):
    s = B0 + 1 + (i - 10)
    put(ws, f"A{i}", i - 9)
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
ch = BarChart()
ch.type, ch.title = "bar", "Risk priority number by root-cause area"
ch.add_data(Reference(ws, min_col=8, min_row=9, max_row=15), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=10, max_row=15))
ch.series[0].graphicalProperties.solidFill = "1F3864"
ch.height, ch.width, ch.legend = 8, 18, None
ws.add_chart(ch, "M9")
widths(ws, {"A": 5, "B": 24, "C": 34, "D": 28, "E": 14, "F": 14, "G": 16, "H": 12, "I": 7, "J": 11, "K": 36})


ws = wb.create_sheet("I2 Impact")
title(ws, "I2. Impact: what improving the top two areas would be worth",
      "Scenario estimate for the two highest-RPN areas.")
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


ws = wb.create_sheet("C1 Control Plan")
title(ws, "C1. Control: keeping the gains", "Monitoring and reaction plan.")
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
widths(ws, {"A": 32, "B": 36, "C": 28, "D": 70})

for sh in wb.worksheets:
    for chrt in sh._charts:
        for ax in (chrt.x_axis, chrt.y_axis):
            ax.delete = False
        for sub in getattr(chrt, "_charts", []):
            sub.y_axis.delete = False
XL = ROOT / "excel"
XL.mkdir(exist_ok=True)
wb.save(XL / "Medical_Device_Recall_Analysis.xlsx")
print("Saved", XL / "Medical_Device_Recall_Analysis.xlsx")
