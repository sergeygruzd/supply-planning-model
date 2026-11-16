"""Builds PlanningDemo-MapleRidgeFoods.xlsx: a fictional manufacturer's planning model driven by live Excel formulas.

Change Inputs!C4 (demand scenario %) or any forecast cell, and production, component requirements,
purchase orders (offset by supplier lead time, rounded to MOQ), shortage alerts and cash all recalculate.
"""
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1] if len(sys.argv) > 1 else "PlanningDemo-MapleRidgeFoods.xlsx"
MONTHS = ["Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27", "Aug-27", "Sep-27", "Oct-27", "Nov-27", "Dec-27"]
N = len(MONTHS)
FC, LC = 3, 3 + N - 1          # first/last month column (C..N)
MC = [L(c) for c in range(FC, LC + 1)]

PRODUCTS = [  # code, name, monthly forecast units
    ("OB-H", "Oat Bar, Honey (each)", [42000, 40000, 45000, 48000, 52000, 55000, 50000, 47000, 53000, 58000, 60000, 62000]),
    ("OB-M", "Oat Bar, Maple (each)", [30000, 28000, 32000, 34000, 33000, 31000, 30000, 33000, 42000, 50000, 55000, 52000]),
    ("GR-500", "Granola 500 g (bag)", [9000, 8500, 9500, 10000, 10500, 9800, 9200, 9600, 11500, 12800, 14000, 13500]),
]
COMPONENTS = [  # code, name, unit, cost, lead time (months), on hand, safety stock, MOQ, payment terms (months)
    ("OATS", "Rolled oats", "kg", 1.10, 1, 9000, 2000, 5000, 1),
    ("HONEY", "Honey", "kg", 6.50, 1, 1200, 300, 500, 1),
    ("MAPLE", "Maple syrup", "L", 9.00, 2, 1400, 250, 400, 1),
    ("ALMOND", "Almonds", "kg", 11.00, 2, 2500, 400, 1000, 2),
    ("WRAP", "Bar wrapper", "each", 0.03, 3, 320000, 80000, 250000, 1),
    ("CASE", "Shipper case", "each", 0.45, 1, 8000, 3000, 5000, 1),
    ("POUCH", "Granola pouch", "each", 0.12, 3, 45000, 10000, 50000, 1),
]
BOM = {  # product -> {component: qty per unit}
    "OB-H": {"OATS": 0.03, "HONEY": 0.008, "ALMOND": 0.005, "WRAP": 1, "CASE": 1 / 24},
    "OB-M": {"OATS": 0.03, "MAPLE": 0.007, "ALMOND": 0.005, "WRAP": 1, "CASE": 1 / 24},
    "GR-500": {"OATS": 0.35, "HONEY": 0.04, "MAPLE": 0.03, "ALMOND": 0.06, "POUCH": 1, "CASE": 1 / 12},
}

NAVY, STEEL, PALE, INPUT, ALERT = "1F3A6E", "2F4A5C", "EEF2F7", "FFF4CC", "FDE2E1"
H = Font(bold=True, color="FFFFFF")
B = Font(bold=True)
thin = Side(style="thin", color="D0D5DD")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
INPUT_FILL = PatternFill("solid", fgColor=INPUT)


def header(ws, row, labels, start=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=t)
        c.font, c.fill, c.alignment = H, PatternFill("solid", fgColor=NAVY), Alignment(horizontal="center", wrap_text=True)


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    ws["A2"] = sub
    ws["A2"].font = Font(italic=True, color="5B625E")


def month_index_row(ws, row):
    """Hidden helper row with 1..12 so formulas can shift by lead time."""
    ws.cell(row=row, column=2, value="month #").font = Font(color="999999", size=8)
    for i, c in enumerate(MC):
        ws[f"{c}{row}"] = i + 1
        ws[f"{c}{row}"].font = Font(color="999999", size=8)


wb = Workbook()

# ---------------- README ----------------
ws = wb.active
ws.title = "Start here"
title(ws, "Maple Ridge Foods: planning model (DEMO, fictional data)", "Built by Sergey Gruzd to show how a forecast-driven planning model works")
lines = [
    "",
    "HOW TO USE",
    "1. Go to 'Inputs' and change the yellow cells: the demand scenario % (C4) or any monthly forecast.",
    "2. Everything else recalculates automatically: Production, Requirements, Purchasing (MRP), Cash, Dashboard.",
    "3. Try +25% on Inputs!C4 and watch the shortage alerts and cash needs on the Dashboard.",
    "",
    "WHAT EACH SHEET DOES",
    "Inputs: demand forecast by product and month, plus one scenario lever.",
    "BOM: how much of each component goes into one unit of each product.",
    "Components: cost, supplier lead time, stock on hand, safety stock, minimum order quantity, payment terms.",
    "Production: planned production = forecast x scenario.",
    "Requirements: component needs per month, exploded through the BOM.",
    "Purchasing: month-by-month projected stock, planned receipts (rounded up to MOQ), and WHEN to place each",
    "            purchase order (receipt month minus lead time). Receipts needed sooner than the lead time are flagged.",
    "Cash: supplier payments by month, using each component's payment terms.",
    "Dashboard: the one page for the monthly planning meeting.",
    "",
    "All data is invented for demonstration. No client information is included.",
]
for i, t in enumerate(lines, start=3):
    ws.cell(row=i, column=1, value=t)
    if t in ("HOW TO USE", "WHAT EACH SHEET DOES"):
        ws.cell(row=i, column=1).font = B
ws.column_dimensions["A"].width = 110

# ---------------- Inputs ----------------
wi = wb.create_sheet("Inputs")
title(wi, "Inputs: change the yellow cells", "Units per month")
wi["B4"] = "Demand scenario vs. forecast"
wi["B4"].font = B
wi["C4"] = 0.0
wi["C4"].number_format = "+0%;-0%;0%"
wi["C4"].fill, wi["C4"].border = INPUT_FILL, BOX
header(wi, 6, ["Code", "Product"] + MONTHS + ["Total"])
for r, (code, name, fc) in enumerate(PRODUCTS, start=7):
    wi.cell(row=r, column=1, value=code)
    wi.cell(row=r, column=2, value=name)
    for i, v in enumerate(fc):
        c = wi.cell(row=r, column=FC + i, value=v)
        c.fill, c.border, c.number_format = INPUT_FILL, BOX, "#,##0"
    wi.cell(row=r, column=LC + 1, value=f"=SUM({MC[0]}{r}:{MC[-1]}{r})").number_format = "#,##0"
P_ROWS = list(range(7, 7 + len(PRODUCTS)))
wi.column_dimensions["B"].width = 26

# ---------------- BOM ----------------
wbm = wb.create_sheet("BOM")
title(wbm, "Bill of materials", "Quantity of each component per ONE unit of product")
header(wbm, 4, ["Product"] + [c[1] + f" ({c[2]})" for c in COMPONENTS])
for r, (code, name, _) in enumerate(PRODUCTS, start=5):
    wbm.cell(row=r, column=1, value=code)
    for j, comp in enumerate(COMPONENTS, start=2):
        c = wbm.cell(row=r, column=j, value=BOM[code].get(comp[0], 0))
        c.number_format, c.fill, c.border = "0.000", INPUT_FILL, BOX
BOM_ROWS = (5, 4 + len(PRODUCTS))
for j in range(1, len(COMPONENTS) + 2):
    wbm.column_dimensions[L(j)].width = 16

# ---------------- Components ----------------
wc = wb.create_sheet("Components")
title(wc, "Component master", "Supplier terms and stock policy")
header(wc, 4, ["Code", "Component", "Unit", "Unit cost ($)", "Lead time (months)", "On hand", "Safety stock", "Min order qty", "Pay terms (months)"])
for r, comp in enumerate(COMPONENTS, start=5):
    for j, v in enumerate(comp, start=1):
        c = wc.cell(row=r, column=j, value=v)
        if j >= 4:
            c.fill, c.border = INPUT_FILL, BOX
            c.number_format = "$#,##0.00" if j == 4 else "#,##0"
COMP_ROW = {comp[0]: r for r, comp in enumerate(COMPONENTS, start=5)}
for j, w in zip("ABCDEFGHI", (10, 18, 8, 12, 12, 12, 12, 13, 12)):
    wc.column_dimensions[j].width = w

# ---------------- Production ----------------
wp = wb.create_sheet("Production")
title(wp, "Production plan", "= forecast x (1 + scenario), rounded to whole units")
header(wp, 4, ["Code", "Product"] + MONTHS + ["Total"])
for k, (code, name, _) in enumerate(PRODUCTS):
    r = 5 + k
    wp.cell(row=r, column=1, value=code)
    wp.cell(row=r, column=2, value=name)
    for c in MC:
        wp[f"{c}{r}"] = f"=ROUND(Inputs!{c}{P_ROWS[k]}*(1+Inputs!$C$4),0)"
        wp[f"{c}{r}"].number_format = "#,##0"
    wp.cell(row=r, column=LC + 1, value=f"=SUM({MC[0]}{r}:{MC[-1]}{r})").number_format = "#,##0"
PROD_ROWS = (5, 4 + len(PRODUCTS))
wp.column_dimensions["B"].width = 26

# ---------------- Requirements ----------------
wr = wb.create_sheet("Requirements")
title(wr, "Component requirements", "Production exploded through the BOM (SUMPRODUCT)")
header(wr, 4, ["Code", "Component"] + MONTHS + ["Total"])
REQ_ROW = {}
for j, comp in enumerate(COMPONENTS):
    r = 5 + j
    REQ_ROW[comp[0]] = r
    bcol = L(2 + j)
    wr.cell(row=r, column=1, value=comp[0])
    wr.cell(row=r, column=2, value=f"{comp[1]} ({comp[2]})")
    for c in MC:
        wr[f"{c}{r}"] = f"=SUMPRODUCT(BOM!${bcol}${BOM_ROWS[0]}:${bcol}${BOM_ROWS[1]},Production!{c}${PROD_ROWS[0]}:{c}${PROD_ROWS[1]})"
        wr[f"{c}{r}"].number_format = "#,##0"
    wr.cell(row=r, column=LC + 1, value=f"=SUM({MC[0]}{r}:{MC[-1]}{r})").number_format = "#,##0"
wr.column_dimensions["B"].width = 24

# ---------------- Purchasing (MRP) ----------------
wm = wb.create_sheet("Purchasing")
title(wm, "Purchasing plan (MRP)", "Receipts keep stock above safety stock, rounded up to MOQ. 'Place PO' = receipt month minus lead time.")
month_index_row(wm, 3)
header(wm, 4, ["Component", "Line", "Start"] + MONTHS[:-1] + [MONTHS[-1]], start=1)
# shift header so months sit in C..N: column C is month 1; put 'Start' in column B label row instead
for i, m in enumerate(MONTHS):
    wm.cell(row=4, column=FC + i, value=m)
wm.cell(row=4, column=1, value="Component")
wm.cell(row=4, column=2, value="Line / start")
wm.cell(row=4, column=LC + 1, value="Total")
MRP = {}
row = 5
for comp in COMPONENTS:
    code, name, unit = comp[0], comp[1], comp[2]
    cr = COMP_ROW[code]
    lt, oh, ss, moq = f"Components!$E${cr}", f"Components!$F${cr}", f"Components!$G${cr}", f"Components!$H${cr}"
    g, rec, proj, po = row + 1, row + 2, row + 3, row + 4
    MRP[code] = {"rec": rec, "po": po, "proj": proj}
    wm.cell(row=row, column=1, value=f"{name} ({unit})").font = B
    wm.cell(row=row, column=2, value=f"=\"lead time \"&{lt}&\" mo, MOQ \"&TEXT({moq},\"#,##0\")")
    wm.cell(row=row, column=2).font = Font(italic=True, color="5B625E")
    wm.cell(row=g, column=1, value="Gross requirement")
    wm.cell(row=rec, column=1, value="Planned receipt")
    wm.cell(row=proj, column=1, value="Projected on hand")
    wm.cell(row=proj, column=2, value=f"={oh}")
    wm.cell(row=po, column=1, value="Place PO (qty)")
    for i, c in enumerate(MC):
        prev = f"$B{proj}" if i == 0 else f"{MC[i-1]}{proj}"
        wm[f"{c}{g}"] = f"=Requirements!{c}{REQ_ROW[code]}"
        wm[f"{c}{rec}"] = f"=IF({prev}-{c}{g}<{ss},CEILING({ss}-({prev}-{c}{g}),{moq}),0)"
        wm[f"{c}{proj}"] = f"={prev}+{c}{rec}-{c}{g}"
        wm[f"{c}{po}"] = f"=IF({c}$3+{lt}<={N},INDEX(${MC[0]}{rec}:${MC[-1]}{rec},{c}$3+{lt}),0)"
        for rr in (g, rec, proj, po):
            wm[f"{c}{rr}"].number_format = "#,##0"
        wm[f"{c}{po}"].font = B
    wm.cell(row=proj, column=2).number_format = "#,##0"
    for rr in (g, rec, po):
        wm.cell(row=rr, column=LC + 1, value=f"=SUM({MC[0]}{rr}:{MC[-1]}{rr})").number_format = "#,##0"
    # receipts needed inside the lead time can't be ordered in time -> highlight
    wm.conditional_formatting.add(f"{MC[0]}{rec}:{MC[-1]}{rec}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=PALE)))
    row += 6
wm.column_dimensions["A"].width = 22
wm.column_dimensions["B"].width = 26

# ---------------- Cash ----------------
wca = wb.create_sheet("Cash")
title(wca, "Supplier payments", "Receipt value paid after each component's payment terms")
month_index_row(wca, 3)
header(wca, 4, ["Code", "Component"] + MONTHS + ["Total"])
for j, comp in enumerate(COMPONENTS):
    r = 5 + j
    code = comp[0]
    cr = COMP_ROW[code]
    rec = MRP[code]["rec"]
    wca.cell(row=r, column=1, value=code)
    wca.cell(row=r, column=2, value=comp[1])
    for c in MC:
        wca[f"{c}{r}"] = (f"=IF({c}$3-Components!$I${cr}>=1,INDEX(Purchasing!${MC[0]}${rec}:${MC[-1]}${rec},"
                          f"{c}$3-Components!$I${cr}),0)*Components!$D${cr}")
        wca[f"{c}{r}"].number_format = "$#,##0"
    wca.cell(row=r, column=LC + 1, value=f"=SUM({MC[0]}{r}:{MC[-1]}{r})").number_format = "$#,##0"
tot = 5 + len(COMPONENTS)
wca.cell(row=tot, column=2, value="Total payments").font = B
for c in MC + [L(LC + 1)]:
    wca[f"{c}{tot}"] = f"=SUM({c}5:{c}{tot-1})"
    wca[f"{c}{tot}"].number_format = "$#,##0"
    wca[f"{c}{tot}"].font = B
CASH_TOT = tot
wca.column_dimensions["B"].width = 18

# ---------------- Dashboard ----------------
wd = wb.create_sheet("Dashboard", 1)
title(wd, "Planning dashboard", "Everything on this page updates from the Inputs sheet")
month_index_row(wd, 3)
wd["A5"], wd["B5"] = "Demand scenario", "=Inputs!C4"
wd["B5"].number_format = "+0%;-0%;0%"
wd["A6"], wd["B6"] = "Units to produce (12 months)", f"=SUM(Production!{L(LC+1)}{PROD_ROWS[0]}:{L(LC+1)}{PROD_ROWS[1]})"
wd["B6"].number_format = "#,##0"
wd["A7"], wd["B7"] = "Supplier payments (12 months)", f"=Cash!{L(LC+1)}{CASH_TOT}"
wd["B7"].number_format = "$#,##0"
wd["A8"], wd["B8"] = "Purchase orders to place this month", "=" + "+".join(f"(Purchasing!{MC[0]}{MRP[c[0]]['po']}>0)" for c in COMPONENTS)
for a in ("A5", "A6", "A7", "A8", "A9"):
    wd[a].font = B
# shortage alerts: receipts required within the lead-time window (can't arrive in time if ordered today)
wd["A9"] = "Components short (need stock sooner than lead time)"
parts = []
for c in COMPONENTS:
    cr, rec = COMP_ROW[c[0]], MRP[c[0]]["rec"]
    parts.append(f"(SUMPRODUCT((Purchasing!${MC[0]}$3:${MC[-1]}$3<=Components!$E${cr})*(Purchasing!${MC[0]}${rec}:${MC[-1]}${rec}>0))>0)")
wd["B9"] = "=" + "+".join(parts)
wd.conditional_formatting.add("B9", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=ALERT), font=Font(bold=True, color="B42318")))
wd.column_dimensions["A"].width = 46
wd.column_dimensions["B"].width = 16

header(wd, 11, ["Month", "Units produced", "Supplier payments"])
for i, c in enumerate(MC):
    r = 12 + i
    wd.cell(row=r, column=1, value=MONTHS[i])
    wd.cell(row=r, column=2, value=f"=SUM(Production!{c}{PROD_ROWS[0]}:{c}{PROD_ROWS[1]})").number_format = "#,##0"
    wd.cell(row=r, column=3, value=f"=Cash!{c}{CASH_TOT}").number_format = "$#,##0"
wd.column_dimensions["C"].width = 18

ch = BarChart()
ch.title, ch.y_axis.title, ch.height, ch.width = "Supplier payments by month", "$", 7.5, 16
ch.add_data(Reference(wd, min_col=3, min_row=11, max_row=11 + N), titles_from_data=True)
ch.set_categories(Reference(wd, min_col=1, min_row=12, max_row=11 + N))
wd.add_chart(ch, "E5")
lc = LineChart()
lc.title, lc.y_axis.title, lc.height, lc.width = "Units produced by month", "units", 7.5, 16
lc.add_data(Reference(wd, min_col=2, min_row=11, max_row=11 + N), titles_from_data=True)
lc.set_categories(Reference(wd, min_col=1, min_row=12, max_row=11 + N))
wd.add_chart(lc, "E21")

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = sheet.title not in ("Start here", "Dashboard")
    sheet.freeze_panes = None
wb.save(OUT)
print("saved", OUT)
