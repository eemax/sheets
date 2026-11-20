# A sales workbook of the kind a team keeps for years, made with XlsxWriter:
# two Excel tables with formula columns and structured references, named
# input cells, a cross-sheet summary, a chart, conditional formats, data
# validations and dynamic-array formulas. Synthetic data. The inputs go to
# the truth file for check.py, which computes the expected end state itself.
import datetime as dt
import json
import sys
import xlsxwriter

work, truth = sys.argv[1], sys.argv[2]
path = f"{work}/sales-book.xlsx"

INPUTS = {"tax": 0.20, "bulk_units": 10, "bulk_disc": 0.05, "target_margin": 0.35}
REGIONS = ["North", "South", "East", "West"]
CATEGORIES = ["Hardware", "Software", "Service"]
# product, category, price, cost
PRODUCTS = [("Widget", "Hardware", 12.50, 7.00), ("Gadget", "Hardware", 30.00, 18.50), ("Gizmo", "Hardware", 7.25, 3.10),
            ("Suite", "Software", 99.00, 22.00), ("Toolkit", "Software", 45.00, 12.00), ("Support", "Service", 60.00, 35.00)]
# date, region, product, units
ORDERS = [("2026-01-05", "North", "Widget", 6), ("2026-01-09", "South", "Gadget", 4), ("2026-01-14", "East", "Suite", 3),
          ("2026-01-19", "West", "Support", 5), ("2026-01-27", "West", "Gizmo", 20), ("2026-02-02", "North", "Toolkit", 8),
          ("2026-02-06", "East", "Widget", 12), ("2026-02-11", "South", "Suite", 2), ("2026-02-16", "North", "Gadget", 9),
          ("2026-02-24", "West", "Toolkit", 4), ("2026-03-02", "South", "Support", 7), ("2026-03-06", "East", "Gizmo", 15),
          ("2026-03-10", "West", "Widget", 10), ("2026-03-13", "North", "Suite", 5)]

# use_future_functions writes SORTBY, UNIQUE and the like with their prefixes.
wb = xlsxwriter.Workbook(path, {"use_future_functions": True})
bold = wb.add_format({"bold": True})
title = wb.add_format({"bold": True, "font_size": 14})
money = wb.add_format({"num_format": "#,##0.00"})
pct = wb.add_format({"num_format": "0%"})
date = wb.add_format({"num_format": "yyyy-mm-dd"})
good = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
bad = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
top = wb.add_format({"bold": True, "font_color": "#1F4E79"})

# --- Inputs
ws = wb.add_worksheet("Inputs")
ws.write_row("A1", ["Assumption", "Value"], bold)
ws.write("A2", "Tax rate"); ws.write("B2", INPUTS["tax"], pct)
ws.write("A3", "Bulk threshold (units)"); ws.write("B3", INPUTS["bulk_units"])
ws.write("A4", "Bulk discount"); ws.write("B4", INPUTS["bulk_disc"], pct)
ws.write("A5", "Target margin"); ws.write("B5", INPUTS["target_margin"], pct)
ws.write("D1", "Regions", bold)
for i, r in enumerate(REGIONS):
    ws.write(1 + i, 3, r)
ws.write("G1", "Categories", bold)
for i, c in enumerate(CATEGORIES):
    ws.write(1 + i, 6, c)
ws.set_column("A:A", 24)
ws.data_validation("B2", {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 1,
                          "input_title": "Tax rate", "input_message": "A rate between 0 and 1",
                          "error_title": "Not a rate", "error_message": "Enter a rate between 0% and 100%."})
for name, ref in {"TaxRate": "B2", "BulkUnits": "B3", "BulkDisc": "B4", "TargetMargin": "B5"}.items():
    wb.define_name(name, f"=Inputs!${ref[0]}${ref[1:]}")
wb.define_name("RegionList", "=Inputs!$D$2:$D$5")
wb.define_name("CategoryList", "=Inputs!$G$2:$G$4")

# --- Products (an Excel table)
pr = wb.add_worksheet("Products")
n = len(PRODUCTS)
pr.add_table(0, 0, n, 4, {
    "name": "Products", "style": "Table Style Medium 2",
    "data": [[p, c, price, cost] for p, c, price, cost in PRODUCTS],
    "columns": [{"header": "Product"}, {"header": "Category"}, {"header": "Price", "format": money}, {"header": "Cost", "format": money},
                {"header": "Margin", "format": pct,
                 "formula": "=(Products[[#This Row],[Price]]-Products[[#This Row],[Cost]])/Products[[#This Row],[Price]]"}]})
pr.data_validation("B2:B50", {"validate": "list", "source": "=CategoryList"})
pr.conditional_format("E2:E7", {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#63BE7B"})
pr.freeze_panes(1, 0)
pr.set_column("A:B", 14)

# --- Orders (an Excel table)
od = wb.add_worksheet("Orders")
look = "MATCH(Orders[[#This Row],[Product]],Products[Product],0)"
rows = []
for d, r, p, u in ORDERS:
    rows.append([dt.datetime.fromisoformat(d), r, p, u])
od.add_table(0, 0, len(ORDERS), 7, {
    "name": "Orders", "style": "Table Style Medium 9",
    "data": rows,
    "columns": [{"header": "Date", "format": date}, {"header": "Region"}, {"header": "Product"}, {"header": "Units"},
                {"header": "Category", "formula": f"=INDEX(Products[Category],{look})"},
                {"header": "Price", "format": money, "formula": f"=INDEX(Products[Price],{look})"},
                {"header": "Revenue", "format": money,
                 "formula": "=Orders[[#This Row],[Units]]*Orders[[#This Row],[Price]]"
                            "*(1-IF(Orders[[#This Row],[Units]]>=BulkUnits,BulkDisc,0))"},
                {"header": "Cost", "format": money, "formula": f"=Orders[[#This Row],[Units]]*INDEX(Products[Cost],{look})"}]})
last = len(ORDERS) + 1
od.data_validation("B2:B100", {"validate": "list", "source": "=RegionList"})
od.data_validation("D2:D100", {"validate": "integer", "criteria": ">=", "value": 1,
                               "error_title": "Units", "error_message": "Whole units, at least 1."})
od.conditional_format(f"D2:D{last}", {"type": "cell", "criteria": ">=", "value": "BulkUnits", "format": good})
od.conditional_format(f"G2:G{last}", {"type": "top", "value": 3, "format": top})
od.freeze_panes(1, 0)
od.set_column("A:A", 12)
od.set_column("B:H", 11)

# --- Summary
sm = wb.add_worksheet("Summary")
sm.write("A1", "Sales summary", title)
sm.write_row("A3", ["Region", "Revenue", "Cost", "Margin", "Share", "Margin %"], bold)
for i, r in enumerate(REGIONS):
    row = 4 + i
    sm.write(f"A{row}", r)
    sm.write_formula(f"B{row}", f"=SUMIFS(Orders[Revenue],Orders[Region],A{row})", money)
    sm.write_formula(f"C{row}", f"=SUMIFS(Orders[Cost],Orders[Region],A{row})", money)
    sm.write_formula(f"D{row}", f"=B{row}-C{row}", money)
    sm.write_formula(f"E{row}", f"=IF(B$8=0,0,B{row}/B$8)", pct)
    sm.write_formula(f"F{row}", f"=IF(B{row}=0,0,D{row}/B{row})", pct)
sm.write("A8", "Total", bold)
for col in "BCD":
    sm.write_formula(f"{col}8", f"=SUM({col}4:{col}7)", money)
sm.write_formula("E8", "=SUM(E4:E7)", pct)
sm.write_formula("F8", "=IF(B8=0,0,D8/B8)", pct)
sm.write("A10", "Tax"); sm.write_formula("B10", "=B8*TaxRate", money)
sm.write("A11", "Revenue after tax"); sm.write_formula("B11", "=B8-B10", money)
sm.write_row("A13", ["Category", "Revenue"], bold)
for i, c in enumerate(CATEGORIES):
    sm.write(f"A{14 + i}", c)
    sm.write_formula(f"B{14 + i}", f"=SUMIFS(Orders[Revenue],Orders[Category],A{14 + i})", money)
sm.write("A17", "Total", bold); sm.write_formula("B17", "=SUM(B14:B16)", money)
sm.write("A19", "Products below target margin")
sm.write_formula("B19", "=SUMPRODUCT(--(Products[Margin]<TargetMargin))")
sm.write("A20", "Top region"); sm.write_formula("B20", "=Insights!A2")
sm.set_column("A:A", 30); sm.set_column("B:H", 12)
sm.conditional_format("F4:F7", {"type": "cell", "criteria": "<", "value": "TargetMargin", "format": bad})
sm.conditional_format("E4:E7", {"type": "data_bar", "bar_color": "#5B9BD5"})
ch = wb.add_chart({"type": "column"})
ch.add_series({"name": "Revenue", "categories": "=Summary!$A$4:$A$7", "values": "=Summary!$B$4:$B$7"})
ch.set_title({"name": "Revenue by region"})
ch.set_legend({"none": True})
sm.insert_chart("J3", ch)

# --- Insights: dynamic arrays (spills of fixed size)
ins = wb.add_worksheet("Insights")
ins.write("A1", "Regions by revenue", bold)
ins.write_dynamic_array_formula("A2:A5", "=SORTBY(Summary!$A$4:$A$7,Summary!$B$4:$B$7,-1)")
ins.write("B1", "Revenue", bold)
ins.write_dynamic_array_formula("B2:B5", "=SORT(Summary!$B$4:$B$7,1,-1)", money)
ins.write("D1", "Three largest orders", bold)
ins.write_dynamic_array_formula("D2:D4", "=LARGE(Orders[Revenue],{1;2;3})", money)
ins.write("F1", "Categories sold", bold)
ins.write_dynamic_array_formula("F2:F4", "=UNIQUE(Orders[Category])")
ins.set_column("A:F", 20)
wb.close()

json.dump({"inputs": INPUTS, "products": PRODUCTS, "orders": ORDERS, "regions": REGIONS, "categories": CATEGORIES}, open(truth, "w"))
