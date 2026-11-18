# A department budget the way a finance team keeps one, made with XlsxWriter
# and saved as Excel saves (results stored): a chart, data validations of
# four kinds, conditional formats of five kinds, a note, frozen panes, and a
# rounded rectangle (a shape, not a text box: it is written into the drawing
# the way Excel writes a shape). Synthetic data. The inputs go to the truth
# file for check.py, which computes the expected end state itself.
import datetime as dt
import json
import re
import sys
import zipfile
import xlsxwriter
from evalkit import excel_saved

work, truth = sys.argv[1], sys.argv[2]
path = f"{work}/budget.xlsx"
# department, owner, Q1..Q4, status, review date
DEPTS = [("Sales", "Anna", 120000, 125000, 130000, 140000, "On track", "2026-06-30"),
         ("Marketing", "Ben", 80000, 95000, 70000, 60000, "At risk", "2026-06-30"),
         ("Engineering", "Chen", 310000, 310000, 320000, 330000, "On track", "2026-07-31"),
         ("Support", "Dara", 90000, 90000, 95000, 95000, "On track", "2026-07-31"),
         ("Finance", "Elin", 60000, 60000, 60000, 62000, "On track", "2026-08-31"),
         ("HR", "Femi", 45000, 45000, 48000, 48000, "Over", "2026-08-31")]
NOTE = "Hiring freeze until Q3."
SHAPE = "Draft: figures in USD, not yet approved"

wb = xlsxwriter.Workbook(path)
ws = wb.add_worksheet("Budget")
money = wb.add_format({"num_format": "#,##0"})
pct = wb.add_format({"num_format": "0.0%"})
date = wb.add_format({"num_format": "yyyy-mm-dd"})
head = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF"})
bold = wb.add_format({"bold": True, "num_format": "#,##0", "top": 1})
boldpct = wb.add_format({"bold": True, "num_format": "0.0%", "top": 1})
boldtext = wb.add_format({"bold": True, "top": 1})
ws.write("A1", "Department budget 2026", wb.add_format({"bold": True, "font_size": 14}))
ws.write_row("A3", ["Department", "Owner", "Q1", "Q2", "Q3", "Q4", "Total", "Share", "Status", "Review date"], head)
grand = sum(sum(d[2:6]) for d in DEPTS)
for i, (dept, owner, *q, status, review) in enumerate(DEPTS):
    r = 4 + i
    ws.write(f"A{r}", dept)
    ws.write(f"B{r}", owner)
    for col, v in zip("CDEF", q):
        ws.write_number(f"{col}{r}", v, money)
    ws.write_formula(f"G{r}", f"=SUM(C{r}:F{r})", money, sum(q))
    ws.write_formula(f"H{r}", f"=G{r}/G$10", pct, sum(q) / grand)
    ws.write(f"I{r}", status)
    ws.write_datetime(f"J{r}", dt.datetime.fromisoformat(review), date)
ws.write("A10", "Total", boldtext)
ws.write("B10", "", boldtext)
for col, k in zip("CDEF", range(2, 6)):
    ws.write_formula(f"{col}10", f"=SUM({col}4:{col}9)", bold, sum(d[k] for d in DEPTS))
ws.write_formula("G10", "=SUM(G4:G9)", bold, grand)
ws.write_formula("H10", "=SUM(H4:H9)", boldpct, 1.0)
for col in "IJ":
    ws.write(f"{col}10", "", boldtext)
ws.write_comment("A9", NOTE)  # HR's row
ws.freeze_panes(3, 1)
ws.set_column("A:A", 16)
ws.set_column("B:B", 10)
ws.set_column("C:H", 11)
ws.set_column("I:J", 13)
# conditional formats of five kinds
ws.conditional_format("C4:F9", {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.conditional_format("G4:G9", {"type": "data_bar", "bar_color": "#5B9BD5"})
ws.conditional_format("H4:H9", {"type": "cell", "criteria": ">", "value": 0.25,
                                "format": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})})
ws.conditional_format("I4:I9", {"type": "cell", "criteria": "==", "value": '"Over"', "format": wb.add_format({"bold": True, "font_color": "#C00000"})})
ws.conditional_format("J4:J9", {"type": "formula", "criteria": "=$J4<DATE(2026,7,15)", "format": wb.add_format({"bg_color": "#FFE699"})})
# data validations of four kinds
ws.data_validation("I4:I20", {"validate": "list", "source": ["On track", "At risk", "Over"],
                              "input_title": "Status", "input_message": "Pick one: On track, At risk or Over."})
ws.data_validation("C4:F20", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 10000000,
                              "error_title": "Budget", "error_message": "A whole amount between 0 and 10,000,000."})
ws.data_validation("J4:J20", {"validate": "date", "criteria": ">=", "value": dt.date(2026, 1, 1)})
ws.data_validation("B4:B20", {"validate": "custom", "value": "=LEN(B4)>=2", "error_message": "An owner needs at least two letters."})
ch = wb.add_chart({"type": "column"})
ch.add_series({"name": "Total", "categories": "=Budget!$A$4:$A$9", "values": "=Budget!$G$4:$G$9"})
ch.set_title({"name": "Budget by department"})
ch.set_legend({"none": True})
ws.insert_chart("A13", ch)
ws.insert_textbox("L3", SHAPE, {"width": 200, "height": 60, "fill": {"color": "#FFF2CC"}, "line": {"color": "#BF9000"}})
wb.close()
excel_saved(path)

# XlsxWriter writes a text box; Excel writes a shape without txBox="1", with the theme's shape style.
STYLE = (b'<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
         b'<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
         b'<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>')
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename.startswith("xl/drawings/drawing"):
            data = (data.replace(b' txBox="1"', b"").replace(b'prst="rect"', b'prst="roundRect"')
                    .replace(b'name="TextBox 2"', b'name="Rounded Rectangle 2"'))
            data = re.sub(rb"<xdr:style>.*?</xdr:style>", STYLE, data, flags=re.S)
        z.writestr(info, data)
json.dump({"depts": DEPTS, "note": NOTE, "shape": SHAPE}, open(truth, "w"))
