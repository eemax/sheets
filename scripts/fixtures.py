"""Generate test workbooks with openpyxl, a different producer than excelize.

They hold what real files carry: charts, validations, comments, links,
Excel tables, styling from another tool. Regenerate with:

    uv run --with openpyxl --with xlsxwriter python scripts/fixtures.py internal/xlsx/testdata
"""
import datetime
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

out = sys.argv[1]

# rich.xlsx: chart, validation, comment, hyperlink, Excel table.
wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Region", "Units"])
for row in [("North", 5), ("South", 7), ("East", 3)]:
    ws.append(row)
chart = BarChart()
chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
ws.add_chart(chart, "E2")
dv = DataValidation(type="list", formula1='"North,South,East,West"')
ws.add_data_validation(dv)
dv.add("A2:A4")
ws["B2"].comment = Comment("checked", "sheets")
ws["A3"].hyperlink = "https://example.com"
t = wb.create_sheet("T")
for row in [("k", "v"), ("a", 1), ("b", 2)]:
    t.append(row)
table = Table(displayName="Things", ref="A1:B3")
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
t.add_table(table)
wb.save(f"{out}/rich.xlsx")

# styled.xlsx: loud formatting from another tool, a merged title, a hidden
# sheet and a defined name.
wb = Workbook()
ws = wb.active
ws.title = "Budget 2026"
ws.sheet_properties.tabColor = "FF0000"
ws["A1"] = "Budget overview"
ws["A1"].font = Font(name="Comic Sans MS", size=20, color="FF00FF")
ws.merge_cells("A1:D1")
for i, h in enumerate(["Department", "Start", "Budget", "Spent", "Share"], 1):
    c = ws.cell(row=3, column=i, value=h)
    c.fill = PatternFill("solid", fgColor="FFFF00")
    c.font = Font(name="Arial", italic=True)
rows = [
    ("Engineering", datetime.date(2026, 1, 1), 1200000, 800000.5),
    ("Marketing", datetime.date(2026, 2, 1), 300000, 310000),
    ("Sales", datetime.date(2026, 3, 15), 450000, 120000.25),
    ("Support", datetime.date(2026, 1, 10), 90000, 45000),
]
for r, row in enumerate(rows, 4):
    for c, v in enumerate(row, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = Border(left=Side("thick", color="00FF00"))
        if c == 2:
            cell.number_format = "mm/dd/yyyy"
        if c in (3, 4):
            cell.number_format = '#,##0.00 "€"'
    ws.cell(row=r, column=5, value=f"=D{r}/C{r}").number_format = "0.00%"
ws["A8"] = "Total"
ws["C8"] = "=SUM(C4:C7)"
ws["D8"] = "=SUM(D4:D7)"
ws.freeze_panes = "C5"
ws.conditional_formatting.add("D4:D7", CellIsRule(operator="greaterThan", formula=["C4"], fill=PatternFill("solid", fgColor="FF9999")))
ws["A11"] = "Source: finance export"
rates = wb.create_sheet("Rates")
rates["A1"] = "EUR"
rates["B1"] = 1.08
rates.sheet_state = "hidden"
wb.save(f"{out}/styled.xlsx")

# tables.xlsx: an Excel table with a totals row, stored (as openpyxl does)
# under an absolute part path, with a note below it. The columns carry the
# totals settings Excel writes (a label, a function).
wb = Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Region", "Units", "Price"])
for row in [("North", 5, 2.5), ("South", 7, 3.0), ("East", 2, 4.0)]:
    ws.append(row)
ws.append(["Total", "=SUBTOTAL(109,SalesT[Units])", None])
# As Excel writes it, the filter leaves out the totals row.
table = Table(displayName="SalesT", ref="A1:C5", totalsRowCount=1, autoFilter=AutoFilter(ref="A1:C4"))
table.tableColumns = [
    TableColumn(id=1, name="Region", totalsRowLabel="Total"),
    TableColumn(id=2, name="Units", totalsRowFunction="sum"),
    TableColumn(id=3, name="Price"),
]
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(table)
ws["A8"] = "Source: CRM"
ws["E2"] = "=SUM(SalesT[Units])"
wb.save(f"{out}/tables.xlsx")

# cached.xlsx: formulas with their results stored, as Excel saves them (no
# recalculation flag); written with XlsxWriter, which can store results.
import xlsxwriter

xw = xlsxwriter.Workbook(f"{out}/cached.xlsx")
data = xw.add_worksheet("Data")
data.write_row(0, 0, ["Region", "Units"])
for i, (region, units) in enumerate([("North", 5), ("South", 7), ("East", 3)], 1):
    data.write_row(i, 0, [region, units])
data.write("A5", "Total")
data.write_formula("B5", "=SUM(B2:B4)", None, 15)
xw.define_name("Rate", "=Data!$D$1")
data.write("D1", 2)
summary = xw.add_worksheet("Summary")
summary.write_formula("A1", "=Data!B5*Rate", None, 30)  # through a total and a name
summary.write_formula("A2", "=_xlfn.LET(_xlpm.x,Data!B2,_xlpm.x*10)", None, 50)  # computed as Data!B2*10
summary.write_formula("A3", "=Data!B3+1", None, 8)  # untouched by an edit of B2
summary.write_formula("A4", "=ROWS(Data!B2:B4)", None, 3)
summary.write_formula("A5", "=_xlfn.XMATCH(5,Data!B2:B4)", None, 1)  # a function the calculator lacks
xw.close()

# Excel drops the recalculation flag XlsxWriter sets when it saves.
import zipfile

path = f"{out}/cached.xlsx"
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename == "xl/workbook.xml":
            data = data.replace(b' fullCalcOnLoad="1"', b"")
        z.writestr(info, data)
