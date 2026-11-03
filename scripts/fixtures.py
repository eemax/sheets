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
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

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

# calccol.xlsx: an Excel table with a calculated column (Total, which Excel
# fills into new rows) and a totals row, with a formula beside it reading a
# column of the table.
wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Rep", "Q1", "Q2", "Q3", "Q4", "Total"])
for name, *q in [("Ann", 1, 2, 3, 4), ("Bob", 5, 6, 7, 8)]:
    ws.append([name, *q, "=SUM(Sales[[#This Row],[Q1]:[Q4]])"])
ws.append(["Total", None, None, None, None, "=SUBTOTAL(109,Sales[Total])"])
table = Table(displayName="Sales", ref="A1:F4", totalsRowCount=1)
columns = [TableColumn(id=i, name=n) for i, n in enumerate(["Rep", "Q1", "Q2", "Q3", "Q4", "Total"], 1)]
columns[0].totalsRowLabel = "Total"
columns[5].calculatedColumnFormula = TableFormula(attr_text="SUM(Sales[[#This Row],[Q1]:[Q4]])")
columns[5].totalsRowFunction = "sum"
table.tableColumns = columns
ws.add_table(table)
ws["H1"] = "Grand"
ws["H2"] = "=SUM(Sales[Total])"
wb.save(f"{out}/calccol.xlsx")

# unprefixed.xlsx: functions added after Excel 2007 written as typed,
# which openpyxl stores without the _xlfn. prefix Excel needs: Excel shows
# #NAME? for them.
wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Code", "Qty", "First PK", "Joined", "Max PK", "Total"])
for code, qty in [("PK1", 10), ("XX2", 20), ("PK3", 30)]:
    ws.append([code, qty])
ws["C2"] = '=IFERROR(INDEX($A$2:$A$4,AGGREGATE(15,6,(ROW($A$2:$A$4)-1)/(LEFT($A$2:$A$4,2)="PK"),1)),"")'
ws["D2"] = '=TEXTJOIN(",",TRUE,A2:A4)'
ws["E2"] = '=MAXIFS(B2:B4,A2:A4,"PK*")'
ws["F2"] = "=SUM(B2:B4)"
wb.save(f"{out}/unprefixed.xlsx")

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

# rules.xlsx: conditional formats and validations as Excel writes them,
# with the Excel 2010 extension (data bars, written by XlsxWriter as Excel
# does), rules whose formulas read their own row or another sheet, and
# merged cells.
xw = xlsxwriter.Workbook(f"{out}/rules.xlsx")
data = xw.add_worksheet("Data")
data.merge_range("A1:D1", "Orders by region", None)
data.write_row(1, 0, ["Region", "Units", "Price", "Flag"])
for i, row in enumerate([("East", 5, 30, 1), ("West", 2, 40, 0), ("North", 9, 10, 1), ("South", 1, 100, 0), ("East", 7, 20, 1)], 2):
    data.write_row(i, 0, row)
red = xw.add_format({"bg_color": "#FFC7CE"})
data.conditional_format("B3:B7", {"type": "data_bar", "bar_color": "#63BE7B", "data_bar_2010": True})
data.conditional_format("C3:C7", {"type": "icon_set", "icon_style": "3_traffic_lights"})
data.conditional_format("A3:D7", {"type": "formula", "criteria": "=$D3=1", "format": red})
data.conditional_format("C3:C7", {"type": "cell", "criteria": ">", "value": "$C$3", "format": red})
data.data_validation("B3:B7", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 100})
data.data_validation("C3:C7", {"validate": "custom", "value": "=C3>$B3"})
other = xw.add_worksheet("Other")
other.write_column(0, 0, [1, 2, 3])
other.conditional_format("A1:A3", {"type": "formula", "criteria": "=Data!$B$4>1", "format": red})
other.data_validation("B1:B3", {"validate": "list", "source": "=Data!$A$3:$A$7"})
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

# threads.xlsx: an Excel 365 comment thread (with its stand-in note for
# older Excel) on Data!B3, a plain note on A1, and the sheet's rows.
wb = Workbook()
ws = wb.active
ws.title = "Data"
for row in [("Name", "Score"), ("Ada", 92), ("Bob", 47), ("Cy", 75)]:
    ws.append(row)
ws["A1"].comment = Comment("Names as entered", "Ann")
thread = "{AAAAAAAA-0000-4000-8000-000000000001}"
ws["B3"].comment = Comment(
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment.\n\n"
    "Comment:\n    Is 47 right?\nReply:\n    Yes, checked.",
    "tc=" + thread,
)
path = f"{out}/threads.xlsx"
wb.save(path)

ns = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
ann, bob = "{11111111-1111-4111-8111-111111111111}", "{22222222-2222-4222-8222-222222222222}"
threads = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<ThreadedComments xmlns="{ns}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    f'<threadedComment ref="B3" dT="2026-10-01T10:00:00.00" personId="{ann}" id="{thread}"><text>Is 47 right?</text></threadedComment>'
    f'<threadedComment ref="B3" dT="2026-10-01T11:00:00.00" personId="{bob}" id="{{AAAAAAAA-0000-4000-8000-000000000002}}" parentId="{thread}"><text>Yes, checked.</text></threadedComment>'
    f"</ThreadedComments>"
)
persons = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<personList xmlns="{ns}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    f'<person displayName="Ann" id="{ann}" userId="ann@example.com" providerId="None"/>'
    f'<person displayName="Bob" id="{bob}" userId="bob@example.com" providerId="None"/>'
    f"</personList>"
)
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename == "[Content_Types].xml":
            data = data.replace(
                b"</Types>",
                b'<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>'
                b'<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
            )
        elif info.filename == "xl/_rels/workbook.xml.rels":
            data = data.replace(
                b"</Relationships>",
                b'<Relationship Id="rIdPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
            )
        elif info.filename == "xl/worksheets/_rels/sheet1.xml.rels":
            data = data.replace(
                b"</Relationships>",
                b'<Relationship Id="rIdThreads" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
            )
        z.writestr(info, data)
    z.writestr("xl/threadedComments/threadedComment1.xml", threads)
    z.writestr("xl/persons/person.xml", persons)

# stale.xlsx: formulas as XlsxWriter writes them: a 0 stored for each, which
# the recalculation flag marks as out of date; a function added after Excel
# 2007 without its prefix; a function the sheets calculator lacks
# (BAHTTEXT) and a formula reading it; an error; an array formula.
xw = xlsxwriter.Workbook(f"{out}/stale.xlsx")
data = xw.add_worksheet("Data")
data.write_row(0, 0, ["Region", "Units"])
for i, (region, units) in enumerate([("North", 5), ("South", 7), ("North", 3)], 1):
    data.write_row(i, 0, [region, units])
data.write_formula("C2", "=SUM(B2:B4)")
data.write_formula("C3", '=MAXIFS(B2:B4,A2:A4,"North")')
data.write_formula("C4", "=BAHTTEXT(B2)")
data.write_formula("C5", '=C4&"!"')
data.write_formula("C6", "=1/0")
data.write_array_formula("C7:C7", "{=SUM(B2:B4*B2:B4)}")
data.write_formula("C8", "=C2*2")
xw.close()

# vouched.xlsx: results as Excel saves them (no recalculation flag), one of
# which the sheets calculator does not agree with (999 stored for =B2*2),
# and one formula stored without its result.
xw = xlsxwriter.Workbook(f"{out}/vouched.xlsx")
data = xw.add_worksheet("Data")
data.write_row(0, 0, ["Region", "Units"])
for i, (region, units) in enumerate([("North", 5), ("South", 7), ("East", 3)], 1):
    data.write_row(i, 0, [region, units])
data.write_formula("C2", "=SUM(B2:B4)", None, 15)
data.write_formula("C3", "=B2*2", None, 999)
data.write_formula("C4", "=B4*3", None, 0)
xw.close()
path = f"{out}/vouched.xlsx"
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename == "xl/workbook.xml":
            data = data.replace(b' fullCalcOnLoad="1"', b"")
        elif info.filename == "xl/worksheets/sheet1.xml":
            data = data.replace(b'<c r="C4"><f>B4*3</f><v>0</v></c>', b'<c r="C4"><f>B4*3</f></c>')
        z.writestr(info, data)

# arrays_openpyxl.xlsx: array formulas entered with Ctrl+Shift+Enter over
# several cells, as openpyxl writes them (the first cell alone).
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.worksheet.formula import ArrayFormula

wb = Workbook()
ws = wb.active
ws.title = "S"
for r, (a, b) in enumerate([(1, 10), (2, 20), (3, 30), (4, 40)], 1):
    ws.cell(r, 1, a)
    ws.cell(r, 2, b)
ws["D1"] = ArrayFormula("D1:D4", "=A1:A4*B1:B4")
ws["F1"] = ArrayFormula("F1:G1", "=TRANSPOSE(A1:A2)")
ws["H1"] = ArrayFormula("H1:H3", "=SUM(A1:A4)")
ws["J1"] = ArrayFormula("J1:J3", '=IF(A1:A3>1,"big","")')  # its first value is ""
ws["K1"] = ArrayFormula("K1:K3", '=IFS(A1:A3>1,"a",TRUE,"b")')
ws["L1"] = ArrayFormula("L1:L3", '=SWITCH(A1:A3,1,"x","y")')
wb.save(f"{out}/arrays_openpyxl.xlsx")

# arrays_xlsxwriter.xlsx: the same from XlsxWriter (a 0 in every cell), and
# one calling a function without its prefix.
xw = xlsxwriter.Workbook(f"{out}/arrays_xlsxwriter.xlsx")
s = xw.add_worksheet("S")
for i, (a, b) in enumerate([(1, 10), (2, 20), (3, 30)], 1):
    s.write_row(i, 0, [a, b])
s.write_array_formula("E2:E4", "{=A2:A4*B2:B4}")
s.write_array_formula("G2:G2", "{=STDEV.S(A2:A4*2)}")
xw.close()

# dates1904.xlsx: a workbook counting dates from 1904.
wb = Workbook()
wb.epoch = CALENDAR_MAC_1904
ws = wb.active
ws.title = "S"
ws["A1"] = datetime.date(2020, 1, 1)
ws["B1"] = "=YEAR(A1)"
ws["C1"] = "=DATE(2020,1,1)"
ws["D1"] = "=A1+1"
ws["E1"] = '=TEXT(A1,"yyyy-mm-dd")'
wb.save(f"{out}/dates1904.xlsx")

# refs.xlsx: ROWS over a range holding the formula itself (not circular),
# volatile functions, and INDIRECT.
wb = Workbook()
ws = wb.active
ws.title = "S"
ws["A2"] = "=ROWS($A$2:A2)"
ws["A3"] = "=ROWS($A$2:A3)"
ws["B2"] = "=COLUMNS(A2:B2)"
ws["C1"] = datetime.date(2026, 1, 1)
ws["C2"] = "=INT(TODAY()-C1+(1))"
ws["C3"] = "=NOW()-NOW()"
ws["C4"] = '=INDIRECT("A2")'
wb.save(f"{out}/refs.xlsx")

# dynamic_xlsxwriter.xlsx: dynamic arrays as XlsxWriter writes them (a 0
# for each, its range the first cell alone).
xw = xlsxwriter.Workbook(f"{out}/dynamic_xlsxwriter.xlsx")
d = xw.add_worksheet("Data")
for i, r in enumerate([("Region", "Units"), ("North", 5), ("South", 7), ("North", 3), ("East", 1)]):
    d.write_row(i, 0, r)
d.write_dynamic_array_formula("D1", "=UNIQUE(A2:A5)")
d.write_formula("E1", "=SUM(B2:B5)")
d.write_dynamic_array_formula("F1", "=SORT(B2:B5,1,-1)")
xw.close()

# arrays_inside.xlsx: a formula of its own inside an array formula's range,
# which Excel does not allow.
wb = Workbook()
ws = wb.active
ws.title = "S"
for r in range(1, 4):
    ws.cell(r, 1, r)
ws["E1"] = ArrayFormula("E1:E3", "=A1:A3*3")
ws["E3"] = "=A1+100"
wb.save(f"{out}/arrays_inside.xlsx")

# arrays_rows.xlsx: arrays and plain formulas next to them, for edits of rows
# and columns: a plain formula over a range (implicit intersection), a
# legacy array over two cells below a plain formula and a value, and
# arrays in single cells, one per row.
wb = Workbook()
ws = wb.active
ws.title = "Sheet"
for r in range(1, 5):
    ws.cell(r, 2, r)
ws["A1"] = "=B1:B2*100"
ws["C1"] = "=B1+100"
ws["C2"] = "keep me"
ws["C3"] = ArrayFormula("C3:C4", "=B1:B2*3")
for r in range(1, 4):
    ws.cell(r, 4).value = ArrayFormula(f"D{r}", f"=SUM(B{r}:B{r}*2)")
wb.save(f"{out}/arrays_rows.xlsx")
