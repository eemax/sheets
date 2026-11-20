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

# kinds_xlsxwriter.xlsx: a formula of each kind, for move, copy and sort (as
# the wave-1 audit's k1.xlsx): plain ones, one reading a range where one
# value goes (F1, implicit intersection), legacy arrays in one cell (C) and
# over two (H1:H2), dynamic arrays in one cell (D, as Excel 365 stores SUM
# over an array) and over two (G1:G2).
xw = xlsxwriter.Workbook(f"{out}/kinds_xlsxwriter.xlsx")
s = xw.add_worksheet("S")
for i, (a, b) in enumerate([(3, 1), (1, 2), (2, 3), (5, 4), (4, 5), (6, 6)]):
    r = i + 1
    s.write_row(i, 0, [a, b])
    if r % 2:
        s.write_array_formula(f"C{r}:C{r}", f"{{=SUM(A{r}:B{r}*2)}}")
    else:
        s.write_formula(f"C{r}", f"=A{r}+B{r}")
    if r % 3 == 0:
        s.write_dynamic_array_formula(f"D{r}", f"=SUM(A{r}:B{r}*3)")
    else:
        s.write_formula(f"D{r}", f"=A{r}*B{r}")
s.write_formula("F1", "=B1:B2*100")
s.write_dynamic_array_formula("G1:G2", "=SEQUENCE(2)")
s.write_array_formula("H1:H2", "{=A1:A2*10}")
xw.close()

# kinds_pending.xlsx: dynamic functions as XlsxWriter's write_formula stores
# them, cm="1" with the range of their own cell and no result yet (Excel
# spills =SEQUENCE(3) in D1 over D1:D3 when it opens the file).
xw = xlsxwriter.Workbook(f"{out}/kinds_pending.xlsx")
s = xw.add_worksheet("S")
for i, (a, b) in enumerate([(3, 1), (1, 2), (2, 3), (5, 4), (4, 5), (6, 6)]):
    s.write_row(i, 0, [a, b])
s.write_formula("D1", "=SEQUENCE(3)")
xw.close()

# kinds_unknown.xlsx: the same with dynamic functions sheets cannot compute
# (MAP over a LAMBDA), so how far they spill is unknown. S has one in D1
# (Excel spills it over D1:D3) with text in D4; U has one in each of D1:D6.
xw = xlsxwriter.Workbook(f"{out}/kinds_unknown.xlsx")
s = xw.add_worksheet("S")
u = xw.add_worksheet("U")
for i, (a, b) in enumerate([(9, 1), (5, 2), (4, 3), (1, 4), (3, 5), (2, 6)]):
    s.write_row(i, 0, [a, b])
    u.write_row(i, 0, [a, b])
s.write_formula("D1", "=_xlfn.MAP(B1:B3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2))")
s.write("D4", "x")
for r in range(1, 7):
    u.write_formula(f"D{r}", f"=_xlfn.MAP(B{r},_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2))")
xw.close()

# unlike_excel.xlsx: one formula of each class WP-01 found computed unlike
# Excel (sheet Cases, A1:D12: the class, the formula, what Excel shows, and
# a formula reading it), over sheet Data.
wb = Workbook()
d = wb.active
d.title = "Data"
for r, (day, n, k) in enumerate([(1, 1, "A"), (2, 2, "B"), (3, 4, "A")], 1):
    d.cell(r, 1, datetime.date(2020, 1, day))
    d.cell(r, 2, n)
    d.cell(r, 3, k)
d["E1"], d["F1"], d["G1"] = "#N/A", "found", "#N/A"
d["H5"], d["H6"] = "A", "B"
c = wb.create_sheet("Cases")
cases = [
    ("TEXT, month names", '=TEXT(Data!A1,"mmmm")', "unknown"),
    ("TEXT, day names", '=TEXT(Data!A1,"dddd")', "unknown"),
    ("TEXT, AM/PM", '=TEXT(0.75,"h:mm AM/PM")', "unknown"),
    ("TEXT, a locale tag", '=TEXT(Data!A1,"[$-41D]mmmm")', "unknown"),
    ("TEXT, the en-US tag", '=TEXT(Data!A1,"[$-409]mmmm")', "January"),
    ("TEXT, Swedish codes", '=TEXT(Data!A1,"ÅÅÅÅ-MM-DD")', "unknown"),
    ("SUMPRODUCT over whole columns", ArrayFormula("B7", "=SUMPRODUCT(Data!A:A=Data!A1)*Data!B:C"), "unknown"),
    ("ERROR.TYPE inside IF", ArrayFormula("B8", '=_xlfn.IFNA(INDEX(Data!E1:G1,MATCH(3,ERROR.TYPE(COLUMN(IF(Data!E1:G1<>"~N/A",Data!E1:G1))),0)),"")'), "found"),
    ("SUMIFS with DATE, criteria range", '=SUMIFS(Data!B1:B3,Data!A1:A3,">="&DATE(2020,1,2),Data!C1:C3,Data!H5:H6)', 0),
    ("SUMIFS with DATE", '=SUMIFS(Data!B1:B3,Data!A1:A3,">="&DATE(2020,1,2))', 6),
    ("text date in a criterion", '=SUMIF(Data!A1:A3,"<1/2/2020",Data!B1:B3)', "unknown"),
    ("ISO date in a criterion", '=SUMIF(Data!A1:A3,"<2020-01-02",Data!B1:B3)', 1),
]
for r, (label, f, excel) in enumerate(cases, 1):
    c.cell(r, 1, label)
    c.cell(r, 2).value = f
    c.cell(r, 3, excel)
    c.cell(r, 4, f"=B{r}")
wb.save(f"{out}/unlike_excel.xlsx")

# against_before.xlsx: a workbook as Excel saves it, results stored, for
# sheets check --against. Data!B2:B6 are inputs Summary's formulas read
# (through a name, TaxRate, too); a column chart grouped with a text box
# (a shape), as Excel saves a group; validations, one of them a list read
# from another sheet, which Excel keeps in its 2010 extension list; and a
# conditional format.
xw = xlsxwriter.Workbook(f"{out}/against_before.xlsx")
data = xw.add_worksheet("Data")
summary = xw.add_worksheet("Summary")
data.write_row(0, 0, ["Region", "Units", "Price", "Revenue"])
rows = [("North", 5, 30), ("South", 7, 20), ("East", 3, 50), ("West", 9, 10), ("Mid", 4, 25)]
for i, (region, units, price) in enumerate(rows, 1):
    data.write_row(i, 0, [region, units, price])
    data.write_formula(i, 3, f"=B{i + 1}*C{i + 1}", None, units * price)
units, revenue = sum(r[1] for r in rows), sum(r[1] * r[2] for r in rows)
data.write("A7", "Total")
data.write_formula("B7", "=SUM(B2:B6)", None, units)
data.write_formula("D7", "=SUM(D2:D6)", None, revenue)
red = xw.add_format({"bg_color": "#FFC7CE"})
data.conditional_format("D2:D6", {"type": "cell", "criteria": ">", "value": 140, "format": red})
data.data_validation("B2:B6", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 1000})
data.data_validation("C2:C6", {"validate": "decimal", "criteria": ">=", "value": 0})
chart = xw.add_chart({"type": "column"})
chart.add_series({"name": "Revenue", "categories": "=Data!$A$2:$A$6", "values": "=Data!$D$2:$D$6"})
data.insert_chart("F2", chart)
data.insert_textbox("F18", "Revenue by region")
xw.define_name("TaxRate", "=Summary!$B$1")
summary.write_row(0, 0, ["Tax rate", 0.25])
tax = revenue * 0.25
for r, (label, f, v) in enumerate([
    ("Units", "=SUM(Data!B2:B6)", units),
    ("Revenue", "=SUM(Data!D2:D6)", revenue),
    ("Tax", "=B3*TaxRate", tax),
    ("Net", "=B3-B4", revenue - tax),
    ("Most units", "=MAX(Data!B2:B6)", max(r[1] for r in rows)),
    ("North units", "=Data!B2", rows[0][1]),
], 2):
    summary.write(f"A{r}", label)
    summary.write_formula(f"B{r}", f, None, v)
summary.write("D1", "Regions")
summary.write_column("D2", [r[0] for r in rows])
xw.close()

# As Excel saves them: no recalculation flag; the list validation read
# from Summary in the extension list; the chart and the text box grouped.
import re

x14_list = (
    b'<extLst><ext uri="{CCE6A557-97BC-4b89-ADB5-D3C9D4A6E9D0}" '
    b'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    b'<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    b'<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">'
    b"<x14:formula1><xm:f>Summary!$D$2:$D$6</xm:f></x14:formula1><xm:sqref>A2:A6</xm:sqref>"
    b"</x14:dataValidation></x14:dataValidations></ext></extLst>"
)


def group(drawing):
    """Groups a drawing's two anchors (a chart, then a shape) into one."""
    anchors = re.findall(rb"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>", drawing, re.S)
    assert len(anchors) == 2, len(anchors)
    chart, shape = anchors
    frame = re.search(rb"<xdr:graphicFrame\b.*?</xdr:graphicFrame>", chart, re.S).group(0)
    sp = re.search(rb"<xdr:sp\b.*?</xdr:sp>", shape, re.S).group(0)
    # The chart's place in EMUs (column F, row 2; 480 by 288 pixels), as
    # a group's members carry it; the text box has its own already.
    x, y, cx, cy = 5 * 609600, 190500, 480 * 9525, 288 * 9525
    frame = frame.replace(b'<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>', b'<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>' % (x, y, cx, cy))
    off = re.search(rb'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', sp)
    sx, sy, scx, scy = map(int, off.groups())
    gx, gy = min(x, sx), min(y, sy)
    gcx, gcy = max(x + cx, sx + scx) - gx, max(y + cy, sy + scy) - gy
    xfrm = b'<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/><a:chOff x="%d" y="%d"/><a:chExt cx="%d" cy="%d"/>' % (gx, gy, gcx, gcy, gx, gy, gcx, gcy)
    start = re.search(rb"<xdr:from>.*?</xdr:from>", chart, re.S).group(0)
    end = re.search(rb"<xdr:to>.*?</xdr:to>", shape, re.S).group(0)
    grouped = (
        b"<xdr:twoCellAnchor>" + start + end + b"<xdr:grpSp><xdr:nvGrpSpPr>"
        b'<xdr:cNvPr id="4" name="Group 3"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
        b"<xdr:grpSpPr><a:xfrm>" + xfrm + b"</a:xfrm></xdr:grpSpPr>" + frame + sp + b"</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>"
    )
    return drawing.replace(chart, b"").replace(shape, grouped)


path = f"{out}/against_before.xlsx"
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename == "xl/workbook.xml":
            data = data.replace(b' fullCalcOnLoad="1"', b"")
        elif info.filename == "xl/worksheets/sheet1.xml":
            data = data.replace(b"</worksheet>", x14_list + b"</worksheet>")
        elif info.filename == "xl/drawings/drawing1.xml":
            data = group(data)
        z.writestr(info, data)

# against_openpyxl.xlsx: the same workbook loaded by openpyxl, Data!B2
# changed, and saved. openpyxl stores no results, keeps no shape, no chart
# in a group and no extension list.
import openpyxl

wb = openpyxl.load_workbook(path)
wb["Data"]["B2"] = 6
wb.save(f"{out}/against_openpyxl.xlsx")

# against_chartsheet.xlsx: a chart on a chart sheet (a sheet that is one
# chart, which sheets does not read as a sheet) beside a worksheet;
# against_chartsheet_deleted.xlsx: openpyxl loads it, deletes the chart
# sheet and saves, which loses the chart without a trace in the cells.
cs = xlsxwriter.Workbook(f"{out}/against_chartsheet.xlsx")
cs_data = cs.add_worksheet("Data")
cs_data.write_column("A1", [1, 2, 3])
cs_chart = cs.add_chart({"type": "bar"})
cs_chart.add_series({"values": "=Data!$A$1:$A$3"})
cs.add_chartsheet("Chart1").set_chart(cs_chart)
cs.close()

wb = openpyxl.load_workbook(f"{out}/against_chartsheet.xlsx")
del wb["Chart1"]
wb.save(f"{out}/against_chartsheet_deleted.xlsx")

# against_1904.xlsx: a workbook counting dates from 1904 (a Mac's), with a
# date typed in and a formula reading it, as XlsxWriter stores them: the
# date 2024-01-31 is the serial 43860 here, 45322 in a workbook counting from
# 1900. check --against against a file that counts from 1900 compares dates.
d1904 = xlsxwriter.Workbook(f"{out}/against_1904.xlsx", {"date_1904": True})
d1904_data = d1904.add_worksheet("Data")
d1904_data.write_datetime("A1", datetime.datetime(2024, 1, 31), d1904.add_format({"num_format": "yyyy-mm-dd"}))
d1904_data.write("B1", 5)
d1904_data.write_formula("C1", "=YEAR(A1)", None, 2024)
d1904.close()

# against_nodata.xlsx: openpyxl deletes the sheet Data, which Summary's
# formulas read, and leaves them as they were (Excel would write #REF!);
# against_noname.xlsx: openpyxl deletes the name TaxRate, which Summary!B4
# uses (Excel leaves it, now #NAME?), and Data!B2 changed; against_facts.xlsx:
# openpyxl renames Data to Facts and leaves the formulas reading Data.
wb = openpyxl.load_workbook(path)
del wb["Data"]
wb.save(f"{out}/against_nodata.xlsx")
wb = openpyxl.load_workbook(path)
del wb.defined_names["TaxRate"]
wb["Data"]["B2"] = 6
wb.save(f"{out}/against_noname.xlsx")
wb = openpyxl.load_workbook(path)
wb["Data"].title = "Facts"
wb.save(f"{out}/against_facts.xlsx")

# against_cut.xlsx: openpyxl cuts the regions' rows Data!A2:D6 and pastes
# them at A12 (move_range, its own formulas translated), and leaves the
# formulas reading them as they were (Excel rewrites them to follow).
wb = openpyxl.load_workbook(path)
wb["Data"].move_range("A2:D6", rows=10, translate=True)
wb.save(f"{out}/against_cut.xlsx")

# against_rtd.xlsx and against_rtd2.xlsx: a rate a server sends (RTD), as
# Excel stored it at two times, 1.08 and then 1.09, and a formula reading
# it. Nothing in the workbook changed.
for name, rate in (("against_rtd", 1.08), ("against_rtd2", 1.09)):
    rtd = xlsxwriter.Workbook(f"{out}/{name}.xlsx")
    rtd_data = rtd.add_worksheet("Data")
    rtd_data.write("A1", "EURUSD")
    rtd_data.write_formula("B1", '=RTD("my.server",,"EURUSD")', None, rate)
    rtd_data.write("A2", "Price")
    rtd_data.write("B2", 100)
    rtd_data.write_formula("C2", "=B2*B1", None, round(100 * rate, 2))
    rtd.close()
    # Excel's results: not marked to be computed again on open.
    with zipfile.ZipFile(f"{out}/{name}.xlsx") as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    with zipfile.ZipFile(f"{out}/{name}.xlsx", "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                data = data.replace(b' fullCalcOnLoad="1"', b"")
            z.writestr(info, data)

# against_legacy.xlsx: openpyxl writes a formula reading a range where one
# value goes as older Excel stores it (implicit intersection: Model!C2 reads
# the row of Data!C2:C3 it is in), with no results stored. A row inserted
# above the range moves it off that row: the value moves with the rows
# added, though the formula reads none of them.
wb = openpyxl.Workbook()
legacy = wb.active
legacy.title = "Data"
for row in (["Item", "Units", "Price"], ["a", 1, 5], ["b", 2, 7], ["c", 3, 9]):
    legacy.append(row)
model = wb.create_sheet("Model")
model["A2"] = "x"
model["C2"] = "=Data!C2:C3*2"
model["D2"] = "=C2+1"
wb.save(f"{out}/against_legacy.xlsx")

# against_addchart.xlsx: against_before.xlsx loaded by openpyxl, a chart of
# other data (Units) added at H20 and another validation (a list typed in,
# E2:E6) with it, and saved. The round trip drops the grouped chart, the text
# box and the list read from Summary (in the 2010 extension list) all the
# same, though a chart and a list validation are now on the sheet.
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.load_workbook(f"{out}/against_before.xlsx")
ws = wb["Data"]
units = BarChart()
units.type = "col"
units.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
ws.add_chart(units, "H20")
dv = DataValidation(type="list", formula1='"North,South,East,West"')
ws.add_data_validation(dv)
dv.add("E2:E6")
wb.save(f"{out}/against_addchart.xlsx")

# against_table.xlsx: an Excel table (Sales, Data!A1:B3) that two formulas
# read by its name. against_table_renamed.xlsx: the table renamed Sales2,
# with the formulas as Excel rewrites them for it.
from openpyxl.worksheet.table import Table

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Data"
for row in (("Region", "Units"), ("North", 5), ("South", 7)):
    ws.append(row)
ws.add_table(Table(displayName="Sales", ref="A1:B3"))
total = wb.create_sheet("Sum")
total["A1"] = "=SUM(Sales[Units])"
total["A2"] = "=Sales[[#This Row],[Units]]"
wb.save(f"{out}/against_table.xlsx")
table = ws.tables["Sales"]
table.name = table.displayName = "Sales2"
total["A1"] = "=SUM(Sales2[Units])"
total["A2"] = "=Sales2[[#This Row],[Units]]"
wb.save(f"{out}/against_table_renamed.xlsx")

# against_block.xlsx: a table (Data!A1:D4) with a block of cells beside it
# (H1:J4) that formulas read, Sum's among them. against_blockcut.xlsx: openpyxl
# cuts the block and pastes it below the table, in other columns
# (move_range, Data!H1:J4 to A30), and leaves the formulas reading the cells
# as they were, which Excel would have rewritten.
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Data"
for row in (("Item", "Units", "Price", "Total"), ("a", 3, 5, "=B2*C2"), ("b", 4, 6, "=B3*C3"), ("c", 5, 7, "=B4*C4")):
    ws.append(row)
for r, row in enumerate((("Key", "Qty", "Cost"), ("x", 10, "=I2*B2"), ("y", 20, "=I3*B3"), ("z", 30, "=SUM(J2:J3)+I4")), 1):
    for c, v in enumerate(row, 8):
        ws.cell(r, c, v)
total = wb.create_sheet("Sum")
total["A1"] = "=SUM(Data!I2:I4)"
total["A2"] = "=Data!D2+Data!J2"
wb.save(f"{out}/against_block.xlsx")
ws.move_range("H1:J4", rows=29, cols=-7)
wb.save(f"{out}/against_blockcut.xlsx")
