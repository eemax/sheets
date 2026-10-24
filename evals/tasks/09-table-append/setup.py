# An Excel table (Insert > Table) with a totals row, and a formula outside
# it that reads the table by name.
import json, sys
import openpyxl
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

work, truth = sys.argv[1], sys.argv[2]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Region", "Units", "Price"])
for row in [("North", 5, 2.5), ("South", 7, 3.0), ("East", 2, 4.0)]:
    ws.append(row)
ws.append(["Total", "=SUBTOTAL(109,SalesT[Units])", None])
t = Table(displayName="SalesT", ref="A1:C5", totalsRowCount=1, autoFilter=AutoFilter(ref="A1:C4"))
t.tableColumns = [TableColumn(id=1, name="Region", totalsRowLabel="Total"),
                  TableColumn(id=2, name="Units", totalsRowFunction="sum"), TableColumn(id=3, name="Price")]
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(t)
ws["E2"] = "=SUM(SalesT[Units])"
wb.save(f"{work}/tables.xlsx")
json.dump({"units": 20}, open(truth, "w"))
