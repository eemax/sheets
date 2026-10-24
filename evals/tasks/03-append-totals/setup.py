# A small table with a total row and a chart of its revenue.
import json, sys
import openpyxl
from openpyxl.chart import BarChart, Reference

work, truth = sys.argv[1], sys.argv[2]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Region", "Units", "Price", "Revenue"])
rows = [("North", 120, 12.5), ("South", 80, 30), ("East", 200, 12.5), ("North", 95, 30)]
for i, (r, u, p) in enumerate(rows, 2):
    ws.append([r, u, p, f"=B{i}*C{i}"])
ws.append(["Total", "=SUM(B2:B5)", None, "=SUM(D2:D5)"])
chart = BarChart()
chart.title = "Revenue"
chart.add_data(Reference(ws, min_col=4, min_row=1, max_row=5), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
ws.add_chart(chart, "F2")
wb.save(f"{work}/sales.xlsx")
units = sum(u for _, u, _ in rows) + 15 + 4
revenue = sum(u * p for _, u, p in rows) + 15 * 120 + 4 * 30
json.dump({"units": units, "revenue": revenue, "orders": 6}, open(truth, "w"))
