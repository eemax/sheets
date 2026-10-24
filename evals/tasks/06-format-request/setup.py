import json, sys
import openpyxl

work, truth = sys.argv[1], sys.argv[2]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Budget"
ws.append(["Department", "Budget", "Spent", "Remaining"])
for i, (d, b, s) in enumerate([("Engineering", 1200000, 800000.5), ("Marketing", 300000, 310000), ("Sales", 450000, 120000.25)], 2):
    ws.append([d, b, s, f"=B{i}-C{i}"])
wb.save(f"{work}/budget.xlsx")
json.dump({"rows": 4}, open(truth, "w"))
