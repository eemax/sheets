# Profit = Revenue - Cost, Margin = Profit / Revenue. Planted mistakes: a
# typed-in profit, a margin dividing by zero, totals that leave out rows.
import json, sys
import openpyxl

work, truth = sys.argv[1], sys.argv[2]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Finance"
ws.append(["Item", "Revenue", "Cost", "Profit", "Margin"])
data = [("Alpha", 1000, 600), ("Beta", 800, 500), ("Gamma", 0, 50), ("Delta", 1500, 900), ("Epsilon", 700, 580), ("Zeta", 400, 100), ("Eta", 950, 300)]
for i, (n, rev, cost) in enumerate(data, 2):
    ws.append([n, rev, cost, f"=B{i}-C{i}", f"=D{i}/B{i}"])
ws["D6"] = 120  # typed over the formula (B6-C6 is 120 today, but will not follow)
ws.append(["Total", "=SUM(B2:B6)", "=SUM(C2:C6)", "=SUM(D2:D6)", None])  # misses rows 7-8
wb.save(f"{work}/finance.xlsx")
rev = sum(r for _, r, _ in data)
cost = sum(c for _, _, c in data)
json.dump({"revenue": rev, "cost": cost, "profit": rev - cost, "rows": len(data)}, open(truth, "w"))
