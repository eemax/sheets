# A report a person made: title, blank row, header, rows with a formula
# column (no stored results, as openpyxl writes them), a total row, a note.
import json, random, sys
import openpyxl

work, truth = sys.argv[1], sys.argv[2]
random.seed(1)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Q3"
ws["A1"] = "Q3 sales report"
ws.append([])
ws.append(["Region", "Product", "Units", "Price", "Revenue"])
totals = {}
for r in range(4, 34):
    region = random.choice(["North", "South", "East", "West"])
    units, price = random.randint(1, 40), random.choice([12.5, 30, 7.25, 120])
    ws.append([region, random.choice(["Widget", "Gadget", "Gizmo"]), units, price, f"=C{r}*D{r}"])
    totals[region] = totals.get(region, 0) + units * price
ws.append(["Total", None, "=SUM(C4:C33)", None, "=SUM(E4:E33)"])
ws.append([])
ws.append(["Source: CRM export, revenue before discounts"])
wb.save(f"{work}/report.xlsx")
json.dump({"revenue": totals}, open(truth, "w"))
