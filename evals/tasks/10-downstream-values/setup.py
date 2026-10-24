# A workbook as Excel saves it: results stored, no recalculation flag.
import json, sys, zipfile
import xlsxwriter

work, truth = sys.argv[1], sys.argv[2]
path = f"{work}/prices.xlsx"
wb = xlsxwriter.Workbook(path)
p = wb.add_worksheet("Products")
prices = {"Widget": 12.5, "Gadget": 30.0, "Gizmo": 7.25}
p.write_row(0, 0, ["Item", "Price"])
for i, (k, v) in enumerate(prices.items(), 1):
    p.write_row(i, 0, [k, v])
o = wb.add_worksheet("Orders")
o.write_row(0, 0, ["Item", "Units", "Revenue"])
orders = [("Widget", 10), ("Gadget", 3), ("Widget", 4), ("Gizmo", 8)]
for i, (k, u) in enumerate(orders, 1):
    o.write_row(i, 0, [k, u])
    o.write_formula(i, 2, f"=B{i + 1}*VLOOKUP(A{i + 1},Products!$A$2:$B$4,2,FALSE)", None, u * prices[k])
o.write(5, 0, "Total")
o.write_formula(5, 2, "=SUM(C2:C5)", None, sum(u * prices[k] for k, u in orders))
wb.close()
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        if info.filename == "xl/workbook.xml":
            data = data.replace(b' fullCalcOnLoad="1"', b"")
        z.writestr(info, data)
prices["Widget"] = 15
json.dump({"total": sum(u * prices[k] for k, u in orders)}, open(truth, "w"))
