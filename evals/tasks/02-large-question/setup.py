# 30,000 orders; Revenue is a formula without stored results.
import datetime, json, random, sys
import openpyxl

work, truth = sys.argv[1], sys.argv[2]
random.seed(2)
wb = openpyxl.Workbook(write_only=True)
ws = wb.create_sheet("Orders")
ws.append(["Order", "Date", "Region", "Status", "Units", "Price", "Revenue"])
count, total = 0, 0.0
d0 = datetime.date(2026, 1, 1)
for i in range(2, 30002):
    status = random.choice(["open", "shipped", "shipped", "cancelled"])
    units, price = random.randint(1, 20), round(random.uniform(5, 200), 2)
    ws.append([f"ORD-{i:06d}", d0 + datetime.timedelta(days=i % 300), random.choice(["North", "South", "East", "West"]),
               status, units, price, f"=E{i}*F{i}"])
    if status == "open":
        count, total = count + 1, total + units * price
wb.save(f"{work}/orders.xlsx")
json.dump({"count": count, "revenue": round(total, 2)}, open(truth, "w"))
