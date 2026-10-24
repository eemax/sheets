# 150,000 order lines with a revenue formula, a styled header, frozen panes
# and a filter, and a Summary of SUMIFS by month; store S-114's March orders
# appended a second time at the end, as a repeated import leaves them.
# Results stored, as Excel saves them.
import datetime, json, random, sys
import xlsxwriter
from evalkit import excel_saved

work, truth = sys.argv[1], sys.argv[2]
rng = random.Random(17)
n = 150_000
stores = [f"S-{101 + i}" for i in range(40)]
products = [(f"P-{200 + i}", round(rng.uniform(4, 180), 2)) for i in range(25)]
start = datetime.date(2026, 1, 1)
orders = []
for i in range(n):
    d = start + datetime.timedelta(days=i * 365 // n)
    p, price = rng.choice(products)
    orders.append((f"SO-{i + 1:06d}", d, d.strftime("%Y-%m"), rng.choice(stores), p, rng.randint(1, 40), price))
dupes = [o for o in orders if o[3] == "S-114" and o[2] == "2026-03"]
rows = orders + dupes
last = len(rows) + 1
path = f"{work}/orders_2026.xlsx"
wb = xlsxwriter.Workbook(path, {"constant_memory": True})
head = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF"})
date = wb.add_format({"num_format": "yyyy-mm-dd"})
money = wb.add_format({"num_format": "#,##0.00"})
d = wb.add_worksheet("Data")
d.write_row(0, 0, ["Order ID", "Date", "Month", "Store", "Product", "Units", "Price", "Revenue"], head)
for r, (oid, day, month, store, prod, units, price) in enumerate(rows, 1):
    d.write_string(r, 0, oid)
    d.write_datetime(r, 1, datetime.datetime.combine(day, datetime.time()), date)
    d.write_string(r, 2, month)
    d.write_string(r, 3, store)
    d.write_string(r, 4, prod)
    d.write_number(r, 5, units)
    d.write_number(r, 6, price, money)
    d.write_formula(r, 7, f"=F{r + 1}*G{r + 1}", money, round(units * price, 2))
d.freeze_panes(1, 0)
d.autofilter(0, 0, last - 1, 7)
d.set_column("A:A", 11)
d.set_column("B:B", 11)
d.set_column("H:H", 12)
s = wb.add_worksheet("Summary")
s.write_row(1, 0, ["Month", "Revenue", "Orders"], head)
months = [f"2026-{m:02d}" for m in range(1, 13)]


def month_rev(rs, m):
    return round(sum(u * p for _, _, mo, _, _, u, p in rs if mo == m), 2)


for i, m in enumerate(months):
    s.write_string(2 + i, 0, m)
    s.write_formula(2 + i, 1, f"=SUMIFS(Data!$H$2:$H${last},Data!$C$2:$C${last},A{3 + i})", money, month_rev(rows, m))
    s.write_formula(2 + i, 2, f"=COUNTIFS(Data!$C$2:$C${last},A{3 + i})", None, sum(1 for o in rows if o[2] == m))
s.write(14, 0, "Total", head)
s.write_formula(14, 1, "=SUM(B3:B14)", money, round(sum(u * p for *_, u, p in rows), 2))
s.write_formula(14, 2, "=SUM(C3:C14)", None, len(rows))
s.set_column("A:C", 14)
wb.close()
excel_saved(path)
json.dump({"rows": n, "dupes": len(dupes), "units": sum(o[5] for o in orders), "march": month_rev(orders, "2026-03"),
           "total": round(sum(u * p for *_, u, p in orders), 2)}, open(truth, "w"))
