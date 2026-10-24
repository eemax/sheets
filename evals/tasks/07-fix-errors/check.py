import openpyxl, re
from evalkit import Checks, args, close, stored, values

work, truth, answer = args()
path = f"{work}/finance.xlsx"
c = Checks()
rows, _ = values(path, "Finance")
errors = [(r, x) for r, v in rows.items() for x in v if isinstance(x, str) and x.startswith("#")]
c.add("no error values", not errors, errors[:5])
ws = openpyxl.load_workbook(path)["Finance"]
profit = [ws.cell(r, 4).value for r in range(2, 2 + truth["rows"])]
c.add("profit is a formula in every row", all(isinstance(p, str) and re.fullmatch(r"=B(\d+)-C\1", p.replace(" ", "")) for p in profit), profit)
tot = next((r for r, v in rows.items() if v and v[0] == "Total"), None)
if c.add("a total row", tot is not None):
    t = rows[tot]
    c.add("revenue total", close(t[1], truth["revenue"]), t[1])
    c.add("cost total", close(t[2], truth["cost"]), t[2])
    c.add("profit total", close(t[3], truth["profit"]), t[3])
    c.add("profit total stored", close(stored(path, "Finance", f"D{tot}"), truth["profit"]), stored(path, "Finance", f"D{tot}"), kind="stored")
c.report()
