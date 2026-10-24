import re, zipfile
from evalkit import Checks, args, close, stored, values

work, truth, answer = args()
path = f"{work}/sales.xlsx"
c = Checks()
rows, _ = values(path, "Sales")
data = {r: v for r, v in rows.items() if v and v[0] not in (None, "Region", "Total")}
total = [r for r, v in rows.items() if v and v[0] == "Total"]
c.add("six orders", len(data) == truth["orders"], f"{len(data)} order rows")
new = [v for v in data.values() if (v[0], v[1], v[2]) in (("West", 15, 120), ("East", 4, 30))]
c.add("both new orders", len(new) == 2, f"found {len(new)}")
c.add("their revenue computes", all(close(v[3], v[1] * v[2]) for v in new), [v[3] for v in new])
c.add("one total row, below the orders", len(total) == 1 and total[0] > max(data, default=0), f"total rows {total}")
if total:
    t = rows[total[0]]
    c.add("units total", close(t[1], truth["units"]), f"got {t[1]}, want {truth['units']}")
    c.add("revenue total", close(t[3], truth["revenue"]), f"got {t[3]}, want {truth['revenue']}")
    c.add("revenue total stored", close(stored(path, "Sales", f"D{total[0]}"), truth["revenue"]),
          stored(path, "Sales", f"D{total[0]}"), kind="stored")
with zipfile.ZipFile(path) as z:
    charts = [n for n in z.namelist() if n.startswith("xl/charts/chart")]
    xml = z.read(charts[0]).decode() if charts else ""
refs = re.findall(r"<(?:c:)?f>([^<]*\$D\$[^<]*)</(?:c:)?f>", xml)
def covers():
    for ref in refs:
        m = re.search(r"\$D\$(\d+)(?::\$D\$(\d+))?", ref)
        lo, hi = int(m.group(1)), int(m.group(2) or m.group(1))
        if lo <= min(data) and hi >= max(data) and (not total or hi < total[0]):
            return True, ref
    return False, refs
c.add("a chart", bool(charts))
c.guard("the chart shows every order and not the total", covers)
c.report()
