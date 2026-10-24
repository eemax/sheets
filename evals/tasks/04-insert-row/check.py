from evalkit import Checks, args, close, stored, values

work, truth, answer = args()
path = f"{work}/model.xlsx"
c = Checks()
rows, _ = values(path, "Data")
regions = [v[0] for r, v in sorted(rows.items()) if v and v[0] not in (None, "Region", "Total")]
c.add("regions in order", regions == truth["order"], regions)
units = {v[0]: v[1] for v in rows.values() if v}
c.add("Central has 4 units", units.get("Central") == 4, units.get("Central"))
c.add("total", close(units.get("Total"), truth["total"]), units.get("Total"))
s, _ = values(path, "Summary!A1:B3")
c.add("summary: South units", close(s[1][1], truth["south"]), s[1][1])
c.add("summary: all units", close(s[2][1], truth["total"]), s[2][1])
c.add("summary: South share", close(s[3][1], truth["south"] / truth["total"]), s[3][1])
c.add("summary stored", close(stored(path, "Summary", "B2"), truth["total"]), stored(path, "Summary", "B2"), kind="stored")
c.report()
