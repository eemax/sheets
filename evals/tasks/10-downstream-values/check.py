from evalkit import Checks, args, close, formula, stored, values

work, truth, answer = args()
path = f"{work}/prices.xlsx"
c = Checks()
p, _ = values(path, "Products")
c.add("Widget costs 15", any(v and v[0] == "Widget" and close(v[1], 15) for v in p.values()), p)
c.add("orders still look prices up", formula(path, "Orders", "C2") is not None, formula(path, "Orders", "C2"))
o, _ = values(path, "Orders!C6")
c.add("total computes", close(o[6][0], truth["total"]), o[6][0])
c.add("total stored, as pandas reads it", close(stored(path, "Orders", "C6"), truth["total"]), stored(path, "Orders", "C6"), kind="stored")
c.report()
