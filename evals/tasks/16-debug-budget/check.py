from evalkit import Checks, args, close, values

work, truth, answer = args()
path = f"{work}/budget.xlsx"
c = Checks()
names, first, total_row = truth["names"], truth["first"], truth["total_row"]
rows, _ = values(path, f"Budget!A{first}:N{total_row}")
lines = {v[0]: v for v in rows.values() if v and v[0]}
c.add("every line is still there", all(n in lines for n in names), [n for n in names if n not in lines])
inputs_ok = all(n in lines and all(close(lines[n][1 + j], truth["grid"][i][j]) for j in range(12)) for i, n in enumerate(names))
t = truth.get("text_line", "Travel")  # runs prepared before the line was chosen
c.add(f"the input figures are right ({t}'s June as a number)", inputs_ok)
c.add(f"{t}'s June counts as a number", t in lines and isinstance(lines[t][6], (int, float)), lines.get(t, [None] * 7)[6])
c.add("every annual total is the sum of its 12 months", all(n in lines and close(lines[n][13], truth["annual"][i]) for i, n in enumerate(names)),
      [(n, lines[n][13], truth["annual"][i]) for i, n in enumerate(names) if n in lines and not close(lines[n][13], truth["annual"][i])][:3])
tot = lines.get("Total")
c.add("every month total takes every line", tot is not None and all(close(tot[1 + j], truth["col_totals"][j]) for j in range(12)),
      tot and [(j, tot[1 + j], truth["col_totals"][j]) for j in range(12) if not close(tot[1 + j], truth["col_totals"][j])])
c.add("the annual grand total", tot is not None and close(tot[13], truth["total"]), tot and tot[13])
s, _ = values(path, "Summary!A1:B20")
summ = {v[0]: v[1] for v in s.values() if v and v[0]}
for label, want in truth["picks"].items():
    c.add(f"Summary {label}", close(summ.get(label), want), f"{summ.get(label)} want {want}")
c.add("Summary total operating budget", close(summ.get("Total operating budget"), truth["total"]), f"{summ.get('Total operating budget')} want {truth['total']}")
c.report()
