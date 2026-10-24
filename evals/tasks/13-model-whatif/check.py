import re
from evalkit import Checks, args, close, formula, mentions, stored, values

work, truth, answer = args()
path = f"{work}/plan.xlsx"
want = truth["downside"]
c = Checks()
inp, _ = values(path, "Inputs!B4:B5")
c.add("growth is 2%", close(inp[4][0], 0.02), inp[4][0])
c.add("COGS is 52%", close(inp[5][0], 0.52), inp[5][0])
c.add("the P&L is still formulas", all(formula(path, "PnL", cell) for cell in ("B3", "F3", "F10", "F11")),
      [formula(path, "PnL", cell) for cell in ("F3", "F11")])
pl, _ = values(path, "PnL!B11:F11")
ni = pl[11]
c.add("P&L net income by year", all(close(a, b) for a, b in zip(ni, truth["ni"])), f"{ni} want {[round(x) for x in truth['ni']]}")
s, _ = values(path, "Summary!B2:B3")
c.add("Summary 2030 net income", close(s[2][0], want["ni_2030"]), f"{s[2][0]} want {want['ni_2030']:.0f}")
c.add("Summary cumulative net income", close(s[3][0], want["cumulative"]), f"{s[3][0]} want {want['cumulative']:.0f}")
# The NPV is judged from its formula and the P&L, not by a calculator.
f = (formula(path, "Summary", "B4") or "").replace(" ", "").upper()
c.add("Summary NPV is still NPV over the net income", re.fullmatch(r"=NPV\((INPUTS!)?\$?B\$?10,PNL!\$?B\$?11:\$?F\$?11\)", f) is not None, f)
npv = stored(path, "Summary", "B4")
c.add("no wrong NPV stored in the file", npv is None or close(npv, want["npv"]), f"stored {npv} want {want['npv']:.0f}")
for key, name in (("ni_2030", "2030 net income"), ("cumulative", "cumulative net income"), ("npv", "NPV")):
    c.add(f"answer gives the {name}", mentions(answer, want[key]), f"want {want[key]:.0f}")
for row, key, name in ((2, "ni_2030", "2030 net income"), (3, "cumulative", "cumulative net income"), (4, "npv", "NPV")):
    v = stored(path, "Summary", f"B{row}")
    c.add(f"Summary {name} stored", close(v, want[key]), v, kind="stored")
c.report()
