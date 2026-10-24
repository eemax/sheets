# A five-year plan as Excel saves it (results stored): inputs, a P&L of
# chained formulas across sheets (bracketed tax by VLOOKUP, depreciation by
# MIN), and a Summary with SUM, NPV and a CAGR.
import json, sys
import xlsxwriter
from evalkit import excel_saved

work, truth = sys.argv[1], sys.argv[2]
years = list(range(2026, 2031))
base = {"rev": 2_400_000, "growth": 0.06, "cogs": 0.45, "opex": 700_000, "opex_g": 0.03, "capex": 120_000, "dep_years": 5, "rate": 0.09}
brackets = [(0, 0.15), (250_000, 0.21), (750_000, 0.25)]


def plan(p):
    rows = {k: [] for k in ("rev", "cogs", "gp", "opex", "ebitda", "dep", "ebit", "tax", "ni")}
    for y in years:
        rev = p["rev"] * (1 + p["growth"]) ** (y - 2026)
        cogs = -rev * p["cogs"]
        opex = -p["opex"] * (1 + p["opex_g"]) ** (y - 2026)
        dep = -min(y - 2025, p["dep_years"]) * p["capex"] / p["dep_years"]
        ebit = rev + cogs + opex + dep
        rate = [r for lo, r in brackets if ebit >= lo][-1] if ebit > 0 else 0
        tax = -(ebit * rate if ebit > 0 else 0)
        for k, v in zip(rows, (rev, cogs, rev + cogs, opex, rev + cogs + opex, dep, ebit, tax, ebit + tax)):
            rows[k].append(v)
    ni = rows["ni"]
    summary = {"ni_2030": ni[-1], "cumulative": sum(ni), "npv": sum(v / (1 + p["rate"]) ** (i + 1) for i, v in enumerate(ni)),
               "cagr": (rows["rev"][-1] / rows["rev"][0]) ** 0.25 - 1}
    return rows, summary


rows, summary = plan(base)
path = f"{work}/plan.xlsx"
wb = xlsxwriter.Workbook(path)
bold = wb.add_format({"bold": True})
money = wb.add_format({"num_format": "#,##0;(#,##0)"})
pct = wb.add_format({"num_format": "0.0%"})
inp = wb.add_worksheet("Inputs")
inp.write("A1", "Plan inputs", bold)
labels = [("Base revenue 2026", "rev", money), ("Revenue growth", "growth", pct), ("COGS (% of revenue)", "cogs", pct),
          ("Base opex 2026", "opex", money), ("Opex growth", "opex_g", pct), ("Capex per year", "capex", money),
          ("Depreciation years", "dep_years", None), ("Discount rate", "rate", pct)]
for i, (label, k, f) in enumerate(labels, 2):
    inp.write(i, 0, label)
    inp.write(i, 1, base[k], f)
inp.write_row("D2", ["Taxable income from", "Tax rate"], bold)
for i, (lo, r) in enumerate(brackets, 2):
    inp.write(i, 3, lo, money)
    inp.write(i, 4, r, pct)
inp.set_column("A:A", 22)
inp.set_column("D:D", 20)

pl = wb.add_worksheet("PnL")
pl.write("A2", "Year", bold)
for j, y in enumerate(years):
    pl.write(1, 1 + j, y, bold)
lines = [("Revenue", "rev", "=Inputs!$B$3*(1+Inputs!$B$4)^({c}$2-2026)"),
         ("COGS", "cogs", "=-{c}3*Inputs!$B$5"),
         ("Gross profit", "gp", "={c}3+{c}4"),
         ("Opex", "opex", "=-Inputs!$B$6*(1+Inputs!$B$7)^({c}$2-2026)"),
         ("EBITDA", "ebitda", "={c}5+{c}6"),
         ("Depreciation", "dep", "=-MIN({c}$2-2025,Inputs!$B$9)*Inputs!$B$8/Inputs!$B$9"),
         ("EBIT", "ebit", "={c}7+{c}8"),
         ("Tax", "tax", "=-IF({c}9>0,{c}9*VLOOKUP({c}9,Inputs!$D$3:$E$5,2,TRUE),0)"),
         ("Net income", "ni", "={c}9+{c}10")]
for i, (label, k, f) in enumerate(lines, 2):
    pl.write(i, 0, label, bold if k in ("gp", "ebitda", "ebit", "ni") else None)
    for j in range(len(years)):
        col = "BCDEF"[j]
        pl.write_formula(i, 1 + j, f.format(c=col), money, rows[k][j])
pl.set_column("A:A", 16)
pl.set_column("B:F", 12)

s = wb.add_worksheet("Summary")
s.write("A1", "Summary", bold)
for i, (label, f, k, fmt) in enumerate([("2030 net income", "=PnL!F11", "ni_2030", money),
                                         ("Cumulative net income 2026-2030", "=SUM(PnL!B11:F11)", "cumulative", money),
                                         ("NPV of net income", "=NPV(Inputs!B10,PnL!B11:F11)", "npv", money),
                                         ("Revenue CAGR", "=(PnL!F3/PnL!B3)^(1/4)-1", "cagr", pct)], 1):
    s.write(i, 0, label)
    s.write_formula(i, 1, f, fmt, summary[k])
s.set_column("A:A", 32)
s.set_column("B:B", 14)
wb.close()
excel_saved(path)

down_rows, downside = plan(dict(base, growth=0.02, cogs=0.52))
json.dump({"base": summary, "downside": downside, "ni": down_rows["ni"], "rate": base["rate"]}, open(truth, "w"))
