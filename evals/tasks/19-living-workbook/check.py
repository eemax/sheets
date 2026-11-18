# Grades the 20 edits on the end state. Formula values come only from
# evalkit.lo_values (a LibreOffice recalculation of a copy), and the
# expected values are computed here, in Python, from the inputs setup.py
# saved: nothing here asks sheets, or the tool that did the edits.
# Features are read from the package's XML or with openpyxl.
import datetime as dt
import re
import warnings
import zipfile
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.utils import range_boundaries
from evalkit import Checks, args, lo_values, mentions, stored

warnings.filterwarnings("ignore")  # openpyxl on extensions it does not read

work, truth, answer = args()
path = f"{work}/sales-book.xlsx"
TARGETS = [("North", 1500), ("South", 800), ("East", 650), ("West", 900)]


def model(t):
    """The workbook after the 20 edits, computed from the inputs."""
    I = dict(t["inputs"])
    products = [dict(name=n, cat=c, price=p, cost=k) for n, c, p, k in t["products"]]
    orders = [dict(date=d, region=r, product=p, units=u) for d, r, p, u in t["orders"]]
    pick = lambda name: next(p for p in products if p["name"] == name)
    by_date = lambda d: next(o for o in orders if o["date"] == d)
    I["tax"] = 0.22                                                                     # 1
    pick("Gadget")["price"] = 34.0                                                      # 2
    products.append(dict(name="Gizmo Pro", cat="Hardware", price=48.0, cost=21.0))      # 3
    orders.append(dict(date="2026-03-18", region="North", product="Gizmo Pro", units=12))  # 4
    by_date("2026-01-14")["units"] = 4                                                  # 5
    by_date("2026-01-19")["region"] = "South"                                           # 6
    I["bulk_units"] = 8                                                                 # 7
    pick("Toolkit")["cat"] = "Service"                                                  # 11
    orders.remove(by_date("2026-01-27"))                                                # 12
    I["target_margin"] = 0.55                                                           # 15
    products.sort(key=lambda p: -p["price"])                                            # 17
    I["ship"] = 4.5                                                                     # 18
    for p in products:
        p["margin"] = (p["price"] - p["cost"]) / p["price"]
    for o in orders:
        p = pick(o["product"])
        bulk = o["units"] >= I["bulk_units"]
        o.update(cat=p["cat"], price=p["price"], rev=o["units"] * p["price"] * (1 - (I["bulk_disc"] if bulk else 0)),
                 cost=o["units"] * p["cost"], ship=0.0 if bulk else I["ship"])
    reg = {}
    for r in t["regions"]:
        mine = [o for o in orders if o["region"] == r]
        rev, cost = sum(o["rev"] for o in mine), sum(o["cost"] for o in mine)
        target = dict(TARGETS)[r]
        reg[r] = dict(rev=rev, cost=cost, margin=rev - cost, n=len(mine), gap=rev - target)
    total = {k: sum(v[k] for v in reg.values()) for k in ("rev", "cost", "margin", "n", "gap")}
    ranked = sorted(t["regions"], key=lambda r: -reg[r]["rev"])
    assert len({round(reg[r]["rev"], 6) for r in reg}) == len(reg), "tie in region revenue"
    cats, seen = {}, []
    for o in orders:
        if o["cat"] not in seen:
            seen.append(o["cat"])
    shipping = sum(o["ship"] for o in orders)
    return dict(I=I, products=products, orders=orders, reg=reg, total=total, ranked=ranked, cats=t["categories"], seen=seen,
                shipping=shipping, top3=sorted((o["rev"] for o in orders), reverse=True)[:3])


M = model(truth)
I, T = M["I"], M["total"]
NORD, NPROD = len(M["orders"]), len(M["products"])
LAST = NORD + 1  # last row of the Orders table


def grids():
    """{sheet: {cell: (value, formula?)}}: every cell of the end state.
    Dynamic arrays: the anchor is the formula, the rest of the spill values."""
    g = {}
    F = lambda v: (v, True)
    C = lambda v: (v, False)
    s = g["Inputs"] = {}
    for a, v in [("A1", "Assumption"), ("B1", "Value"), ("A2", "Tax rate"), ("B2", I["tax"]), ("A3", "Bulk threshold (units)"),
                 ("B3", I["bulk_units"]), ("A4", "Bulk discount"), ("B4", I["bulk_disc"]), ("A5", "Target margin"),
                 ("B5", I["target_margin"]), ("A6", "Shipping per order"), ("B6", I["ship"]), ("D1", "Regions"), ("G1", "Categories")]:
        s[a] = C(v)
    for i, r in enumerate(truth["regions"]):
        s[f"D{2 + i}"] = C(r)
    for i, r in enumerate(M["cats"]):
        s[f"G{2 + i}"] = C(r)
    s = g["Products"] = {}
    for a, v in zip("ABCDE", ["Product", "Category", "Price", "Cost", "Margin"]):
        s[f"{a}1"] = C(v)
    for i, p in enumerate(M["products"]):
        r = i + 2
        s[f"A{r}"], s[f"B{r}"], s[f"C{r}"], s[f"D{r}"], s[f"E{r}"] = C(p["name"]), C(p["cat"]), C(p["price"]), C(p["cost"]), F(p["margin"])
    s = g["Orders"] = {}
    for a, v in zip("ABCDEFGHI", ["Date", "Region", "Product", "Units", "Category", "Price", "Revenue", "Cost", "Shipping"]):
        s[f"{a}1"] = C(v)
    for i, o in enumerate(M["orders"]):
        r = i + 2
        for a, v in zip("ABCD", [o["date"], o["region"], o["product"], o["units"]]):
            s[f"{a}{r}"] = C(v)
        for a, k in zip("EFGHI", ["cat", "price", "rev", "cost", "ship"]):
            s[f"{a}{r}"] = F(o[k])
    s = g["Summary"] = {}
    s["A1"] = C("Sales summary")
    for a, v in zip("ABCDEFGH", ["Region", "Revenue", "Cost", "Margin", "Share", "Margin %", "Orders", "Gap to target"]):
        s[f"{a}3"] = C(v)
    for i, r in enumerate(truth["regions"]):
        row, d = 4 + i, M["reg"][r]
        s[f"A{row}"] = C(r)
        for a, v in zip("BCDEFGH", [d["rev"], d["cost"], d["margin"], d["rev"] / T["rev"], d["margin"] / d["rev"], d["n"], d["gap"]]):
            s[f"{a}{row}"] = F(v)
    s["A8"] = C("Total")
    for a, v in zip("BCDEFGH", [T["rev"], T["cost"], T["margin"], 1.0, T["margin"] / T["rev"], T["n"], T["gap"]]):
        s[f"{a}8"] = F(v)
    s["A10"], s["B10"] = C("Tax"), F(T["rev"] * I["tax"])
    s["A11"], s["B11"] = C("Revenue after tax"), F(T["rev"] * (1 - I["tax"]))
    s["A13"], s["B13"] = C("Category"), C("Revenue")
    for i, cat in enumerate(M["cats"]):
        s[f"A{14 + i}"], s[f"B{14 + i}"] = C(cat), F(sum(o["rev"] for o in M["orders"] if o["cat"] == cat))
    s["A17"], s["B17"] = C("Total"), F(T["rev"])
    s["A19"], s["B19"] = C("Products below target margin"), F(sum(p["margin"] < I["target_margin"] for p in M["products"]))
    s["A20"], s["B20"] = C("Top region"), F(M["ranked"][0])
    s["A21"], s["B21"] = C("Average order value"), F(T["rev"] / T["n"])
    s["A22"], s["B22"] = C("Shipping cost"), F(M["shipping"])
    s["A23"], s["B23"] = C("Margin after shipping"), F(T["margin"] - M["shipping"])
    s = g["Rankings"] = {}
    for a, v in [("A1", "Regions by revenue"), ("B1", "Revenue"), ("D1", "Three largest orders"), ("F1", "Categories sold")]:
        s[a] = C(v)
    for i, r in enumerate(M["ranked"]):
        s[f"A{2 + i}"], s[f"B{2 + i}"] = (F if i == 0 else C)(r), (F if i == 0 else C)(M["reg"][r]["rev"])
    for i, v in enumerate(M["top3"]):
        s[f"D{2 + i}"] = (F if i == 0 else C)(v)
    for i, v in enumerate(M["seen"]):
        s[f"F{2 + i}"] = (F if i == 0 else C)(v)
    s = g["Targets"] = {"A1": C("Region"), "B1": C("Target")}
    for i, (r, v) in enumerate(TARGETS):
        s[f"A{2 + i}"], s[f"B{2 + i}"] = C(r), C(v)
    return g


G = grids()
RANGE = {"Inputs": "A1:H12", "Products": "A1:G14", "Orders": "A1:K24", "Summary": "A1:J30", "Rankings": "A1:H12", "Targets": "A1:D9"}


def norm(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return (v.date() if isinstance(v, dt.datetime) else v).isoformat()
    if isinstance(v, str):
        return v.strip()
    return v


def same(a, b):
    a, b = norm(a), norm(b)
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return abs(a - b) <= 1e-6 * max(1, abs(b))
    return a == b


c = Checks()
try:
    wb = openpyxl.load_workbook(path)
except Exception as e:
    c.add("the workbook opens", False, e)
    c.report()
    raise SystemExit
names = wb.sheetnames
_lo = {}


def lo(sheet):
    """{cell: value} of a sheet's expected range, from LibreOffice."""
    if sheet not in _lo:
        rows, cols = lo_values(path, f"'{sheet}'!{RANGE[sheet]}")
        _lo[sheet] = {f"{col}{r}": v for r, vals in rows.items() for col, v in zip(cols, vals)}
    return _lo[sheet]


def isformula(sheet, addr):
    v = wb[sheet][addr].value
    return (isinstance(v, str) and v.startswith("=")) or type(v).__name__ == "ArrayFormula"


def ftext(sheet, addr):
    v = wb[sheet][addr].value
    return getattr(v, "text", v) if v is not None else ""


def cell_ok(sheet, addr):
    return same(lo(sheet)[addr], G[sheet][addr][0])


def cells_ok(sheet, addrs, need_formula=False):
    bad = [f"{a}: {lo(sheet)[a]!r}, expected {G[sheet][a][0]!r}" for a in addrs if not cell_ok(sheet, a)]
    bad += [f"{a} is not a formula" for a in addrs if need_formula and not isformula(sheet, a)]
    return not bad, "; ".join(bad[:3])


def find(sheet, col, value, first_row=2):
    """The row where a column of the actual sheet holds a value."""
    for addr, v in lo(sheet).items():
        if addr.startswith(col) and addr[len(col):].isdigit() and int(addr[len(col):]) >= first_row and norm(v) == norm(value):
            return int(addr[len(col):])
    return None


def expected_row(sheet, key_col, key):
    for addr, (v, _) in G[sheet].items():
        if addr.startswith(key_col) and addr[len(key_col):].isdigit() and int(addr[len(key_col):]) >= 2 and norm(v) == norm(key):
            return int(addr[len(key_col):])


def need(sheet):
    if sheet not in names:
        raise AssertionError(f"no sheet {sheet}")


def table_ref(sheet, name):
    return wb[sheet].tables[name].ref if sheet in names and name in wb[sheet].tables else None


# ---- the 20 edits, each by its direct effect
def e1():
    return cells_ok("Inputs", ["B2"])


def e2():
    r = find("Products", "A", "Gadget")
    return r is not None and same(lo("Products")[f"C{r}"], 34), f"Gadget row {r}: {lo('Products').get(f'C{r}')}"


def e3():
    r = find("Products", "A", "Gizmo Pro")
    if r is None:
        return False, "no Gizmo Pro"
    got = [lo("Products")[f"{a}{r}"] for a in "BCD"]
    ok = got == ["Hardware", 48, 21] and isformula("Products", f"E{r}") and same(lo("Products")[f"E{r}"], 27 / 48)
    t = table_ref("Products", "Products")
    return ok and t and range_boundaries(t)[3] >= r, f"{got}, formula {isformula('Products', f'E{r}')}, table {t}"


def e4():
    r = find("Orders", "A", "2026-03-18")
    if r is None:
        return False, "no order of 2026-03-18 (as a date)"
    d = lo("Orders")[f"A{r}"]
    exp = M["orders"][-1]
    ok = (isinstance(d, dt.datetime) and [lo("Orders")[f"{a}{r}"] for a in "BCD"] == ["North", "Gizmo Pro", 12]
          and all(isformula("Orders", f"{a}{r}") for a in "EFGH") and cells_ok("Orders", [f"{a}{r}" for a in "EFGH"])[0])
    t = table_ref("Orders", "Orders")
    return ok and t and range_boundaries(t)[3] >= r, f"row {r}: {[lo('Orders')[f'{a}{r}'] for a in 'ABCDEFGH']}, table {t}"


def e5():
    r = find("Orders", "A", "2026-01-14")
    return r is not None and lo("Orders")[f"B{r}"] == "East" and lo("Orders")[f"D{r}"] == 4, f"row {r}: {r and lo('Orders')[f'D{r}']}"


def e6():
    r = find("Orders", "A", "2026-01-19")
    return r is not None and lo("Orders")[f"C{r}"] == "Support" and lo("Orders")[f"B{r}"] == "South", f"row {r}: {r and lo('Orders')[f'B{r}']}"


def e7():
    return cells_ok("Inputs", ["B3"])


def e8():
    s = lo("Summary")
    ok, d = cells_ok("Summary", ["G3", "G4", "G5", "G6", "G7", "G8"])
    return ok and all(isformula("Summary", f"G{r}") for r in range(4, 9)), d or "G4:G8 must be formulas"


def e9():
    ok, d = cells_ok("Summary", ["A21", "B21"], need_formula=False)
    return ok and isformula("Summary", "B21"), d or "B21 must be a formula"


def e10():
    gone = "Insights" not in names and "Rankings" in names
    t = ftext("Summary", "B20")
    return gone and "Rankings" in str(t) and cell_ok("Summary", "B20"), f"sheets {names}, B20 = {t}"


def e11():
    r = find("Products", "A", "Toolkit")
    return r is not None and lo("Products")[f"B{r}"] == "Service", f"row {r}"


def e12():
    r = find("Orders", "A", "2026-01-27")
    t = table_ref("Orders", "Orders")
    blank = [a for a in range(2, LAST + 1) if lo("Orders").get(f"A{a}") is None]
    return r is None and t and range_boundaries(t)[3] == LAST and not blank, f"2026-01-27 at {r}, table {t}, blank rows {blank}"


def e13():
    ok = bool(names) and names[-1] == "Targets"
    return ok and cells_ok("Targets", list(G["Targets"]))[0], f"sheets {names}; " + cells_ok("Targets", list(G["Targets"]))[1]


def e14():
    ok, d = cells_ok("Summary", ["H3", "H4", "H5", "H6", "H7", "H8"])
    uses = all("Targets" in str(ftext("Summary", f"H{r}")) for r in range(4, 8))
    return ok and uses, d or "H4:H7 must look the targets up"


def e15():
    return cells_ok("Inputs", ["B5"])


def e16():
    fmts = [wb["Summary"][f"F{r}"].number_format for r in range(4, 9)]
    return all(f == "0.0%" for f in fmts), fmts


def e17():
    got = [lo("Products").get(f"A{r}") for r in range(2, NPROD + 2)]
    want = [p["name"] for p in M["products"]]
    ok, d = cells_ok("Products", [f"{a}{r}" for r in range(2, NPROD + 2) for a in "ABCD"])
    return got == want and ok, f"{got}; {d}"


def e18():
    ok, d = cells_ok("Inputs", ["A6", "B6"])
    n = wb.defined_names.get("ShipPerOrder")
    ref = n.attr_text.replace("$", "") if n else None
    return ok and ref == "Inputs!B6", f"{d} name -> {ref}"


def e19():
    t = table_ref("Orders", "Orders")
    cols = [col.name for col in wb["Orders"].tables["Orders"].tableColumns] if t else []
    rows = range(2, LAST + 1)
    uses = all("ShipPerOrder" in str(ftext("Orders", f"I{r}")) and "4.5" not in str(ftext("Orders", f"I{r}")) for r in rows)
    ok, d = cells_ok("Orders", ["I1"] + [f"I{r}" for r in rows])
    return t and t.endswith(f"I{LAST}") and cols[-1:] == ["Shipping"] and uses and ok, f"table {t}, {cols[-1:]}, names used {uses}; {d}"


def e20():
    ok, d = cells_ok("Summary", ["A22", "B22", "A23", "B23"])
    return ok and isformula("Summary", "B22") and isformula("Summary", "B23"), d or "B22 and B23 must be formulas"


def e21():
    return mentions(text, T["rev"]) and mentions(text, T["margin"] - M["shipping"]), f"wanted {T['rev']:.2f} and {T['margin'] - M['shipping']:.2f}"


text = answer
EDITS = [
    ("edit 1: the tax rate is 22%", e1), ("edit 2: Gadget costs 34.00", e2), ("edit 3: Gizmo Pro is a product, with its margin formula", e3),
    ("edit 4: the 2026-03-18 order is in the table, a date, with its formulas", e4), ("edit 5: the 2026-01-14 order is 4 units", e5),
    ("edit 6: the 2026-01-19 order is South's", e6), ("edit 7: the bulk threshold is 8 units", e7),
    ("edit 8: Summary counts orders per region, by formula", e8), ("edit 9: Summary has the average order value", e9),
    ("edit 10: Insights is renamed Rankings and Summary follows", e10), ("edit 11: Toolkit is a Service", e11),
    ("edit 12: the 2026-01-27 order is gone and the table has no gap", e12), ("edit 13: the Targets sheet is last, as asked", e13),
    ("edit 14: Summary has the gap to target, from Targets", e14), ("edit 15: the target margin is 55%", e15),
    ("edit 16: Margin % shows one decimal", e16), ("edit 17: Products sorted by price, highest first", e17),
    ("edit 18: the shipping input exists and is named", e18), ("edit 19: Orders has a Shipping column, by formula and name", e19),
    ("edit 20: Summary has shipping cost and margin after shipping", e20), ("the reply states the total revenue and the margin after shipping", e21),
]
for name, fn in EDITS:
    c.guard(name, fn)

# ---- every cell as the edits leave it (values by LibreOffice), and formulas stay formulas
for sheet in G:
    def cells():
        need(sheet)
        got, bad = lo(sheet), []
        for addr in sorted(set(got) | set(G[sheet])):
            exp = G[sheet].get(addr, (None, False))[0]
            if not same(got.get(addr), exp):
                bad.append(f"{addr}: {got.get(addr)!r}, expected {exp!r}")
        return not bad, f"{len(bad)} differ: " + "; ".join(bad[:3])
    c.guard(f"{sheet}: every cell is as the edits leave it", cells)
for sheet in ("Products", "Orders", "Summary", "Rankings"):
    def formulas():
        need(sheet)
        lost = [a for a, (_, f) in G[sheet].items() if f and not isformula(sheet, a)]
        return not lost, f"not formulas any more: {lost[:6]}"
    c.guard(f"{sheet}: formulas are still formulas", formulas)

# ---- what the workbook held stays
z = zipfile.ZipFile(path)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def sheet_xml(name):
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = next(s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
               for s in wbx.find("m:sheets", NS) if s.get("name") == name)
    target = next(r.get("Target") for r in rels if r.get("Id") == rid).lstrip("/")
    return z.read(target if target.startswith("xl/") else "xl/" + target).decode()


def covers(sqref, rows, col):
    """Whether a range list covers column COL's rows (first, last)."""
    for part in sqref.split():
        c1, r1, c2, r2 = range_boundaries(part)
        if c1 <= openpyxl.utils.column_index_from_string(col) <= c2 and r1 <= rows[0] and r2 >= rows[1]:
            return True
    return False


def cf_rules(sheet):
    x = ET.fromstring(sheet_xml(sheet))
    out = []
    for cf in x.findall("m:conditionalFormatting", NS):
        for r in cf.findall("m:cfRule", NS):
            out.append((r.get("type"), r.get("operator"), cf.get("sqref"), "".join(f.text or "" for f in r.findall("m:formula", NS)), r.get("rank")))
    return out


def dvs(sheet):
    x = ET.fromstring(sheet_xml(sheet))
    d = x.find("m:dataValidations", NS)
    return [] if d is None else [(v.get("type"), v.get("operator"), v.get("sqref"), (v.findtext("m:formula1", "", NS)), (v.findtext("m:formula2", "", NS)), v) for v in d]


def f_tables():
    exp = {("Products", "Products"): ["Product", "Category", "Price", "Cost", "Margin"],
           ("Orders", "Orders"): ["Date", "Region", "Product", "Units", "Category", "Price", "Revenue", "Cost", "Shipping"]}
    bad = []
    for (sheet, name), cols in exp.items():
        t = wb[sheet].tables.get(name) if sheet in names else None
        if t is None:
            bad.append(f"no table {name}")
            continue
        if [x.name for x in t.tableColumns] != cols:
            bad.append(f"{name} columns {[x.name for x in t.tableColumns]}")
        if t.tableStyleInfo is None or t.tableStyleInfo.name != ("TableStyleMedium2" if name == "Products" else "TableStyleMedium9"):
            bad.append(f"{name} style")
    return not bad, "; ".join(bad)


def f_names():
    want = {"TaxRate": "Inputs!B2", "BulkUnits": "Inputs!B3", "BulkDisc": "Inputs!B4", "TargetMargin": "Inputs!B5",
            "RegionList": "Inputs!D2:D5", "CategoryList": "Inputs!G2:G4"}
    got = {n: d.attr_text.replace("$", "") for n, d in wb.defined_names.items()}
    bad = [f"{n} -> {got.get(n)}" for n, ref in want.items() if got.get(n) != ref]
    return not bad, "; ".join(bad)


def f_chart():
    charts = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    if len(charts) != 1:
        return False, f"{len(charts)} chart parts"
    x = z.read(charts[0]).decode()
    on_summary = any("charts/chart" in z.read(n).decode() for n in z.namelist() if re.match(r"xl/drawings/_rels/drawing\d+\.xml\.rels$", n))
    refs = [f.replace("'", "").replace("$", "") for f in re.findall(r"<(?:c:)?f>([^<]*)</(?:c:)?f>", x)]
    ok = "Revenue by region" in x and "barChart" in x and "Summary!A4:A7" in refs and "Summary!B4:B7" in refs and on_summary
    return ok, f"title {'Revenue by region' in x}, bar {'barChart' in x}, series {refs}"


def f_cf_orders():
    need("Orders")
    r = cf_rules("Orders")
    ok = (any(t == "cellIs" and op == "greaterThanOrEqual" and covers(s, (2, LAST), "D") and f == "BulkUnits" for t, op, s, f, _ in r)
          and any(t == "top10" and rank == "3" and covers(s, (2, LAST), "G") for t, _, s, _, rank in r))
    return ok, str(r)


def f_cf_summary():
    need("Summary")
    r = cf_rules("Summary")
    ok = (any(t == "cellIs" and op == "lessThan" and covers(s, (4, 7), "F") and f == "TargetMargin" for t, op, s, f, _ in r)
          and any(t == "dataBar" and covers(s, (4, 7), "E") for t, _, s, _, _ in r))
    return ok, str(r)


def f_cf_products():
    need("Products")
    r = cf_rules("Products")
    return any(t == "colorScale" and covers(s, (2, NPROD + 1), "E") for t, _, s, _, _ in r), str(r)


def f_dv_orders():
    need("Orders")
    d = dvs("Orders")
    ok = (any(t == "list" and covers(s, (2, LAST), "B") and f1 == "RegionList" for t, _, s, f1, _, _ in d)
          and any(t == "whole" and op == "greaterThanOrEqual" and covers(s, (2, LAST), "D") and f1 == "1" for t, op, s, f1, _, _ in d))
    return ok, str([x[:5] for x in d])


def f_dv_products():
    need("Products")
    d = dvs("Products")
    return any(t == "list" and covers(s, (2, NPROD + 1), "B") and f1 == "CategoryList" for t, _, s, f1, _, _ in d), str([x[:5] for x in d])


def f_dv_inputs():
    need("Inputs")
    d = dvs("Inputs")
    ok = any(t == "decimal" and covers(s, (2, 2), "B") and f1 == "0" and f2 == "1" and v.get("promptTitle") == "Tax rate"
             and v.get("errorTitle") == "Not a rate" and "0% and 100%" in (v.get("error") or "") for t, _, s, f1, f2, v in d)
    return ok, str([x[:5] for x in d])


def f_arrays():
    need("Rankings")
    x = sheet_xml("Rankings")
    bad = []
    for addr, ref, fn in [("A2", "A2:A5", "SORTBY"), ("B2", "B2:B5", "SORT("), ("D2", "D2:D4", "LARGE"), ("F2", "F2:F4", "UNIQUE")]:
        m = re.search(rf'<c r="{addr}"([^>]*)>(.*?)</c>', x, re.S)
        if not m or 'cm="1"' not in m.group(1) or f't="array" ref="{ref}"' not in m.group(2) or fn not in m.group(2):
            bad.append(f"{addr}: {m.group(0)[:120] if m else 'missing'}")
    if "xl/metadata.xml" not in z.namelist():
        bad.append("no metadata.xml")
    return not bad, "; ".join(bad[:2])


def f_view():
    bad = [s for s in ("Orders", "Products") if s in names and 'state="frozen"' not in sheet_xml(s)]
    return not bad, f"not frozen: {bad}"


def f_formats():
    o, p, s = wb["Orders"], wb["Products"], wb["Summary"]
    bad = []
    for r in range(2, LAST + 1):
        got = [o[f"A{r}"].number_format] + [o[f"{a}{r}"].number_format for a in "FGH"]
        if got != ["yyyy-mm-dd", "#,##0.00", "#,##0.00", "#,##0.00"]:
            bad.append(f"Orders row {r}: {got}")
    for r in range(2, NPROD + 2):
        got = [p[f"{a}{r}"].number_format for a in "CDE"]
        if got != ["#,##0.00", "#,##0.00", "0%"]:
            bad.append(f"Products row {r}: {got}")
    got = [s[f"B{r}"].number_format for r in range(4, 9)] + [s[f"E{r}"].number_format for r in range(4, 9)]
    if got != ["#,##0.00"] * 5 + ["0%"] * 5:
        bad.append(f"Summary: {got}")
    return not bad, "; ".join(bad[:3])


for name, fn in [("tables keep their names, columns and style", f_tables), ("the named cells and ranges are kept", f_names),
                 ("the chart is kept, on the region revenue", f_chart), ("Orders' conditional formats cover the orders", f_cf_orders),
                 ("Summary's conditional formats are kept", f_cf_summary), ("Products' colour scale is kept", f_cf_products),
                 ("Orders' validations are kept", f_dv_orders), ("Products' category list is kept", f_dv_products),
                 ("the tax rate's validation and messages are kept", f_dv_inputs), ("the dynamic arrays are still dynamic arrays", f_arrays),
                 ("the frozen header rows are kept", f_view), ("number formats are kept, new rows look like the rest", f_formats)]:
    c.guard(name, fn)

# ---- stored results, as pandas and previews read them (not part of the pass)
def st(sheet, addr, want):
    v = stored(path, sheet, addr)
    return same(v, want), v


c.guard("total revenue is stored", lambda: st("Summary", "B8", T["rev"]), kind="stored")
c.guard("margin after shipping is stored", lambda: st("Summary", "B23", T["margin"] - M["shipping"]), kind="stored")
c.guard("an order's revenue is stored", lambda: st("Orders", "G2", M["orders"][0]["rev"]), kind="stored")
c.report()
