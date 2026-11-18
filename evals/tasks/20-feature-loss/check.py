# Grades six edits to a workbook that holds a chart, validations, conditional
# formats, a note, frozen panes and a shape: the edits on the end state, and
# each feature kept. Formula values come only from evalkit.lo_values (a
# LibreOffice recalculation of a copy); the expected values are computed
# here, in Python, from the inputs setup.py saved. Features are read from
# the package's XML or with openpyxl.
import datetime as dt
import re
import warnings
import zipfile
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.utils import column_index_from_string, range_boundaries
from evalkit import Checks, args, lo_values, mentions, stored

warnings.filterwarnings("ignore")  # openpyxl on extensions it does not read

work, truth, answer = args()
path = f"{work}/budget.xlsx"


def model(t):
    """The budget after the six edits: rows of [dept, owner, q1..q4, status, review]."""
    rows = [list(d) for d in t["depts"]]
    at = lambda name: next(r for r in rows if r[0] == name)
    legal = ["Legal", "Priya", 40000, 42000, 42000, 45000, "On track", "2026-06-30"]
    rows.insert([r[0] for r in rows].index("HR"), legal)                  # 1
    eng = at("Engineering")
    eng[4], eng[5] = round(eng[4] * 1.1), round(eng[5] * 1.1)             # 2
    at("Support")[6] = "At risk"                                          # 3
    at("HR")[1] = "Lena"                                                  # 4
    at("Marketing")[7] = "2026-09-30"                                     # 5
    return rows


ROWS = model(truth)
N = len(ROWS)
TOT = 4 + N            # the Total row
AVG = TOT + 1          # the row under it
LASTD = TOT - 1        # the last department row
QS = [sum(r[k] for r in ROWS) for k in range(2, 6)]
GRAND = sum(QS)
EXP = {}               # cell: (value, formula?)
for a, v in zip("ABCDEFGHIJ", ["Department", "Owner", "Q1", "Q2", "Q3", "Q4", "Total", "Share", "Status", "Review date"]):
    EXP[f"{a}3"] = (v, False)
EXP["A1"] = ("Department budget 2026", False)
for i, r in enumerate(ROWS):
    n = 4 + i
    for a, v in zip("ABCDEF", r[:6]):
        EXP[f"{a}{n}"] = (v, False)
    EXP[f"G{n}"], EXP[f"H{n}"] = (sum(r[2:6]), True), (sum(r[2:6]) / GRAND, True)
    EXP[f"I{n}"], EXP[f"J{n}"] = (r[6], False), (r[7], False)
EXP[f"A{TOT}"] = ("Total", False)
for a, v in zip("CDEFGH", QS + [GRAND, 1.0]):
    EXP[f"{a}{TOT}"] = (v, True)
EXP[f"A{AVG}"] = ("Quarterly average", False)
for a, v in zip("CDEF", QS):
    EXP[f"{a}{AVG}"] = (v / N, True)
ENG_SHARE = sum(next(r for r in ROWS if r[0] == "Engineering")[2:6]) / GRAND


def norm(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return (v.date() if isinstance(v, dt.datetime) else v).isoformat()
    return v.strip() if isinstance(v, str) else v


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
    ws = wb["Budget"]
except Exception as e:
    c.add("the workbook opens", False, e)
    c.report()
    raise SystemExit
rows_, cols_ = lo_values(path, "Budget!A1:N24")
LO = {f"{col}{r}": v for r, vals in rows_.items() for col, v in zip(cols_, vals)}


def isformula(addr):
    v = ws[addr].value
    return (isinstance(v, str) and v.startswith("=")) or type(v).__name__ == "ArrayFormula"


def cells_ok(addrs, need_formula=False):
    bad = [f"{a}: {LO.get(a)!r}, expected {EXP[a][0]!r}" for a in addrs if not same(LO.get(a), EXP[a][0])]
    bad += [f"{a} is not a formula" for a in addrs if need_formula and not isformula(a)]
    return not bad, "; ".join(bad[:3])


def row_of(name):
    return next((int(a[1:]) for a, v in LO.items() if a[0] == "A" and a[1:].isdigit() and int(a[1:]) >= 4 and v == name), None)


def row_cells(name):
    n = row_of(name)
    return n, [f"{a}{n}" for a in "ABCDEFGHIJ"]


# ---- the edits
def e1():
    n = row_of("Legal")
    order = [LO.get(f"A{r}") for r in range(4, TOT)]
    want = [r[0] for r in ROWS]
    ok, d = cells_ok([f"{a}{n}" for a in "ABCDEFGHIJ"], need_formula=False) if n else (False, "no Legal row")
    return n and order == want and ok and isformula(f"G{n}") and isformula(f"H{n}") and isinstance(LO.get(f"J{n}"), dt.datetime), f"rows {order}; {d}"


def e2():
    n = row_of("Engineering")
    return n and cells_ok([f"E{n}", f"F{n}"])[0] and LO[f"E{n}"] == 352000 and LO[f"F{n}"] == 363000, f"row {n}: {n and (LO[f'E{n}'], LO[f'F{n}'])}"


def e3():
    n = row_of("Support")
    return n and LO[f"I{n}"] == "At risk", n and LO[f"I{n}"]


def e4():
    n = row_of("HR")
    return n and LO[f"B{n}"] == "Lena", n and LO[f"B{n}"]


def e5():
    n = row_of("Marketing")
    return n and isinstance(LO[f"J{n}"], dt.datetime) and norm(LO[f"J{n}"]) == "2026-09-30", n and LO[f"J{n}"]


def e6():
    ok, d = cells_ok([f"A{AVG}"] + [f"{a}{AVG}" for a in "CDEF"])
    return ok and all(isformula(f"{a}{AVG}") for a in "CDEF"), d or "the averages must be formulas"


def e6b():
    ok, d = cells_ok([f"{a}{TOT}" for a in "CDEFGH"], need_formula=True)
    return ok and cells_ok([f"{a}{n}" for n in range(4, TOT) for a in "GH"], need_formula=True)[0], d or "a total or a share misses a row"


def e7():
    return mentions(answer, GRAND) and (mentions(answer, ENG_SHARE * 100, 0.01) or mentions(answer, ENG_SHARE, 0.01)), \
        f"wanted {GRAND:,} and Engineering's {ENG_SHARE:.1%}"


for name, fn in [("edit 1: Legal sits between Finance and HR, with its formulas", e1), ("edit 2: Engineering's Q3 and Q4 are up 10%", e2),
                 ("edit 3: Support is At risk", e3), ("edit 4: HR's owner is Lena", e4), ("edit 5: Marketing's review is 2026-09-30, a date", e5),
                 ("edit 6: the Quarterly average row is under the total, by formula", e6),
                 ("the totals and shares cover every department", e6b),
                 ("the reply states the new total and Engineering's share", e7)]:
    c.guard(name, fn)


def whole():
    bad = [f"{a}: {LO.get(a)!r}, expected {EXP.get(a, (None,))[0]!r}" for a in sorted(set(LO) | set(EXP)) if not same(LO.get(a), EXP.get(a, (None,))[0])]
    return not bad, f"{len(bad)} differ: " + "; ".join(bad[:3])


def formulas():
    lost = [a for a, (_, f) in EXP.items() if f and not isformula(a)]
    return not lost, f"not formulas any more: {lost[:6]}"


c.guard("every cell is as the edits leave it", whole)
c.guard("formulas are still formulas", formulas)

# ---- what the workbook held stays
z = zipfile.ZipFile(path)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def sheet_xml():
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = next(s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
               for s in wbx.find("m:sheets", NS) if s.get("name") == "Budget")
    target = next(r.get("Target") for r in rels if r.get("Id") == rid).lstrip("/")
    return z.read(target if target.startswith("xl/") else "xl/" + target).decode()


SX = sheet_xml()
DRAW = "".join(z.read(n).decode() for n in z.namelist() if re.match(r"xl/drawings/drawing\d+\.xml$", n))


def covered(sqrefs, cols):
    """Whether the ranges together cover every department row of the columns
    (a tool may split a range around a new row; what counts is that the
    row is covered)."""
    have = set()
    for sq in sqrefs:
        for part in sq.split():
            c1, r1, c2, r2 = range_boundaries(part)
            have |= {(c, r) for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}
    return all((column_index_from_string(k), r) in have for k in cols for r in range(4, LASTD + 1))


def cf():
    x = ET.fromstring(SX)
    return [(r.get("type"), r.get("operator"), cf_.get("sqref"), "".join(f.text or "" for f in r.findall("m:formula", NS)), r)
            for cf_ in x.findall("m:conditionalFormatting", NS) for r in cf_.findall("m:cfRule", NS)]


def dvs():
    x = ET.fromstring(SX)
    d = x.find("m:dataValidations", NS)
    return [] if d is None else list(d)


def has_cf(kind, cols, pred=lambda *a: True):  # pred(operator, formula, rule, sqref)
    rules = cf()
    mine = [s for t, op, s, f, r in rules if t == kind and pred(op, f, r, s)]
    return covered(mine, cols), str([(t, op, s, f) for t, op, s, f, _ in rules])


def f_chart():
    charts = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    if len(charts) != 1:
        return False, f"{len(charts)} chart parts"
    x = z.read(charts[0]).decode()
    refs = [f.replace("'", "").replace("$", "") for f in re.findall(r"<(?:c:)?f>([^<]*)</(?:c:)?f>", x)]
    first = int(re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", DRAW.split("graphicFrame")[0], re.S).group(1)) if "graphicFrame" in DRAW else -1
    ok = "Budget by department" in x and "barChart" in x and f"Budget!A4:A{LASTD}" in refs and f"Budget!G4:G{LASTD}" in refs and first >= AVG
    return ok, f"series {refs}, anchored at row {first + 1}"


def f_shape():
    sp = re.findall(r"<xdr:sp\b.*?</xdr:sp>", DRAW, re.S)
    ok = any('prst="roundRect"' in s and truth["shape"] in s and 'txBox="1"' not in s for s in sp)
    kinds = [re.findall(r'prst="(\w+)"', s) for s in sp]
    return ok, f"{len(sp)} shapes: {kinds}"


def f_cf_scale():
    return has_cf("colorScale", "CDEF")


def f_cf_bar():
    return has_cf("dataBar", "G")


def f_cf_share():
    return has_cf("cellIs", "H", lambda op, f, r, s: op == "greaterThan" and f == "0.25" and r.get("dxfId") is not None)


def f_cf_status():
    return has_cf("cellIs", "I", lambda op, f, r, s: op == "equal" and f == '"Over"' and r.get("dxfId") is not None)


def f_cf_date():
    return has_cf("expression", "J", lambda op, f, r, s: f.replace(" ", "") == f"$J{top(s)[1:]}<DATE(2026,7,15)" and r.get("dxfId") is not None)


def has_dv(cols, pred):
    d = dvs()
    return covered([v.get("sqref") for v in d if pred(v)], cols), str([(v.get("type"), v.get("sqref")) for v in d])


def top(sqref):
    return sqref.split()[0].split(":")[0]


def f1(v):
    return v.findtext("m:formula1", "", NS)


def f_dv_status():
    return has_dv("I", lambda v: v.get("type") == "list" and f1(v) == '"On track,At risk,Over"'
                  and v.get("promptTitle") == "Status" and "Pick one" in (v.get("prompt") or ""))


def f_dv_amount():
    return has_dv("CDEF", lambda v: v.get("type") == "whole" and f1(v) == "0" and v.findtext("m:formula2", "", NS) == "10000000"
                  and v.get("errorTitle") == "Budget")


def f_dv_date():
    return has_dv("J", lambda v: v.get("type") == "date" and v.get("operator") == "greaterThanOrEqual" and f1(v) == "46023")


def f_dv_owner():
    # the formula is relative to the first cell of its range
    return has_dv("B", lambda v: v.get("type") == "custom" and f1(v).replace(" ", "") == f"LEN({top(v.get('sqref'))})>=2")


def f_view():
    pane = re.search(r"<pane [^>]*>", SX)
    return bool(pane) and 'state="frozen"' in pane.group(0) and 'xSplit="1"' in pane.group(0) and 'ySplit="3"' in pane.group(0), pane and pane.group(0)


def f_note():
    n = row_of("HR")
    cm = ws[f"A{n}"].comment if n else None
    others = [a for a in ("A4", "A5", "A6", "A7", "A8", "A9", f"A{TOT}") if ws[a].comment is not None and a != f"A{n}"]
    return bool(cm) and truth["note"] in cm.text and not others, f"HR is row {n}; note {cm and cm.text!r}; stray notes {others}"


def f_formats():
    bad = []
    for r in range(4, TOT):
        got = [ws[f"{a}{r}"].number_format for a in "CDEFG"] + [ws[f"H{r}"].number_format, ws[f"J{r}"].number_format]
        if got != ["#,##0"] * 5 + ["0.0%", "yyyy-mm-dd"]:
            bad.append(f"row {r}: {got}")
    widths = [round(ws.column_dimensions[a].width or 0) for a in "ABC"]
    if widths != [17, 11, 12]:
        bad.append(f"widths {widths}")
    return not bad, "; ".join(bad[:3])


for name, fn in [("the chart is kept and covers every department", f_chart), ("the shape is kept: a rounded rectangle with its text", f_shape),
                 ("the colour scale covers every quarter", f_cf_scale), ("the data bars cover every total", f_cf_bar),
                 ("the share rule (over 25%) covers every row", f_cf_share), ("the status rule (Over) covers every row", f_cf_status),
                 ("the review-date rule keeps its formula", f_cf_date), ("the status dropdown and its prompt are kept", f_dv_status),
                 ("the amount validation is kept", f_dv_amount), ("the review-date validation is kept", f_dv_date),
                 ("the owner validation is kept", f_dv_owner), ("the frozen panes are kept", f_view),
                 ("the note stays on HR's row", f_note), ("number formats and widths are kept, the new row included", f_formats)]:
    c.guard(name, fn)


def st(addr, want):
    v = stored(path, "Budget", addr)
    return same(v, want), v


c.guard("the grand total is stored", lambda: st(f"G{TOT}", GRAND), kind="stored")
c.guard("Engineering's share is stored", lambda: st(f"H{row_of('Engineering')}", ENG_SHARE), kind="stored")
c.report()
