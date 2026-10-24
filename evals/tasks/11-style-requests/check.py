import re
import openpyxl
from openpyxl.utils import range_boundaries
from evalkit import Checks, args, close, stored, values

work, truth, answer = args()
path = f"{work}/orders.xlsx"
c = Checks()
wb = openpyxl.load_workbook(path)
ws = wb["Orders"]
rows, total = truth["rows"], truth["total_row"]

fmts = [ws.cell(r, 6).number_format for r in rows + [total]]
c.add("revenue in dollars without decimals, total included",
      all("$" in f and not re.search(r"\.0|\.#", f) for f in fmts), sorted(set(fmts)))


def covers(sqref, cells):
    boxes = [range_boundaries(p) for p in str(sqref).split()]
    return all(any(b[0] <= col <= b[2] and b[1] <= row <= b[3] for b in boxes) for row, col in cells)


def reddish(dxf):
    if dxf is None or dxf.fill is None:
        return False
    for color in (dxf.fill.bgColor, dxf.fill.fgColor, getattr(dxf.fill, "start_color", None)):
        rgb = getattr(color, "rgb", None)
        if isinstance(rgb, str) and len(rgb) >= 6:
            r, g, b = (int(rgb[-6:][i:i + 2], 16) for i in (0, 2, 4))
            if r > 200 and r > g and r > b:
                return True
    return False


# Evaluate each rule's formula on each data row: "$G2" anchored to the
# first row, comparing the status to Late.
# Only $G2 works across a row: a relative G2 tests another column in each
# cell.
def marks(formula, row):
    status = ws.cell(row, 7).value
    f = formula.replace("$G2", f'"{status}"')
    m = re.fullmatch(r'"([^"]*)"\s*=\s*"([^"]*)"', f)
    if m:
        return m.group(1).lower() == m.group(2).lower()
    m = re.fullmatch(r'EXACT\("([^"]*)",\s*"([^"]*)"\)', f, re.I)
    if m:
        return m.group(1) == m.group(2)
    m = re.fullmatch(r'(?:ISNUMBER\()?SEARCH\("([^"]*)",\s*"([^"]*)"\)\)?', f, re.I)
    if m:
        return m.group(1).lower() in m.group(2).lower()
    m = re.fullmatch(r'COUNTIF\("([^"]*)",\s*"([^"]*)"\)\s*(>\s*0|=\s*1)?', f, re.I)
    if m:
        return m.group(1).lower() == m.group(2).lower()
    return None


ok, detail = False, []
for cf in ws.conditional_formatting:
    for rule in cf.rules:
        if rule.type != "expression" or not rule.formula:
            continue
        detail.append((str(cf.sqref), rule.formula[0]))
        cells = [(r, col) for r in rows for col in range(1, 8)]
        got = [marks(rule.formula[0], r) for r in rows]
        if covers(cf.sqref, cells) and None not in got and [r for r, g in zip(rows, got) if g] == truth["late"] and reddish(rule.dxf):
            ok = True
c.add("late rows highlighted in light red, by a rule on Status", ok, detail)

dv_ok, dv_detail = False, []
for dv in ws.data_validations.dataValidation:
    dv_detail.append((dv.type, dv.formula1, str(dv.sqref)))
    items = set()
    f1 = dv.formula1 or ""
    if f1.startswith('"'):
        items = {x.strip() for x in f1.strip('"').split(",")}
    else:
        try:
            ref = f1.lstrip("=")
            sheet, rng = ref.rsplit("!", 1) if "!" in ref else ("Orders", ref)
            src = wb[sheet.strip("'")]
            items = {cell.value for row in src[rng.replace("$", "")] for cell in row if cell.value}
        except Exception:
            pass
    if dv.type == "list" and items == {"Open", "Shipped", "Late"} and covers(dv.sqref, [(r, 7) for r in rows]):
        dv_ok = True
c.add("status dropdown of Open, Shipped, Late", dv_ok, dv_detail)


def blue(cell):
    rgb = (cell.fill.fgColor.rgb or "") if cell.fill and cell.fill.fill_type == "solid" else ""
    if not isinstance(rgb, str) or len(rgb) < 6:
        return False
    r, g, b = (int(rgb[-6:][i:i + 2], 16) for i in (0, 2, 4))
    return b >= r and b >= g and b > 150 and r + g + b > 450


head = [ws.cell(1, col) for col in range(1, 8)]
c.add("header bold on light blue", all(h.font.b and blue(h) for h in head), [(h.font.b, h.fill.fgColor.rgb) for h in head])
c.add("header row and Order ID column frozen", ws.freeze_panes == "B2", ws.freeze_panes)
vals, _ = values(path, "Orders")
c.add("data and formulas intact", ws["F2"].value == "=D2*E2" and ws["F32"].value == "=SUM(F2:F31)" and close(vals[2][5], truth["f2"])
      and [vals[r][6] for r in rows] == [ws.cell(r, 7).value for r in rows], (ws["F2"].value, ws["F32"].value))
c.add("revenue stored", close(stored(path, "Orders", "F2") or 0, truth["f2"]), stored(path, "Orders", "F2"), kind="stored")
c.report()
