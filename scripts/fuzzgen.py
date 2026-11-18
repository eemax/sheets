"""Generate the fuzzer's pool of random workbooks (internal/fuzz/testdata).

Half are written by openpyxl, half by XlsxWriter, each from its own seed:
values of every type, formulas of all three kinds (plain, legacy array
entered with Ctrl+Shift+Enter, dynamic array), Excel tables, defined names,
charts, validations and conditional formats.

Every formula carries its kind in its text, as N("p"), N("l") or N("d")
(N of a text is 0, so the marker changes no result). The fuzzer checks after
each edit that every formula is still stored as the kind it names. After
saving, this script reads each file back and refuses one whose stored kinds
do not match their markers: a producer that stores a kind of its own choice
(XlsxWriter makes any formula calling SEQUENCE a dynamic array) would
otherwise put a false finding into the pool.

It also writes the checkers' must-refuse cases (testdata/refuse/): pool
files with one defect each, which the fuzzer's invariants must catch.

    uv run --with openpyxl --with xlsxwriter python scripts/fuzzgen.py internal/fuzz/testdata

--seed changes the pool; --count the number of files (default 40).
"""
import argparse
import datetime
import io
import os
import random
import re
import sys
import zipfile

import openpyxl
import xlsxwriter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

FIXED_TIME = datetime.datetime(2026, 11, 18, 12, 0, 0)
HEADERS = ["Item", "Units", "Price", "Cost", "Date", "Flag"]
ITEMS = ["Apple", "Pear", "Fig", "Kiwi", "Plum", "Lime", "Date", "Sloe", "N/A", "apple", "Acme Inc.", ""]
SECOND = ["Calc", "Summary", "Q 2026", "Notes", "Model"]


def col(c):
    """Column letter of a 1-based column number."""
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def mark(text, kind):
    """A formula's text with its kind marker added."""
    return f'{text}+N("{kind}")'


class Plan:
    """What one random workbook holds, independent of the producer: cells as
    (sheet, row, col) -> value, formula ("=...") or None, and its arrays,
    tables, names, charts, validations and conditional formats."""

    def __init__(self, rng, dynamic):
        self.rng = rng
        self.sheets = ["Data"]
        n2 = rng.choice([0, 1, 1, 2])
        for name in rng.sample(SECOND, n2):
            self.sheets.append(name)
        self.cells = {}  # (sheet, r, c) -> value
        self.legacy = []  # (sheet, ref, text): Ctrl+Shift+Enter arrays
        self.dynamic = []  # (sheet, ref, text): dynamic arrays
        self.tables = []  # (sheet, name, ref)
        self.names = []  # (name, ref)
        self.charts = []  # (sheet, data rows, anchor)
        self.validations = []  # (sheet, kind, ref)
        self.conds = []  # (sheet, kind, ref)
        self.n = rng.randint(4, 14)  # data rows on Data
        self.build(dynamic)

    def build(self, dynamic):
        rng, n = self.rng, self.n
        last = n + 1
        for c, h in enumerate(HEADERS, 1):
            self.cells[("Data", 1, c)] = h
        for r in range(2, last + 1):
            self.cells[("Data", r, 1)] = rng.choice(ITEMS) or None
            self.cells[("Data", r, 2)] = rng.choice([rng.randint(-5, 40), rng.randint(0, 9), None])
            self.cells[("Data", r, 3)] = rng.choice([round(rng.uniform(0, 100), 2), rng.randint(1, 9)])
            self.cells[("Data", r, 4)] = rng.randint(0, 500)
            self.cells[("Data", r, 5)] = datetime.date(2026, rng.randint(1, 12), rng.randint(1, 28))
            self.cells[("Data", r, 6)] = rng.choice([True, False])
        # A formula column, filled down (plain).
        self.cells[("Data", 1, 7)] = "Total"
        fill = rng.choice(["=B{r}*C{r}", "=IF(B{r}>3,B{r},0)", "=ROUND(C{r}*1.1,2)", "=B{r}+D{r}"])
        for r in range(2, last + 1):
            self.cells[("Data", r, 7)] = "=" + mark(fill[1:].replace("{r}", str(r)), "p")
        # A total row below the data, now and then.
        if rng.random() < 0.5:
            self.cells[("Data", last + 2, 1)] = "Total"
            self.cells[("Data", last + 2, 2)] = "=" + mark(f"SUM(B2:B{last})", "p")
            self.cells[("Data", last + 2, 7)] = "=" + mark(f"SUM(G2:G{last})", "p")
        # Plain formulas reading a range where one value goes: kept plain
        # (Excel shows them with @), the kind WP-18 found move and copy
        # turning into dynamic arrays.
        row = 2
        for _ in range(rng.randint(1, 3)):
            a = rng.randint(2, max(2, last - 2))
            self.cells[("Data", row, 9)] = "=" + mark(f"B{a}:B{a + 1}*{rng.choice([2, 10, 100])}", "p")
            row += 1
        # Legacy arrays: a single cell and one over several cells.
        if rng.random() < 0.7:
            self.legacy.append(("Data", f"J{row}", "=" + mark(f"SUM(B2:B{last}*C2:C{last})", "l")))
            row += 1
        if rng.random() < 0.6:
            h = rng.randint(2, 3)
            self.legacy.append(("Data", f"K2:K{1 + h}", "=" + mark(f"B2:B{1 + h}*{rng.choice([2, 3])}", "l")))
        # Dynamic arrays (XlsxWriter only: openpyxl cannot write them).
        if dynamic:
            if rng.random() < 0.7:
                h = rng.randint(2, 4)
                self.dynamic.append(("Data", f"M2:M{1 + h}", "=" + mark(f"_xlfn.SEQUENCE({h})", "d")))
            if rng.random() < 0.7:
                self.dynamic.append(("Data", "N2", "=" + mark(f"SUM(IF(B2:B{last}>3,B2:B{last}))", "d")))
            if rng.random() < 0.5:
                h = rng.randint(2, 3)
                self.dynamic.append(("Data", f"O2:O{1 + h}", "=" + mark(f"C2:C{1 + h}*2", "d")))
        # Second sheets: values and formulas reading Data.
        for s in self.sheets[1:]:
            ref = rng.choice(["Data", "'Data'"])  # quoting a plain name is fine too
            self.cells[(s, 1, 1)] = "Key"
            self.cells[(s, 1, 2)] = "Value"
            for r in range(2, rng.randint(4, 8)):
                self.cells[(s, r, 1)] = rng.choice(ITEMS[:8])
                self.cells[(s, r, 2)] = rng.choice(
                    [
                        "=" + mark(f"{ref}!B{rng.randint(2, last)}*2", "p"),
                        "=" + mark(f"SUM({ref}!D2:D{last})", "p"),
                        rng.randint(1, 99),
                        "=" + mark(f"COUNTIF({ref}!B2:B{last},\">3\")", "p"),
                    ]
                )
            if rng.random() < 0.4:
                self.legacy.append((s, "D2", "=" + mark(f"MAX({ref}!B2:B{last}-{ref}!D2:D{last})", "l")))
            if rng.random() < 0.5:
                self.cells[(s, 2, 3)] = "=" + mark(f"{ref}!C2:C3*{rng.choice([2, 10])}", "p")
            if dynamic and rng.random() < 0.4:
                self.dynamic.append((s, "E2:E3", "=" + mark(f"{ref}!B2:B3*10", "d")))
        # A defined name, used by a formula.
        if rng.random() < 0.6:
            self.names.append(("Rate", "Data!$C$2"))
            self.cells[("Data", 2, 8)] = "=" + mark("B2*Rate", "p")
        if rng.random() < 0.4:
            self.names.append(("Units", f"Data!$B$2:$B${last}"))
            self.cells[("Data", 3, 8)] = "=" + mark("SUM(Units)", "p")
        # An Excel table over the data, or a chart, validation, conditional
        # format.
        if rng.random() < 0.5:
            self.tables.append(("Data", f"Sales{rng.randint(1, 99)}", f"A1:G{last}"))
        if rng.random() < 0.6:
            self.charts.append(("Data", last, f"Q{rng.randint(2, 6)}"))
        if rng.random() < 0.6:
            self.validations.append(("Data", rng.choice(["list", "whole"]), f"A2:B{last}"))
        if rng.random() < 0.6:
            self.conds.append(("Data", rng.choice(["cell", "expr"]), f"B2:D{last}"))


def write_openpyxl(plan, path):
    wb = openpyxl.Workbook()
    wb.properties.created = FIXED_TIME
    wb.properties.modified = FIXED_TIME
    sheets = {}
    for i, name in enumerate(plan.sheets):
        ws = wb.active if i == 0 else wb.create_sheet(name)
        ws.title = name
        sheets[name] = ws
    for (s, r, c), v in sorted(plan.cells.items(), key=lambda kv: (plan.sheets.index(kv[0][0]), kv[0][1], kv[0][2])):
        if v is not None:
            sheets[s].cell(row=r, column=c, value=v)
    for s, ref, text in plan.legacy:
        sheets[s][ref.split(":")[0]] = ArrayFormula(ref, text)
    for name, ref in plan.names:
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    for s, name, ref in plan.tables:
        t = Table(displayName=name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheets[s].add_table(t)
    for s, last, anchor in plan.charts:
        ch = BarChart() if plan.rng.random() < 0.5 else LineChart()
        ch.add_data(Reference(sheets[s], min_col=2, min_row=1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(sheets[s], min_col=1, min_row=2, max_row=last))
        sheets[s].add_chart(ch, anchor)
    for s, kind, ref in plan.validations:
        dv = DataValidation(type="list", formula1='"Apple,Pear,Fig"') if kind == "list" else DataValidation(type="whole", operator="between", formula1="0", formula2="100")
        sheets[s].add_data_validation(dv)
        dv.add(ref)
    for s, kind, ref in plan.conds:
        fill = PatternFill("solid", fgColor="FFC7CE")
        top = ref.split(":")[0]
        rule = CellIsRule(operator="greaterThan", formula=["5"], fill=fill) if kind == "cell" else FormulaRule(formula=[f"${top[0]}{top[1:]}>10"], fill=fill)
        sheets[s].conditional_formatting.add(ref, rule)
    wb.save(path)


def write_xlsxwriter(plan, path):
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False})
    wb.set_properties({"created": FIXED_TIME})
    date = wb.add_format({"num_format": "yyyy-mm-dd"})
    sheets = {name: wb.add_worksheet(name) for name in plan.sheets}
    for (s, r, c), v in plan.cells.items():
        ws, cell = sheets[s], f"{col(c)}{r}"
        if v is None:
            continue
        if isinstance(v, str) and v.startswith("="):
            ws.write_formula(cell, v)
        elif isinstance(v, datetime.date):
            ws.write_datetime(cell, datetime.datetime(v.year, v.month, v.day), date)
        else:
            ws.write(cell, v)
    for s, ref, text in plan.legacy:
        ws = sheets[s]
        a = ref.split(":")
        ws.write_array_formula(f"{a[0]}:{a[-1]}", "{" + text + "}")
    for s, ref, text in plan.dynamic:
        a = ref.split(":")
        sheets[s].write_dynamic_array_formula(f"{a[0]}:{a[-1]}", text)
    for name, ref in plan.names:
        wb.define_name(name, "=" + ref)
    for s, name, ref in plan.tables:
        # The table's cells are the plan's, written above; add_table writes
        # the same headers again.
        cols = [{"header": h} for h in HEADERS + ["Total"]]
        sheets[s].add_table(ref, {"name": name, "columns": cols, "autofilter": True, "style": "Table Style Medium 2"})
    for s, last, anchor in plan.charts:
        ch = wb.add_chart({"type": plan.rng.choice(["column", "bar", "line"])})
        ch.add_series({"name": f"={s}!$B$1", "categories": f"={s}!$A$2:$A${last}", "values": f"={s}!$B$2:$B${last}"})
        sheets[s].insert_chart(anchor, ch)
    for s, kind, ref in plan.validations:
        if kind == "list":
            sheets[s].data_validation(ref, {"validate": "list", "source": ["Apple", "Pear", "Fig"]})
        else:
            sheets[s].data_validation(ref, {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 100})
    for s, kind, ref in plan.conds:
        red = wb.add_format({"bg_color": "#FFC7CE"})
        top = ref.split(":")[0]
        if kind == "cell":
            sheets[s].conditional_format(ref, {"type": "cell", "criteria": ">", "value": 5, "format": red})
        else:
            sheets[s].conditional_format(ref, {"type": "formula", "criteria": f"=${top[0]}{top[1:]}>10", "format": red})
    wb.close()


CELL = re.compile(r"<c ([^>]*?)(/>|>(.*?)</c>)", re.S)
FORMULA = re.compile(r"<f([^>]*)>(.*?)</f>|<f([^>]*)/>", re.S)
MARKER = re.compile(r'N\(&quot;([pld])&quot;\)|N\("([pld])"\)')


def kinds_ok(path):
    """Every formula in the file is stored as the kind its marker names.
    Returns the mismatches."""
    bad = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            xml = z.read(name).decode()
            for m in CELL.finditer(xml):
                attrs, body = m.group(1), m.group(3) or ""
                f = FORMULA.search(body)
                if not f:
                    continue
                fattrs = f.group(1) if f.group(1) is not None else f.group(3)
                text = f.group(2) or ""
                if not text:
                    continue  # a shared formula's follower
                mk = MARKER.search(text)
                if not mk:
                    bad.append(f"{name} {attrs}: no marker in {text}")
                    continue
                want = mk.group(1) or mk.group(2)
                array = 't="array"' in fattrs
                cm = re.search(r'\bcm="([0-9]+)"', attrs)
                got = "d" if array and cm and cm.group(1) != "0" else ("l" if array else "p")
                if got != want:
                    bad.append(f"{name} {attrs}: stored {got}, marked {want}: {text}")
    return bad


def rewrite(src, dst, part, fn):
    """Copy a package, passing one part's text through fn."""
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info.filename)
            if info.filename == part:
                new = fn(data.decode())
                if new == data.decode():
                    raise SystemExit(f"{dst}: the mutation changed nothing")
                data = new.encode()
            zo.writestr(info, data)


def refuse_cases(pool, out):
    """Pool files with one defect each, named <invariant>-<what>.xlsx."""
    os.makedirs(out, exist_ok=True)
    sheet = "xl/worksheets/sheet1.xml"

    def first(pred):
        for name in sorted(pool):
            if pred(name):
                return name
        raise SystemExit("no pool file for a refuse case")

    def has(text):
        def pred(name):
            with zipfile.ZipFile(name) as z:
                return text in z.read(sheet).decode()
        return pred

    plain = first(lambda n: True)
    # xml: a part that is not well-formed.
    rewrite(plain, f"{out}/xml-malformed.xlsx", sheet, lambda x: x.replace("</sheetData>", "</sheetDat>", 1))
    # xml: rows out of order (rows 2 and 3 swapped).
    def swap_rows(x):
        m2 = re.search(r'<row r="2"[^>]*>.*?</row>', x, re.S)
        m3 = re.search(r'<row r="3"[^>]*>.*?</row>', x, re.S)
        a, b = m2.group(0), m3.group(0)
        return x[: m2.start()] + b + a + x[m3.end():]
    rewrite(plain, f"{out}/xml-rows-order.xlsx", sheet, swap_rows)
    # xml: a cell twice in its row.
    def dup_cell(x):
        m = re.search(r'<c r="B2"[^>]*?(/>|>.*?</c>)', x, re.S)
        return x[: m.end()] + m.group(0) + x[m.end():]
    rewrite(plain, f"{out}/xml-cell-twice.xlsx", sheet, dup_cell)
    # kinds: a dynamic array stored without its cell metadata (plain now).
    dyn = first(has('cm="1"'))
    rewrite(dyn, f"{out}/kinds-dynamic-lost.xlsx", sheet, lambda x: re.sub(r' cm="1"', "", x, count=1))
    # kinds: a legacy array stored as a plain formula.
    leg = first(lambda n: re.search(r'<c r="[A-Z]+\d+"( s="\d+")?><f t="array" ref="[^"]+">[^<]*N\((&quot;|")l', zipfile.ZipFile(n).read(sheet).decode()) is not None)
    rewrite(leg, f"{out}/kinds-legacy-lost.xlsx", sheet, lambda x: re.sub(r'(<c r="[A-Z]+\d+"(?: s="\d+")?>)<f t="array" ref="[^"]+">([^<]*N\((?:&quot;|")l)', r"\1<f>\2", x, count=1))
    # kinds: a plain formula stored as a dynamic array.
    rewrite(dyn, f"{out}/kinds-plain-to-dynamic.xlsx", sheet, lambda x: re.sub(r'<c r="(G2)"((?: s="\d+")?)><f>', r'<c r="\1"\2 cm="1"><f t="array" ref="\1">', x, count=1))
    # editors: one value changed; compared against the file it came from
    # (named in the refuse case's own name after the double dash).
    base = os.path.basename(plain)[: -len(".xlsx")]
    rewrite(plain, f"{out}/editors-value--{base}.xlsx", sheet, lambda x: re.sub(r'(<c r="D2"[^>]*><v>)(\d+)', lambda m: m.group(1) + str(int(m.group(2)) + 1), x, count=1))
    rewrite(plain, f"{out}/editors-formula--{base}.xlsx", sheet, lambda x: x.replace("+N(&quot;p&quot;)", "+N(&quot;p&quot;)+1", 1) if "&quot;" in x else x.replace('+N("p")', '+N("p")+1', 1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--seed", type=int, default=20261118)
    ap.add_argument("--count", type=int, default=40)
    args = ap.parse_args()
    pool = os.path.join(args.out, "pool")
    os.makedirs(pool, exist_ok=True)
    made = []
    for i in range(args.count):
        rng = random.Random(args.seed * 1000 + i)
        producer = "openpyxl" if i % 2 == 0 else "xlsxwriter"
        plan = Plan(rng, dynamic=producer == "xlsxwriter")
        path = os.path.join(pool, f"{producer}-{i:02d}.xlsx")
        (write_openpyxl if producer == "openpyxl" else write_xlsxwriter)(plan, path)
        bad = kinds_ok(path)
        if bad:
            raise SystemExit(f"{path}: kinds stored unlike their markers:\n  " + "\n  ".join(bad))
        made.append(path)
    refuse_cases(made, os.path.join(args.out, "refuse"))
    total = sum(os.path.getsize(p) for p in made)
    print(f"{len(made)} workbooks, {total} bytes, in {pool}")


if __name__ == "__main__":
    main()
