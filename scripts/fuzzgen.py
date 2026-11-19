"""Generate the fuzzer's pool of random workbooks (internal/fuzz/testdata).

Half are written by openpyxl, half by XlsxWriter, each from its own seed:
values of every type, formulas of all three kinds (plain, legacy array
entered with Ctrl+Shift+Enter, dynamic array), Excel tables, defined names,
charts, validations and conditional formats. With --objects (long mode's
pools of runs from seed 2,000,000: fix round 3 of the fuzzer's long mode) it
adds merges, notes, links, hidden rows and columns and a styled header, which
the reach invariant judges outside an edit's reach. They come from a random
stream of their own, and only with the flag, so every seed without it makes
the files it always made, byte for byte: the pools the repros of known.txt
replay on need no change. --notes adds only the notes of --objects (the same
ones, for a seed), which the committed pool holds: the default mode judges a
note outside an edit's reach too (fix round 4: a note lost went unseen in it).

Every formula carries its kind in its text, as the condition N("p")=0,
N("l")=0 or N("d")=0 of an IF that is always true: IF(N("p")=0,B2*C2). N of
a text is 0, so the marker changes no result and, unlike an addition, no
result's type: a text, an empty text, a boolean and an error stay what they
are. The fuzzer checks after each edit that every formula is still stored as
the kind it names. After saving, this script reads each file back and
refuses one whose stored kinds do not match their markers: a producer that
stores a kind of its own choice (XlsxWriter makes any formula calling
SEQUENCE a dynamic array) would otherwise put a false finding into the pool.

Results of every type are in the pool. Formulas return numbers, texts, empty
texts (IF(B2>1000,1,"")), booleans and errors. XlsxWriter stores each
formula's result, as it does for a formula given one (text as t="str",
a boolean as t="b", an empty text as an empty <v>); openpyxl stores none,
which sheets reads as pending. About half the XlsxWriter files store an empty
text as Excel does (t="str" with an empty <v>) and half as XlsxWriter does
(an empty <v> with no type, which reads as no result). The filled-down column of about two files in
three is rewritten as shared formulas (<f t="shared" si="0" ref=...>, and
followers <f t="shared" si="0"/>), the way Excel and LibreOffice store it:
neither producer writes them.

It also writes the checkers' must-refuse cases (testdata/refuse/): pool
files with one defect each, which the fuzzer's invariants must catch.

    uv run --with openpyxl==3.1.5 --with xlsxwriter==3.3.1 python scripts/fuzzgen.py internal/fuzz/testdata

The versions are pinned: another release may write other bytes, which would
change every edit each seed makes and need known.txt made again (the script
refuses a release but these). The files are the same each time, byte for
byte: the dates in them (the document properties, and the time of each part
in the package) are fixed.

--seed changes the pool; --count the number of files (default 40).
--pool-only writes the pool and no must-refuse cases: the fuzzer's long mode
(SHEETS_FUZZ=long) makes a fresh pool this way from each seed it runs, and
the must-refuse cases are made from the committed pool's files (another seed
may hold none that fits a case). --objects adds the objects named above,
--notes the notes among them alone.
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
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

PINNED = {"openpyxl": "3.1.5", "xlsxwriter": "3.3.1"}
FIXED_TIME = datetime.datetime(2026, 11, 18, 12, 0, 0)
HEADERS = ["Item", "Units", "Price", "Cost", "Date", "Flag"]
ITEMS = ["Apple", "Pear", "Fig", "Kiwi", "Plum", "Lime", "Date", "Sloe", "N/A", "apple", "Acme Inc.", ""]
# Texts a writer must escape or quote: an ampersand, angle brackets, double
# quotes and an apostrophe, spaces at the ends (xml:space="preserve"), and
# non-ASCII. They are in the data of every file now and then (column A of
# Data, and the keys of the second sheets), and the fuzzer's edits type more
# (internal/fuzz/edits.go), so the streaming editor's own text writer and the
# full editor's both have text to escape (xlsx/stream_sheet.go newCell).
SPECIAL = ["R&D", "<b>x</b>", "Smith & Sons", 'say "hi"', "it's", " padded ", "Åäö 日本語"]
# Second sheets: the last two are names to escape in the workbook's parts
# (workbook.xml, the relationships, formulas that quote them).
SECOND = ["Calc", "Summary", "Q 2026", "Notes", "Model", "R&D", "Bob's Plan"]


def col(c):
    """Column letter of a 1-based column number."""
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def mark(text, kind):
    """A formula's text with its kind marker: IF(N("p")=0,text). The marker is
    the condition of an IF that is always true, so the result keeps its type."""
    return f'IF(N("{kind}")=0,{text})'


# Formulas whose results are not numbers, by the row they sit in: the text of
# the formula ({r} is the row; columns as on Data) and its result, from the
# cells the formula reads (v: the value of column letter at row r, None when
# empty). A producer that stores results gets them from here.
def _a(v, r):
    return v("A", r) or ""


def _b(v, r):
    return v("B", r) or 0


TEXTS = [
    ('A{r}&"-x"', lambda v, r: _a(v, r) + "-x"),
    ('IF(B{r}>3,"hi","lo")', lambda v, r: "hi" if _b(v, r) > 3 else "lo"),
    ("UPPER(A{r})", lambda v, r: _a(v, r).upper()),
    ('IF(A{r}="","none",A{r})', lambda v, r: _a(v, r) or "none"),
]
EMPTY_TEXTS = [  # empty for some rows or all
    ('IF(B{r}>1000,1,"")', lambda v, r: ""),
    ('IF(A{r}="Fig","x","")', lambda v, r: "x" if _a(v, r).lower() == "fig" else ""),
    ('IF(F{r},"yes","")', lambda v, r: "yes" if v("F", r) else ""),
]
BOOLS = [
    ("B{r}>3", lambda v, r: _b(v, r) > 3),
    ("AND(B{r}>3,C{r}<50)", lambda v, r: _b(v, r) > 3 and v("C", r) < 50),
    ("ISNUMBER(B{r})", lambda v, r: v("B", r) is not None),
    ("NOT(F{r})", lambda v, r: not v("F", r)),
    ("OR(F{r},D{r}>250)", lambda v, r: bool(v("F", r)) or v("D", r) > 250),
]


class Plan:
    """What one random workbook holds, independent of the producer: cells as
    (sheet, row, col) -> value, formula ("=...") or None, and its arrays,
    tables, names, charts, validations and conditional formats."""

    def __init__(self, rng, dynamic, objects=False, notes=False):
        self.rng = rng
        self.sheets = ["Data"]
        n2 = rng.choice([0, 1, 1, 2])
        for name in rng.sample(SECOND, n2):
            self.sheets.append(name)
        self.cells = {}  # (sheet, r, c) -> value
        self.legacy = []  # (sheet, ref, text): Ctrl+Shift+Enter arrays
        self.dynamic = []  # (sheet, ref, text): dynamic arrays
        self.tables = []  # (sheet, name, ref, headers)
        self.names = []  # (name, ref)
        self.charts = []  # (sheet, data rows, anchor)
        self.validations = []  # (sheet, kind, ref)
        self.conds = []  # (sheet, kind, ref)
        self.results = {}  # (sheet, "A1") -> the result of a formula that is no number
        self.shared = []  # (column letter, first row, last row): shared formula groups on Data
        self.typed_empty = False  # XlsxWriter: an empty text result stored as Excel does (t="str")
        # About half the files hold arrays over several cells (legacy or
        # dynamic). The others hold none, so the streaming editor takes edits
        # on them itself: it hands the full editor every set once a workbook
        # has a dynamic array, and every edit of rows once it has one or a
        # legacy array over several cells (xlsx/stream.go).
        self.arrays = rng.random() < 0.5
        self.n = rng.randint(4, 14)  # data rows on Data
        self.build(dynamic)
        self.extras(objects, notes)

    def build(self, dynamic):
        rng, n = self.rng, self.n
        last = n + 1
        for c, h in enumerate(HEADERS, 1):
            self.cells[("Data", 1, c)] = h
        for r in range(2, last + 1):
            self.cells[("Data", r, 1)] = rng.choice(ITEMS + SPECIAL) or None
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
        if self.arrays and rng.random() < 0.6:
            h = rng.randint(2, 3)
            self.legacy.append(("Data", f"K2:K{1 + h}", "=" + mark(f"B2:B{1 + h}*{rng.choice([2, 3])}", "l")))
        # Dynamic arrays (XlsxWriter only: openpyxl cannot write them).
        if dynamic and self.arrays:
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
                self.cells[(s, r, 1)] = rng.choice(ITEMS[:8] + SPECIAL[:3])
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
            if dynamic and self.arrays and rng.random() < 0.4:
                self.dynamic.append((s, "E2:E3", "=" + mark(f"{ref}!B2:B3*10", "d")))
        # A defined name, used by a formula.
        if rng.random() < 0.6:
            self.names.append(("Rate", "Data!$C$2"))
            self.cells[("Data", 2, 8)] = "=" + mark("B2*Rate", "p")
        if rng.random() < 0.4:
            self.names.append(("Units", f"Data!$B$2:$B${last}"))
            self.cells[("Data", 3, 8)] = "=" + mark("SUM(Units)", "p")
        # An Excel table, over the data (sheets does not move cells into,
        # out of or within one) or in a block of its own, away from what the
        # formulas read; or a chart, validation, conditional format.
        t = rng.random()
        if t < 0.25:
            self.tables.append(("Data", f"Sales{rng.randint(1, 99)}", f"A1:G{last}", HEADERS + ["Total"]))
        elif t < 0.5:
            headers = ["Aux1", "Aux2", "Aux3"]
            for c, h in enumerate(headers, 23):
                self.cells[("Data", 1, c)] = h
            for r in range(2, 6):
                for c in range(23, 26):
                    self.cells[("Data", r, c)] = rng.randint(0, 99)
            # (Not "Aux7": that reads as a cell, column AUX row 7, which a table's
            # name may not, and a command given it takes it for the cell.)
            self.tables.append(("Data", f"Block{rng.randint(1, 99)}", "W1:Y5", headers))
        if rng.random() < 0.6:
            self.charts.append(("Data", last, f"Q{rng.randint(2, 6)}"))
        if rng.random() < 0.6:
            self.validations.append(("Data", rng.choice(["list", "whole"]), f"A2:B{last}"))
        if rng.random() < 0.6:
            self.conds.append(("Data", rng.choice(["cell", "expr"]), f"B2:D{last}"))
        self.typed(last)
        # The filled-down column as shared formulas, in one group or two.
        if rng.random() < 0.65:
            if last - 1 >= 6 and rng.random() < 0.4:
                cut = rng.randint(3, last - 2)
                self.shared = [("G", 2, cut), ("G", cut + 1, last)]
            else:
                self.shared = [("G", 2, last)]
        self.typed_empty = rng.random() < 0.5

    def extras(self, objects, notes=False):
        """What a sheet holds besides cells and the objects above, with
        --objects: a merge, a note, a link (to the web, or to a cell), a hidden
        row and column, and a header in bold on a fill. The fuzzer's reach
        invariant judges these outside an edit's reach (internal/fuzz/objects.go).
        They come from a random stream of their own, and only with the flag, so
        the cells and objects above are the same for a seed as before they were
        added. With --notes alone, only the notes: the ones --objects makes for
        the seed (the draw before theirs is the merge's, taken and not used)."""
        self.merges = []  # (sheet, ref)
        self.notes = []  # (sheet, cell, text)
        self.links = []  # (sheet, cell, target): "https://..." or "Data!B2"
        self.hidden_cols = []  # (sheet, column letter)
        self.hidden_rows = []  # (sheet, row)
        self.styled = False
        if not objects and not notes:
            return
        o = random.Random(str(self.rng.getstate()[1][:8]) + "objects")
        if not objects:
            o.random()
            if o.random() < 0.25:
                self.notes.append(("Data", f"{o.choice('BCDEF')}1", "check this"))
            return
        last = self.n + 1
        if o.random() < 0.4:
            r = last + 5
            self.cells[("Data", r, 1)] = "Remarks"
            self.merges.append(("Data", f"A{r}:C{r}"))
        if o.random() < 0.25:
            self.notes.append(("Data", f"{o.choice('BCDEF')}1", "check this"))
        if o.random() < 0.4:
            r = last + 7
            self.cells[("Data", r, 1)] = "Source"
            self.links.append(("Data", f"A{r}", o.choice(["https://example.com/data", "Data!B2"])))
        if o.random() < 0.4:
            self.hidden_cols.append(("Data", "V"))
        if o.random() < 0.3:
            self.hidden_rows.append(("Data", last + 9))
        self.styled = o.random() < 0.5

    def value(self, sheet, c, r):
        """What the cell holds, for the results typed formulas compute."""
        v = self.cells.get((sheet, r, ord(c) - 64))
        return None if v == "" else v

    def typed(self, last):
        """Formulas whose results are texts, empty texts and booleans: filled
        down columns L (text) and P (boolean) of Data, a text array and a
        dynamic text array (T, U), and one on each second sheet."""
        rng = self.rng

        def v(c, r):
            return self.value("Data", c, r)

        def put(sheet, r, c, template, fn, kind):
            self.cells[(sheet, r, c)] = "=" + mark(template.replace("{r}", str(r)), kind)
            self.results[(sheet, f"{col(c)}{r}")] = fn(v, r)

        # Text and empty texts, mixed or all one: column L.
        template, fn = rng.choice(TEXTS + EMPTY_TEXTS)
        self.cells[("Data", 1, 12)] = "Label"
        for r in range(2, last + 1):
            put("Data", r, 12, template, fn, "p")
        # Booleans: column P.
        template, fn = rng.choice(BOOLS)
        self.cells[("Data", 1, 16)] = "Test"
        for r in range(2, last + 1):
            put("Data", r, 16, template, fn, "p")
        # A text in a legacy array (several cells), and in a dynamic one.
        h = rng.randint(2, 3)
        if self.arrays and rng.random() < 0.6:
            self.legacy.append(("Data", f"T2:T{1 + h}", "=" + mark(f"UPPER(A2:A{1 + h})", "l")))
            self.results[("Data", "T2")] = _a(v, 2).upper()
        if self.arrays and rng.random() < 0.6:
            self.dynamic.append(("Data", f"U2:U{1 + h}", "=" + mark(f'A2:A{1 + h}&"!"', "d")))
            self.results[("Data", "U2")] = _a(v, 2) + "!"
        # On each second sheet: a text or empty text, and a boolean, reading
        # Data.
        for s in self.sheets[1:]:
            ref = rng.choice(["Data", "'Data'"])
            if rng.random() < 0.7:
                self.cells[(s, 2, 6)] = "=" + mark(f'IF({ref}!B2>3,"hi","")', "p")
                self.results[(s, "F2")] = "hi" if (v("B", 2) or 0) > 3 else ""
            if rng.random() < 0.6:
                self.cells[(s, 3, 6)] = "=" + mark(f"{ref}!B2>3", "p")
                self.results[(s, "F3")] = (v("B", 2) or 0) > 3


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
    for s, name, ref, _ in plan.tables:
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
    for s, ref in plan.merges:
        sheets[s].merge_cells(ref)
    for s, ref, text in plan.notes:
        sheets[s][ref].comment = Comment(text, "fuzz")
    for s, ref, target in plan.links:
        if target.startswith("http"):
            sheets[s][ref].hyperlink = target
        else:
            sheets[s][ref].hyperlink = Hyperlink(ref=ref, location=target)
    for s, letter in plan.hidden_cols:
        sheets[s].column_dimensions[letter].hidden = True
    for s, r in plan.hidden_rows:
        sheets[s].row_dimensions[r].hidden = True
    if plan.styled:
        for c in range(1, 8):
            cell = sheets["Data"].cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
    wb.save(path)
    # openpyxl stamps the time of saving over the property set above.
    tmp = path + ".tmp"
    rewrite(path, tmp, "docProps/core.xml", lambda x: re.sub(r"(<dcterms:modified[^>]*>)[^<]*", r"\g<1>" + FIXED_TIME.isoformat() + "Z", x), strict=False)
    os.replace(tmp, path)


def write_xlsxwriter(plan, path):
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False})
    wb.set_properties({"created": FIXED_TIME})
    date = wb.add_format({"num_format": "yyyy-mm-dd"})
    head = wb.add_format({"bold": True, "bg_color": "#DDEBF7"}) if plan.styled else None
    sheets = {name: wb.add_worksheet(name) for name in plan.sheets}
    for (s, r, c), v in plan.cells.items():
        ws, cell = sheets[s], f"{col(c)}{r}"
        if v is None:
            continue
        if head is not None and s == "Data" and r == 1 and c <= 7:
            ws.write(cell, v, head)
        elif isinstance(v, str) and v.startswith("="):
            ws.write_formula(cell, v, None, plan.results.get((s, cell), 0))
        elif isinstance(v, datetime.date):
            ws.write_datetime(cell, datetime.datetime(v.year, v.month, v.day), date)
        else:
            ws.write(cell, v)
    for s, ref, text in plan.legacy:
        ws = sheets[s]
        a = ref.split(":")
        ws.write_array_formula(f"{a[0]}:{a[-1]}", "{" + text + "}", None, plan.results.get((s, a[0]), 0))
    for s, ref, text in plan.dynamic:
        a = ref.split(":")
        sheets[s].write_dynamic_array_formula(f"{a[0]}:{a[-1]}", text, None, plan.results.get((s, a[0]), 0))
    for name, ref in plan.names:
        wb.define_name(name, "=" + ref)
    for s, name, ref, headers in plan.tables:
        # The table's cells are the plan's, written above; add_table writes
        # the same headers again.
        cols = [{"header": h} for h in headers]
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
    for s, ref in plan.merges:
        sheets[s].merge_range(ref, plan.cells[(s, int(ref.split(":")[0][1:]), 1)])
    for s, ref, text in plan.notes:
        sheets[s].write_comment(ref, text)
    for s, ref, target in plan.links:
        sheets[s].write_url(ref, target if target.startswith("http") else "internal:" + target, string=plan.cells[(s, int(ref[1:]), 1)])
    for s, letter in plan.hidden_cols:
        sheets[s].set_column(f"{letter}:{letter}", None, None, {"hidden": True})
    for s, r in plan.hidden_rows:
        sheets[s].set_row(r - 1, None, None, {"hidden": True})
    wb.close()


def type_empty_texts(path):
    """XlsxWriter stores a formula's empty text result as an empty <v> with no
    type, which reads as no result. Excel and LibreOffice store it as t="str"
    with an empty <v>: this rewrites the cells that way, in place."""
    part = "xl/worksheets/sheet1.xml"

    def fn(xml):
        return re.sub(r'(<c r="[A-Z]+[0-9]+")(><f[^>]*>[^<]*</f><v></v></c>)', r'\1 t="str"\2', xml)

    tmp = path + ".tmp"
    rewrite(path, tmp, part, fn, strict=False)  # a file may have no empty text
    os.replace(tmp, path)


def share_formulas(path, groups):
    """Rewrite filled-down formulas of Data as shared formulas, in place: each
    group's first cell holds the formula text with the group's range, and the
    others only point to it (si), the way Excel and LibreOffice store a
    formula filled down. Neither producer writes them."""
    part = "xl/worksheets/sheet1.xml"  # Data is the first sheet

    def fn(xml):
        for si, (c, r1, r2) in enumerate(groups):
            for r in range(r1, r2 + 1):
                pat = re.compile(rf'(<c r="{c}{r}"[^>]*>)<f>(.*?)</f>', re.S)
                if r == r1:
                    sub = lambda m: f'{m.group(1)}<f t="shared" ref="{c}{r1}:{c}{r2}" si="{si}">{m.group(2)}</f>'
                else:
                    sub = lambda m: f'{m.group(1)}<f t="shared" si="{si}"/>'
                xml, n = pat.subn(sub, xml, count=1)
                if n != 1:
                    raise SystemExit(f"{path}: no plain formula at {c}{r} to share")
        return xml

    tmp = path + ".tmp"
    rewrite(path, tmp, part, fn)
    os.replace(tmp, path)


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


def fix_dates(path):
    """Give every part of a package the fixed time. openpyxl stamps each part
    with the time it saves at, which would make each run's files different."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            info.date_time = FIXED_TIME.timetuple()[:6]
            zo.writestr(info, zi.read(info.filename))
    os.replace(tmp, path)


def rewrite(src, dst, part, fn, strict=True):
    """Copy a package, passing one part's text through fn. A mutation that
    changes nothing is an error, unless strict is False."""
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info.filename)
            if info.filename == part:
                new = fn(data.decode())
                if strict and new == data.decode():
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
    # xml: well-formedness Go's decoder lets through but expat, and Excel,
    # reject: an attribute twice on one element, and a prefix no namespace
    # declaration binds.
    rewrite(plain, f"{out}/xml-attribute-twice.xlsx", sheet, lambda x: re.sub(r'<c r="([A-Z]+[0-9]+)"', r'<c r="\1" r="\1"', x, count=1))
    rewrite(plain, f"{out}/xml-prefix-unbound.xlsx", sheet, lambda x: x.replace("<sheetData>", "<sheetData><zz:x/>", 1))
    # xml: a text with spaces at its ends stored without xml:space="preserve"
    # (Excel drops the spaces): in the shared strings of a file that holds one.
    spaced = first(lambda n: 'xml:space="preserve"' in zipfile.ZipFile(n).read("xl/sharedStrings.xml").decode() if "xl/sharedStrings.xml" in zipfile.ZipFile(n).namelist() else False)
    rewrite(spaced, f"{out}/xml-space-lost--{os.path.basename(spaced)}", "xl/sharedStrings.xml", lambda x: x.replace(' xml:space="preserve"', "", 1))
    # kinds: a dynamic array stored without its cell metadata (plain now).
    # (Its file keeps G2 a plain formula: not one of a shared group.)
    dyn = first(lambda n: has('cm="1"')(n) and not has('t="shared"')(n))
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
    # editors: a stored text, boolean or empty text result that differs, each
    # in a file from XlsxWriter (it stores results) that holds one: the
    # results the editors must agree on as they do on numbers.
    def stores(pattern):
        def pred(n):
            return "xlsxwriter" in n and re.search(pattern, zipfile.ZipFile(n).read(sheet).decode()) is not None
        return pred

    def case(what, pattern, fn):
        n = first(stores(pattern))
        rewrite(n, f"{out}/editors-{what}--{os.path.basename(n)}", sheet, fn)

    case("text", r'<c r="[A-Z]+\d+" t="str"><f>[^<]*</f><v>[^<]+</v>',
         lambda x: re.sub(r'(<c r="[A-Z]+\d+" t="str"><f>[^<]*</f><v>)([^<]+)(</v>)', r"\1\2x\3", x, count=1))
    case("boolean", r'<c r="[A-Z]+\d+" t="b"><f>[^<]*</f><v>[01]</v>',
         lambda x: re.sub(r'(<c r="[A-Z]+\d+" t="b"><f>[^<]*</f><v>)([01])', lambda m: m.group(1) + str(1 - int(m.group(2))), x, count=1))
    case("empty-text-lost", r'<c r="[A-Z]+\d+" t="str"><f>[^<]*</f><v></v>',
         lambda x: re.sub(r'(<c r="[A-Z]+\d+") t="str">(<f>[^<]*</f>)<v></v>', r"\1>\2", x, count=1))
    rewrite(plain, f"{out}/editors-formula--{base}.xlsx", sheet, lambda x: x.replace("N(&quot;p&quot;)=0,", "N(&quot;p&quot;)=0,+", 1) if "&quot;" in x else x.replace('N("p")=0,', 'N("p")=0,+', 1))


def main():
    for mod in (openpyxl, xlsxwriter):
        have, want = mod.__version__, PINNED[mod.__name__]
        if have != want:
            raise SystemExit(f"{mod.__name__} {have}, not {want}: another release may write other bytes and change every edit the fuzzer makes (run with uv run --with openpyxl=={PINNED['openpyxl']} --with xlsxwriter=={PINNED['xlsxwriter']})")
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--seed", type=int, default=20261118)
    ap.add_argument("--count", type=int, default=40)
    ap.add_argument("--pool-only", action="store_true", help="no must-refuse cases (long mode's fresh pools)")
    ap.add_argument("--objects", action="store_true", help="merges, notes, links, hidden rows and columns, a styled header (long mode's pools of runs from seed 2,000,000)")
    ap.add_argument("--notes", action="store_true", help="the notes of --objects alone (the committed pool's)")
    args = ap.parse_args()
    pool = os.path.join(args.out, "pool")
    os.makedirs(pool, exist_ok=True)
    made = []
    for i in range(args.count):
        rng = random.Random(args.seed * 1000 + i)
        producer = "openpyxl" if i % 2 == 0 else "xlsxwriter"
        plan = Plan(rng, dynamic=producer == "xlsxwriter", objects=args.objects, notes=args.notes)
        path = os.path.join(pool, f"{producer}-{i:02d}.xlsx")
        (write_openpyxl if producer == "openpyxl" else write_xlsxwriter)(plan, path)
        if plan.shared:
            share_formulas(path, plan.shared)
        if producer == "xlsxwriter" and plan.typed_empty:
            type_empty_texts(path)
        bad = kinds_ok(path)
        if bad:
            raise SystemExit(f"{path}: kinds stored unlike their markers:\n  " + "\n  ".join(bad))
        fix_dates(path)
        made.append(path)
    if not args.pool_only:
        refuse_cases(made, os.path.join(args.out, "refuse"))
    total = sum(os.path.getsize(p) for p in made)
    print(f"{len(made)} workbooks, {total} bytes, in {pool}")


if __name__ == "__main__":
    main()
