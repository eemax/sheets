"""Helpers for task checks: read workbooks the way a grader needs to.

Values come from `sheets read --json`, which computes formulas a file
stores no result for, so a formula is judged by what it computes whoever
wrote it. Checks that must not depend on the tool under test (the claim
tasks, 19 and 20) read formula values with `lo_values` instead: a copy
recalculated by LibreOffice, as ssb.py's grading does. Stored results (what
pandas or a preview shows) are read separately with openpyxl, data_only.
"""
import atexit
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import openpyxl

SHEETS = os.environ.get("SHEETS_BIN", "sheets")


def values(path, ref):
    """Rows of computed values of a range, as {row: [values...]}."""
    out = subprocess.run([SHEETS, "read", path, ref, "--json", "-n", "100000"],
                         capture_output=True, text=True, env={**os.environ, "SHEETS_UNDO": "off"})
    if out.returncode != 0:
        raise RuntimeError(out.stderr.strip())
    data = json.loads(out.stdout)
    return {r["row"]: r["values"] for r in data["rows"]}, data["columns"]


def value(path, sheet, cell):
    """The computed value of one cell (None when empty, "?" when unknown)."""
    rows, _ = values(path, f"'{sheet}'!{cell}")
    for vals in rows.values():
        return vals[0] if vals else None
    return None


_LO = {"profile": None, "cache": {}, "dir": None}
_RECALC_ALWAYS = (
    '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
    'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
    '<value>0</value></prop></item>\n</oor:items>\n')


def _lo_recalculated(path):
    """A copy of PATH that LibreOffice (headless) recalculated, as an
    openpyxl workbook with data_only. One conversion per file state and
    process; the profile is the process's own, so concurrent checks do not
    share one."""
    st = os.stat(path)
    key = (os.path.realpath(path), st.st_mtime_ns, st.st_size)
    if key in _LO["cache"]:
        return _LO["cache"][key]
    if _LO["dir"] is None:
        _LO["dir"] = tempfile.mkdtemp(prefix="evalkit-lo-")
        atexit.register(shutil.rmtree, _LO["dir"], True)
        _LO["profile"] = os.path.join(_LO["dir"], "profile")
        xcu = os.path.join(_LO["profile"], "user", "registrymodifications.xcu")
        os.makedirs(os.path.dirname(xcu))
        with open(xcu, "w") as f:
            f.write(_RECALC_ALWAYS)
    src = os.path.join(_LO["dir"], f"in{len(_LO['cache'])}")
    dst = os.path.join(_LO["dir"], f"out{len(_LO['cache'])}")
    os.makedirs(src)
    shutil.copyfile(path, os.path.join(src, "book.xlsx"))
    soffice = os.environ.get("EVAL_SOFFICE") or shutil.which("soffice") or "/opt/homebrew/bin/soffice"
    cmd = [soffice, f"-env:UserInstallation=file://{_LO['profile']}", "--headless", "--calc", "--convert-to",
           "xlsx:Calc MS Excel 2007 XML", "--outdir", dst, os.path.join(src, "book.xlsx")]
    # A second try, for a conversion LibreOffice drops under load.
    for _ in range(2):
        p = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
        try:
            p.wait(timeout=180)
        except subprocess.TimeoutExpired:
            os.killpg(p.pid, 9)
            p.wait()
        if os.path.exists(os.path.join(dst, "book.xlsx")):
            break
    else:
        raise RuntimeError("LibreOffice could not convert " + os.path.basename(path))
    wb = openpyxl.load_workbook(os.path.join(dst, "book.xlsx"), data_only=True)
    _LO["cache"][key] = wb
    return wb


def lo_values(path, ref):
    """Rows of computed values of a range, as {row: [values...]} and the
    column letters, the way `values` gives them, but computed by LibreOffice
    on a copy of the file instead of by sheets. REF is `Sheet!A1:C5`,
    `'My sheet'!B2` or a sheet name alone (its used range, trimmed). Numbers come
    back as numbers, dates as datetimes, text as text, errors as their text
    (`#DIV/0!`) and empty cells as None. Use this where a check must not
    depend on the tool under test."""
    sheet, _, rng = ref.rpartition("!")
    if not sheet:
        sheet, rng = rng, ""
    sheet = sheet[1:-1].replace("''", "'") if sheet[:1] == "'" else sheet
    ws = _lo_recalculated(path)[sheet]
    if rng:
        cells = ws[rng]
    else:  # the used range, less the empty rows and columns a validation or a format leaves at its end
        cells = ws[ws.dimensions]
        filled = [c for row in cells for c in row if c.value is not None]
        if filled:
            last_row, last_col = max(c.row for c in filled), max(c.column for c in filled)
            cells = tuple(tuple(c for c in row if c.column <= last_col) for row in cells if row[0].row <= last_row)
    if not isinstance(cells, tuple):  # one cell
        cells = ((cells,),)
    rows = {r[0].row: [c.value for c in r] for r in cells}
    return rows, [c.column_letter for c in cells[0]]


def lo_value(path, sheet, cell):
    """The value of one cell as LibreOffice computes it (None when empty)."""
    rows, _ = lo_values(path, f"'{sheet}'!{cell}")
    for vals in rows.values():
        return vals[0]
    return None


def stored(path, sheet, cell):
    """The result a file stores for a cell, as pandas and previews see it."""
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb[sheet][cell].value


def formula(path, sheet, cell):
    wb = openpyxl.load_workbook(path)
    v = wb[sheet][cell].value
    return v if isinstance(v, str) and v.startswith("=") else None


def close(a, b, tol=0.005):
    try:
        return abs(float(a) - float(b)) <= tol * max(1, abs(float(b)))
    except (TypeError, ValueError):
        return False


_UNITS = {"k": 1e3, "thousand": 1e3, "m": 1e6, "mn": 1e6, "million": 1e6, "bn": 1e9, "billion": 1e9}


def numbers(text):
    """The numbers a reply mentions: 1,234.5, $1.2M, 338k, (1,200), -5%."""
    found = []
    pattern = r"([(\-\u2212])?\s?[$€£]?\s?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s*(million|thousand|billion|mn|bn|m|k)?\b"
    for m in re.finditer(pattern, text, re.I):
        n = float(m.group(2).replace(",", "") + (m.group(3) or ""))
        n *= _UNITS.get((m.group(4) or "").lower(), 1)
        found.append(-n if m.group(1) else n)
    return found


def mentions(text, value, rel=0.005):
    """Whether a reply states a number, to within rel (sign aside)."""
    return any(abs(abs(n) - abs(value)) <= rel * max(1, abs(value)) for n in numbers(text))


def excel_saved(path):
    """Make a file xlsxwriter wrote look as Excel saves it: results stored,
    no flag asking the next program to recalculate."""
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                data = data.replace(b' fullCalcOnLoad="1"', b"")
            z.writestr(info, data)


class Checks:
    """Collects named checks; report() prints the JSON the runner reads."""

    def __init__(self):
        self.items = []

    def add(self, name, ok, detail="", kind="task"):
        self.items.append({"name": name, "ok": bool(ok), "detail": str(detail)[:300], "kind": kind})
        return ok

    def guard(self, name, fn, kind="task"):
        try:
            ok, detail = fn()
        except Exception as e:  # a crash in a check fails it
            ok, detail = False, f"{type(e).__name__}: {e}"
        return self.add(name, ok, detail, kind)

    def report(self):
        task = [c for c in self.items if c["kind"] == "task"]
        print(json.dumps({"passed": all(c["ok"] for c in task), "checks": self.items}))


def args():
    """work dir, truth file, answer file (the agent's final reply)."""
    work, truth, answer = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(truth) as f:
        t = json.load(f)
    text = open(answer).read() if os.path.exists(answer) else ""
    return work, t, text
