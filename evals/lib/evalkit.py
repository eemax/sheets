"""Helpers for task checks: read workbooks the way a grader needs to.

Values come from `sheets read --json`, which computes formulas a file
stores no result for, so a formula is judged by what it computes whoever
wrote it. Stored results (what pandas or a preview shows) are read
separately with openpyxl, data_only.
"""
import json
import os
import subprocess
import sys

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
