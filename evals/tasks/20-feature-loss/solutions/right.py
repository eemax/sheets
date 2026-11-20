# The right way: with sheets, whose row insert moves the chart, the shape,
# the note, the validations and the conditional formats with the cells and
# widens the ranges that span the new row; and `copy` makes the new row look
# like its neighbour, formulas included.
import json, os, subprocess, sys

work = sys.argv[1]
book = os.path.join(work, "budget.xlsx")
SHEETS = os.environ.get("SHEETS_BIN", "sheets")


def S(*args):
    n = 2 if args[0] in ("sheet", "rows") else 1
    r = subprocess.run([SHEETS, *args[:n], book, *args[n:]], capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f"sheets {' '.join(args)} failed: {r.stderr.strip()}")
    return r.stdout


S("rows", "insert", "Budget!9")                                                  # 1: before HR
S("copy", "Budget!A8:J8", "Budget!A9")
S("set", "Budget!A9", "Legal", "Budget!B9", "Priya", "Budget!C9", "40000", "Budget!D9", "42000", "Budget!E9", "42000",
  "Budget!F9", "45000", "Budget!I9", "On track", "Budget!J9", "2026-06-30")
S("set", "Budget!E6", "352000", "Budget!F6", "363000")                          # 2
S("set", "Budget!I7", "At risk")                                                # 3
S("set", "Budget!B10", "Lena")                                                  # 4
S("set", "Budget!J5", "2026-09-30")                                             # 5
S("set", "Budget!A12", "Quarterly average", "Budget!C12:F12", "=AVERAGE(C4:C10)")  # 6
S("style", "Budget!C12:F12", "format=#,##0")


def cell(ref):
    return json.loads(S("read", ref, "--json"))["rows"][0]["values"][0]


print(f"Done. The budget total is now {cell('Budget!G11'):,.0f}; Engineering is {cell('Budget!H6'):.1%} of it.")
