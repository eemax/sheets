# The 20 edits, done the way that keeps the workbook whole: with the sheets
# CLI (it keeps tables, names, the chart, conditional formats and
# validations, and moves references with the cells), and with direct edits
# of the package where it has no command or does not yet keep a part:
# a defined name, one more column of an Excel table, and the spill ranges of
# the dynamic arrays (it saves each as a one-cell array that Excel would
# show as #SPILL!).
import json, os, re, shutil, subprocess, sys, zipfile

work = sys.argv[1]
book = os.path.join(work, "sales-book.xlsx")
SHEETS = os.environ.get("SHEETS_BIN", "sheets")


def S(*args, stdin=None):
    """sheets COMMAND FILE ARGS; `sheet` and `rows` take a subcommand first."""
    n = 2 if args[0] in ("sheet", "rows") else 1
    r = subprocess.run([SHEETS, *args[:n], book, *args[n:]], input=stdin, capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f"sheets {' '.join(args)} failed: {r.stderr.strip()}")
    return r.stdout


def patch(edit):
    """Rewrite the package: edit(name, bytes) -> bytes for each part."""
    tmp = book + ".tmp"
    with zipfile.ZipFile(book) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, edit(item.filename, zin.read(item.filename)))
    shutil.move(tmp, book)


S("set", "Inputs!B2", "0.22")                                              # 1
S("set", "Products!C3", "34")                                              # 2
S("write", "Products", "--append", stdin="Gizmo Pro\tHardware\t48\t21\n")  # 3
S("write", "Orders", "--append", stdin="2026-03-18\tNorth\tGizmo Pro\t12\n")  # 4
S("set", "Orders!D4", "4")                                                 # 5
S("set", "Orders!B5", "South")                                             # 6
S("set", "Inputs!B3", "8")                                                 # 7
S("set", "Summary!G3", "Orders", "Summary!G4:G7", "=COUNTIFS(Orders[Region],A4)", "Summary!G8", "=SUM(G4:G7)")  # 8
S("set", "Summary!A21", "Average order value", "Summary!B21", "=B8/G8")    # 9
S("sheet", "rename", "Insights", "Rankings")                               # 10
S("set", "Products!B6", "Service")                                         # 11
S("rows", "delete", "Orders!6")                                            # 12
S("sheet", "add", "Targets")                                               # 13
S("write", "Targets!A1", stdin="Region\tTarget\nNorth\t1500\nSouth\t800\nEast\t650\nWest\t900\n")
S("set", "Summary!H3", "Gap to target", "Summary!H4:H7", "=B4-INDEX(Targets!$B$2:$B$5,MATCH(A4,Targets!$A$2:$A$5,0))",
  "Summary!H8", "=SUM(H4:H7)")                                             # 14
S("set", "Inputs!B5", "0.55")                                              # 15
S("style", "Summary!F4:F8", "format=0.0%")                                 # 16
S("sort", "Products", "Price:desc")                                        # 17
S("set", "Inputs!A6", "Shipping per order", "Inputs!B6", "4.5")            # 18
S("set", "Orders!I1", "Shipping")                                          # 19: the column joins the table, with its formula
LAST = 15  # 14 orders


def join_table(name, data):
    if name.startswith("xl/tables/") and b'name="Orders"' in data:
        data = data.replace(b'ref="A1:H15"', b'ref="A1:I15"')
        data = re.sub(rb'<tableColumns count="8">(.*)</tableColumns>', lambda m: b'<tableColumns count="9">' + m.group(1) + (
            b'<tableColumn id="9" name="Shipping"><calculatedColumnFormula>IF(Orders[[#This Row],[Units]]&gt;=BulkUnits,0,ShipPerOrder)'
            b'</calculatedColumnFormula></tableColumn></tableColumns>'), data, flags=re.S)
    elif name == "xl/workbook.xml":  # 18: the name
        data = data.replace(b"</definedNames>", b'<definedName name="ShipPerOrder">Inputs!$B$6</definedName></definedNames>')
    return data


patch(join_table)
S("set", f"Orders!I2:I{LAST}", "=IF(Orders[[#This Row],[Units]]>=BulkUnits,0,ShipPerOrder)")
S("set", "Summary!A22", "Shipping cost", "Summary!B22", "=SUM(Orders[Shipping])",
  "Summary!A23", "Margin after shipping", "Summary!B23", "=D8-B22")        # 20
# the new rows and the new column look like their neighbours
S("style", f"Orders!A2:A{LAST}", "format=yyyy-mm-dd")
S("style", f"Orders!F2:I{LAST}", "format=#,##0.00")
S("style", "Products!C2:D8", "format=#,##0.00")
S("style", "Products!E2:E8", "format=0%")
S("style", "Summary!G4:G8", "format=0")
S("style", "Summary!B21:B23", "format=#,##0.00")


def spills(name, data):
    if name.startswith("xl/worksheets/sheet5.xml"):
        text = data.decode()
        for anchor, ref in [("A2", "A2:A5"), ("B2", "B2:B5"), ("D2", "D2:D4"), ("F2", "F2:F4")]:
            text = re.sub(rf'(<c r="{anchor}"[^>]*?) t="e"([^>]*><f t="array" ref=")[^"]*("[^>]*>[^<]*</f>)<v>[^<]*</v>',
                          rf"\1\2{ref}\3", text)
            text = re.sub(rf'(<c r="{anchor}"[^>]*><f t="array" ref=")[^"]*"', rf'\g<1>{ref}"', text)
        data = text.encode()
    return data


patch(spills)


def cell(ref):
    return json.loads(S("read", ref, "--json"))["rows"][0]["values"][0]


print(f"All 20 changes are in sales-book.xlsx. Total revenue is {cell('Summary!B8'):,.2f}; "
      f"margin after shipping is {cell('Summary!B23'):,.2f}.")
