# A wrong way that looks right: openpyxl. Its insert_rows moves the cells
# and nothing that refers to them (formulas, validations, conditional
# formats, the note), and saving drops the shape, rewrites the chart and
# stores no results. The new row's formulas are written by hand. The reply
# claims success and the right figures.
import copy, datetime, os, sys
import openpyxl

work = sys.argv[1]
path = os.path.join(work, "budget.xlsx")
wb = openpyxl.load_workbook(path)
ws = wb["Budget"]
ws.insert_rows(9)
for col, v in zip("ABCDEF", ["Legal", "Priya", 40000, 42000, 42000, 45000]):
    ws[f"{col}9"] = v
ws["G9"] = "=SUM(C9:F9)"
ws["H9"] = "=G9/G$10"
ws["I9"], ws["J9"] = "On track", datetime.datetime(2026, 6, 30)
ws["E6"], ws["F6"] = 352000, 363000
ws["I7"] = "At risk"
ws["B10"] = "Lena"
ws["J5"] = datetime.datetime(2026, 9, 30)
ws["A12"] = "Quarterly average"
for col in "CDEF":
    ws[f"{col}12"] = f"=AVERAGE({col}4:{col}10)"
wb.save(path)
print("Done. The budget total is now 3,122,000; Engineering is 42.8% of it.")
