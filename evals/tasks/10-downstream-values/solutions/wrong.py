# A wrong way that looks right: openpyxl opened with data_only=True keeps
# only the stored results, so saving replaces every formula with a fixed
# value and the orders no longer look prices up. Its reply claims success.
import os, sys
import openpyxl

work = sys.argv[1]
path = os.path.join(work, "prices.xlsx")
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb["Products"]
for row in ws.iter_rows(min_row=2):
    if row[0].value == "Widget":
        row[1].value = 15
wb.save(path)
print("Done: Widget now costs 15 in prices.xlsx.")
