# A departmental budget: 24 lines by month with annual totals, a total row,
# and a Summary that picks lines out. Planted mistakes: an annual total that
# misses December, a month total that misses the last line, a figure stored
# as text, a Summary line pointing at the wrong row, and a Summary total
# that sums the total row too. Results stored, as Excel saves them.
import json, random, sys
import xlsxwriter
from evalkit import excel_saved

work, truth = sys.argv[1], sys.argv[2]
rng = random.Random(16)
names = ["Payroll", "Benefits", "Contractors", "Rent", "Utilities", "Marketing", "Events", "Travel", "Software", "Hardware",
         "Cloud hosting", "Telecom", "Insurance", "Legal", "Audit", "Training", "Recruiting", "Office supplies", "Cleaning",
         "Security", "Maintenance", "Subscriptions", "Bank fees", "Miscellaneous"]
grid = []
for i, n in enumerate(names):
    base = rng.choice([400, 900, 1500, 2500, 4800, 7500, 12000, 41000]) if n != "Payroll" else 152000
    grid.append([round(base * rng.uniform(0.9, 1.1), -1) for _ in range(12)])
# A June figure typed as text ("4,850"), on a line whose months run in the
# thousands, so the figure itself is plausible.
text_row = next(i for i, row in enumerate(grid) if 2000 <= row[5] < 10000 and names[i] not in ("Payroll", "Software", "Marketing"))
text_cell = (text_row, 5)
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
cols = "BCDEFGHIJKLM"
first, last = 4, 3 + len(names)  # sheet rows of the lines
total_row = last + 1
path = f"{work}/budget.xlsx"
wb = xlsxwriter.Workbook(path)
bold = wb.add_format({"bold": True})
num = wb.add_format({"num_format": "#,##0"})
tot = wb.add_format({"bold": True, "num_format": "#,##0", "top": 1})
inputs = wb.add_format({"num_format": "#,##0", "font_color": "#0000FF"})
b = wb.add_worksheet("Budget")
b.write("A1", "Operating budget 2026", wb.add_format({"bold": True, "font_size": 14}))
b.write_row("A3", ["Line"] + months + ["Annual"], bold)
annual = []
for i, (n, row) in enumerate(zip(names, grid)):
    r = first + i
    b.write(r - 1, 0, n)
    for j, v in enumerate(row):
        if (i, j) == text_cell:
            b.write_string(r - 1, 1 + j, f"{v:,.0f}")
        else:
            b.write_number(r - 1, 1 + j, v, inputs)
    stored_row = [v for j, v in enumerate(row) if (i, j) != text_cell]
    if n == "Software":  # misses December
        b.write_formula(r - 1, 13, f"=SUM(B{r}:L{r})", num, sum(row[:11]))
    else:
        b.write_formula(r - 1, 13, f"=SUM(B{r}:M{r})", num, sum(stored_row))
    annual.append(sum(row))
b.write(total_row - 1, 0, "Total", bold)
col_totals = []
for j, col in enumerate(cols):
    end = last - 1 if col == "I" else last  # August misses the last line
    vals = [grid[i][j] for i in range(end - first + 1) if (i, j) != text_cell]
    b.write_formula(total_row - 1, 1 + j, f"=SUM({col}{first}:{col}{end})", tot, sum(vals))
    col_totals.append(sum(grid[i][j] for i in range(len(names))))
stored_annual = [sum(v for j, v in enumerate(row) if (i, j) != text_cell) if names[i] != "Software" else sum(row[:11]) for i, row in enumerate(grid)]
b.write_formula(total_row - 1, 13, f"=SUM(N{first}:N{last})", tot, sum(stored_annual))
b.set_column("A:A", 18)
b.set_column("B:N", 10)
b.freeze_panes(3, 1)

s = wb.add_worksheet("Summary")
s.write("A1", "Budget summary 2026", wb.add_format({"bold": True, "font_size": 14}))
picks = [("Payroll", "Payroll"), ("Marketing", "Marketing"), (names[text_row], names[text_row]), ("Software", "Software"), ("Cloud hosting", "Cloud hosting")]
for k, (label, n) in enumerate(picks, 3):
    i = names.index(n)
    r = first + i + (1 if label == "Marketing" else 0)  # Marketing points at Events
    s.write(k - 1, 0, label)
    s.write_formula(k - 1, 1, f"=Budget!N{r}", num, stored_annual[r - first])
k = 3 + len(picks)
s.write(k, 0, "Total operating budget", bold)
s.write_formula(k, 1, f"=SUM(Budget!N{first}:N{total_row})", tot, 2 * sum(stored_annual))  # counts the total row too
s.set_column("A:A", 26)
s.set_column("B:B", 14)
wb.close()
excel_saved(path)
json.dump({"names": names, "grid": grid, "text_line": names[text_row], "annual": annual, "col_totals": col_totals, "total": sum(annual),
           "first": first, "total_row": total_row, "summary_total_row": k + 1,
           "picks": {label: annual[names.index(n)] for label, n in picks}}, open(truth, "w"))
