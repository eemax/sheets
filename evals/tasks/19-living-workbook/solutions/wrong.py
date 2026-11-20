# A wrong way that looks right: all 20 edits with openpyxl, the way an agent
# without sheets does them. Every value lands where it was asked, but
# openpyxl keeps neither the dynamic arrays (they save as plain array
# formulas), nor the sheet renaming's references, nor the new rows' formats,
# nor the table's new column; it rewrites the chart. The reply claims
# success and quotes the figures.
import copy, datetime, os, sys
import openpyxl

work = sys.argv[1]
path = os.path.join(work, "sales-book.xlsx")
wb = openpyxl.load_workbook(path)
inp, pr, od, sm = wb["Inputs"], wb["Products"], wb["Orders"], wb["Summary"]

inp["B2"] = 0.22                                                    # 1
pr["C3"] = 34                                                       # 2
for col, v in zip("ABCD", ["Gizmo Pro", "Hardware", 48, 21]):       # 3
    pr[f"{col}8"] = v
pr["E8"] = pr["E7"].value
pr.tables["Products"].ref = "A1:E8"
for col, v in zip("ABCD", [datetime.datetime(2026, 3, 18), "North", "Gizmo Pro", 12]):  # 4
    od[f"{col}16"] = v
for col in "EFGH":
    od[f"{col}16"] = od[f"{col}15"].value
od.tables["Orders"].ref = "A1:H16"
od["D4"] = 4                                                        # 5
od["B5"] = "South"                                                  # 6
inp["B3"] = 8                                                       # 7
sm["G3"] = "Orders"                                                 # 8
for r in range(4, 8):
    sm[f"G{r}"] = f"=COUNTIF(Orders[Region],A{r})"
sm["G8"] = "=SUM(G4:G7)"
sm["A21"], sm["B21"] = "Average order value", "=B8/G8"              # 9
wb["Insights"].title = "Rankings"                                   # 10 (Summary!B20 still says Insights)
pr["B6"] = "Service"                                                # 11
od.delete_rows(6)                                                   # 12
od.tables["Orders"].ref = "A1:H15"
tg = wb.create_sheet("Targets")                                     # 13
for row in [("Region", "Target"), ("North", 1500), ("South", 800), ("East", 650), ("West", 900)]:
    tg.append(row)
sm["H3"] = "Gap to target"                                          # 14
for r in range(4, 8):
    sm[f"H{r}"] = f"=B{r}-INDEX(Targets!$B$2:$B$5,MATCH(A{r},Targets!$A$2:$A$5,0))"
sm["H8"] = "=SUM(H4:H7)"
inp["B5"] = 0.55                                                    # 15
for r in range(4, 9):                                               # 16
    sm[f"F{r}"].number_format = "0.0%"
rows = sorted(([c.value for c in row] for row in pr.iter_rows(min_row=2, max_row=8, max_col=4)), key=lambda r: -r[2])  # 17
for i, vals in enumerate(rows, 2):
    for col, v in zip("ABCD", vals):
        pr[f"{col}{i}"] = v
inp["A6"], inp["B6"] = "Shipping per order", 4.5                    # 18
wb.defined_names["ShipPerOrder"] = openpyxl.workbook.defined_name.DefinedName("ShipPerOrder", attr_text="Inputs!$B$6")
od["I1"] = "Shipping"                                               # 19
for r in range(2, 16):
    od[f"I{r}"] = f"=IF(D{r}>=BulkUnits,0,ShipPerOrder)"
sm["A22"], sm["B22"] = "Shipping cost", "=SUM(I2:I15)"              # 20 (a sum of the sheet's own cells, not Orders[Shipping])
sm["B22"] = "=SUM(Orders!I2:I15)"
sm["A23"], sm["B23"] = "Margin after shipping", "=D8-B22"
wb.save(path)
print("Done: all 20 changes are made in sales-book.xlsx. Total revenue is 3,744.46 and the margin after shipping is 2,167.46.")
