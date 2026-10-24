# Rows, columns and sheets

    sheets rows insert book.xlsx Sales!5 -n 3     # three empty rows before row 5
    sheets rows delete book.xlsx Sales!5:9
    sheets cols insert book.xlsx Sales!C
    sheets cols delete book.xlsx Sales!C:D
    sheets sheet add book.xlsx Notes --after Sales
    sheets sheet rename book.xlsx Sales 'Sales 2026'
    sheets sheet copy book.xlsx Sales Backup
    sheets sheet move book.xlsx Summary --first
    sheets sheet hide book.xlsx Backup            # and: show
    sheets sheet delete book.xlsx Old

- References follow the cells, on every sheet: formulas, names, merged
  cells, conditional formats, validations, filters and tables. A formula
  that pointed into deleted rows or columns becomes #REF!, as in Excel, and
  the output lists those formulas.
- A conditional format's or validation's formula is written for the first
  cell of its range (`=$D2=1` over A2:F50). Deleting that first row keeps
  the formula reading each cell's own row, as Excel does.
- Chart data ranges and Excel tables (Insert › Table) move with their
  cells too. Deleting an Excel table's header row or all of its data rows,
  and inserting or deleting columns inside it, are refused with the
  reason. Deleting its totals row leaves a table without one.
- Renaming a sheet updates every reference to it, including in charts.
- `sheet delete` refuses while formulas elsewhere still refer to the sheet,
  and lists them. `--force` deletes anyway; those references become #REF!.
- Excel table column names follow their header cells, as Excel requires:
  changing a header renames the column, and an emptied header becomes
  `Column3`. Formulas naming the old column (`Sales[Units]`, `[@Units]`)
  are rewritten to the new name, as Excel does.
- `write --clear` and `query --into` refuse a sheet holding an Excel
  table.
