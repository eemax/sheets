# Checking: lint and diff

    sheets lint book.xlsx                 # errors, broken patterns, totals that miss rows
    sheets lint book.xlsx --rules         # what it checks
    sheets diff book.xlsx                 # what the last sheets command changed
    sheets diff before.xlsx after.xlsx    # two versions of a workbook

- `lint` finds error values, `#REF!` from deleted cells, unknown functions,
  cells that break their column's formula pattern, rows a formula column
  misses, totals that leave out rows, sums that add a total together with
  the rows it totals (double counting), a line reading another line than
  its label names (Summary "Marketing" reading the Events row), numbers
  stored as text, references to Excel tables that do not exist, formulas
  over ranges stored for older Excel (which shows them with `@`), and
  functions stored without the prefix Excel needs (openpyxl leaves it
  off; Excel shows `#NAME?`). It exits 1 on errors
  (`--fail-on warning|never` to change that). Run it before handing a
  workbook over, and on any workbook a script wrote.
- `diff` aligns rows and columns first, so an inserted row is one added
  row, not every row below it changing. It lists sheets added, removed or
  renamed, then rows, columns and cells. With one file it compares with
  the state before the last change (`--since 3`: the last three), from the
  undo history: the way to check what a command did.
