# Checking: lint and diff

    sheets lint book.xlsx                 # errors, broken patterns, totals that miss rows
    sheets lint book.xlsx --rules         # what it checks
    sheets diff book.xlsx                 # what the last sheets command changed
    sheets diff before.xlsx after.xlsx    # two versions of a workbook

- `lint` finds error values, `#REF!` from deleted cells, unknown functions,
  cells that break their column's formula pattern, totals that leave out
  rows, numbers stored as text, and more. It exits 1 on errors
  (`--fail-on warning|never` to change that). Run it before handing a
  workbook over.
- `diff` aligns rows and columns first, so an inserted row is one added
  row, not every row below it changing. It lists sheets added, removed or
  renamed, then rows, columns and cells. With one file it compares with
  the state before the last change (`--since 3`: the last three), from the
  undo history: the way to check what a command did.
