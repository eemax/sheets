# Checking: check, lint and diff

    sheets check book.xlsx                # formula results, unknowns, errors, and lint
    sheets check --fix book.xlsx          # store every result sheets can compute, add prefixes
    sheets lint book.xlsx                 # errors, broken patterns, totals that miss rows
    sheets lint book.xlsx --rules         # what it checks
    sheets diff book.xlsx                 # what the last sheets command changed
    sheets diff before.xlsx after.xlsx    # two versions of a workbook

- Finish with `sheets check --fix` on any workbook a script wrote
  (openpyxl, pandas, XlsxWriter), then hand it over. Python's files have
  silent defects: openpyxl stores no formula results, so pandas, previews
  and other programs read blanks; XlsxWriter stores 0 for every formula;
  and both store functions added after Excel 2007 (XLOOKUP, TEXTJOIN,
  MAXIFS, ...) without the `_xlfn.` prefix Excel needs, so Excel shows
  `#NAME?`.
- `check` writes nothing. It counts formulas by their stored results
  (vouched for, missing, or stale: the file asks Excel to recalculate on
  open), and reports:
  - `missing-result` and `stale-result` (errors): results `--fix` would
    store, with an example;
  - `unknown-result` (warning): formulas whose results sheets cannot know,
    grouped by why (a function the calculator lacks, a link to another
    workbook, a circular reference, a formula they read); Excel computes
    them on open;
  - `calc-differs` (note): results Excel stored that the sheets
    calculator does not agree with; Excel's are kept;
  - everything `lint` finds, formula errors and missing prefixes included.

  It exits 1 on errors, as `lint` does (`--fail-on warning|note|never`).
- `check --fix` stores every result sheets can compute and adds missing
  prefixes, as one change (`undo` takes it back), and lists what it
  changed, then what is left. A result Excel stored is never replaced, and
  nothing but results and prefixes changes; a stale result sheets cannot
  recompute (a link to another workbook among them) is removed rather
  than vouched for. An array formula over several cells, entered with
  Ctrl+Shift+Enter or spilling, gets every cell of its result. In a
  workbook counting dates from 1904, results that work with dates stay
  unknown. Running it again changes nothing; a file that needs nothing is
  not written. `--dry-run` shows the changes without saving. On a CSV
  file, `check` runs lint.
- On 150,000 rows (300,000 formulas openpyxl left without results),
  `check --fix` takes about 15 seconds, and a second run about 4.
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
