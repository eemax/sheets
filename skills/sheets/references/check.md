# Checking: check, lint and diff

    sheets check book.xlsx                # formula results, unknowns, errors, and lint
    sheets check --fix book.xlsx          # store every result sheets can compute, add prefixes
    sheets check book.xlsx --against before.xlsx  # what changed since before.xlsx, and why
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
  Ctrl+Shift+Enter or spilling, gets every cell of its result; one in a
  single cell whose result is larger stays unknown (Excel shows its first
  value, or spills it if it was a dynamic array whose mark openpyxl
  dropped). In a workbook counting dates from 1904, dates count in its
  system; only a date written as text and read as a number stays unknown
  there. Running it again changes nothing; a file that needs nothing is
  not written. `--dry-run` shows the changes without saving. On a CSV
  file, `check` runs lint.
- On 150,000 rows (300,000 formulas openpyxl left without results),
  `check --fix` takes about 15 seconds, and a second run about 4.
- `check book.xlsx --against before.xlsx` says what a change did, whoever
  made it (a script, openpyxl, LibreOffice, Excel, sheets). Keep a copy
  before you change a workbook, then run it before handing the workbook
  over. It reports:
  - `changed`: inputs and formulas that changed. Flagged, as a likely
    mistake: one row of a formula column changed unlike the rows around it,
    or made a typed-in value; and a formula left as it was when rows or
    columns were inserted or deleted in what it reads (openpyxl's
    `insert_rows` and `delete_rows` rewrite no formula), or a sheet it reads
    was renamed (openpyxl's `ws.title`), with what Excel would have written;
  - `moved`: computed values that moved, each with its cause upstream (an
    input or formula change, rows added, removed or moved, a sheet or a
    name removed, a name changed), traced through what each formula reads.
    A formula changed is listed too, with what it computes now. A value
    that moved because rows added or removed moved its formula or what it
    reads (`=ROW()`, a range older Excel reads one cell of) has them as its
    cause;
  - `unexplained`: values that moved with nothing upstream changed, the
    sign of a broken formula or a wrong stored result;
  - `features_lost`: charts (on a chart sheet too), validations,
    conditional formats, shapes, pictures, names and tables the old
    version had (openpyxl drops shapes, a chart grouped with a shape, and
    validations in Excel 2010's extension list, such as a list read from
    another sheet);
  - `unknown`: values sheets cannot compare, with why (a function the
    calculator lacks, INDIRECT or OFFSET, a link to another workbook, a
    value from outside the workbook such as RTD, WEBSERVICE, STOCKHISTORY
    or the CUBE functions, or a formula reading one of them): never called
    unexplained;
  - `moved_rows`: rows moved as by a sort, which are not changes: each row
    whole, every column of its table (the columns between empty ones)
    together. A sort of a table with cells past an empty column that stay
    (a note, a list, a formula reading the table) is rows moved in its
    columns, as `Data!A2:D2, were A4:D4`, with `columns` in `--json`. A
    sort that left some of a table's columns behind (its labels, or a
    column next to it: Excel warns of data next to the selection), or two
    values swapped in a column, is inputs changed where they are, each the
    cause of what it moves;
  - then the check of book.xlsx, as above.

  Formulas Excel rewrites for rows or columns added or removed (a column
  inserted between the columns a formula reads too, or a row inserted on
  the sheet it reads), a sheet renamed, or a sheet deleted (`#REF!`), are
  not changes, nor is LibreOffice's `TRUE()` for `TRUE`. Dates compare as
  dates, whether a workbook counts them from 1900 or from 1904 (a Mac's).
  Rows a sort moved are moved rows, though rows were inserted or deleted
  above or inside them. Text lists ten of each section, the check's
  findings too, with counts, and cuts long values and formulas at 120
  characters; `--json` lists everything, whole (each moved value with
  `cell`, `before`, `after` and `cause`). It exits 1 when a value is
  unexplained, a feature is lost, or the check finds errors. On a CSV file
  it compares values only. On 150,000 rows it takes about 4 to 6 seconds
  (15,000 rows changed, or all of them sorted), and about 9 when openpyxl
  left one version without results.
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
