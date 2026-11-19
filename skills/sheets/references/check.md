# Checking: check, lint and diff

    sheets check book.xlsx                # formula results, unknowns, errors, and lint
    sheets check --fix book.xlsx          # store every result sheets can compute, add prefixes
    sheets check book.xlsx --against before.xlsx  # what changed since before.xlsx, and why
    sheets check book.xlsx --against before.xlsx --receipt receipt.md  # the same, for a person
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
    `insert_rows` and `delete_rows` rewrite no formula), a sheet it reads
    was renamed (openpyxl's `ws.title`), or the cells it reads were cut and
    pasted elsewhere (openpyxl's `move_range`), with what Excel would have
    written. A defined name left as it was in the same way is flagged in
    `structure` (`name not rewritten`, with what Excel writes), and the
    values it moved name it as their cause. Inputs compare exactly: a
    barcode, an order number or an amount in the millions changed in its
    last digit is a change, in a CSV file too; computed results compare to
    Excel's 15 digits;
  - `moved`: computed values that moved, each with its cause upstream (an
    input or formula change, rows added, removed or moved, a sheet or a
    name removed, a name changed), traced through what each formula reads.
    A value that several changed cells reach names each of them, the first
    five and then `and more` (`input Data!B2, input Data!B3`: a MAX that
    one of them moved is the work of the one that moved it, but both are
    among its sources). An input that changed is always named, whatever
    rows were added, removed or moved around it: a structure change that
    reaches the value comes first, the inputs after it. Rows a sort moved
    are a cause only of what reads them by position (a lookup, `=Data!B2`,
    a sum of some of them): a total over all of them (SUM, MAX, COUNT, an
    average, SUMIF or SUMPRODUCT over the same rows, a table's column)
    did not move for the sort, and names the input that moved it. Rows
    removed or added are a cause only of what they can move that way, a
    formula Excel rewrote for them included, over a range, a table's
    column or a defined name alike: a MAX that rose for a row removed, a
    COUNT that rose, a SUM that fell though they held 0 or less (or by
    more than they held), or a total that moved for an empty row inserted
    in its range, did not move for them, and stays `unexplained` unless
    something else moved it. A formula changed is listed too, with what
    it computes now. A value that moved because rows added or removed moved
    its formula or what it reads (`=ROW()`, a range older Excel reads one
    cell of) has them as its cause;
  - `unexplained`: values that moved with nothing upstream changed, the
    sign of a broken formula or a wrong stored result (a total whose rows
    were sorted too: no sort moves it). What it moved is in `moved` with
    it as its cause; the cells that read it and hold what they held are
    listed with it (`held` in `--json`), not as unknown;
  - `features_lost`: charts (on a chart sheet too), validations,
    conditional formats, shapes, pictures, names and tables the old
    version had (openpyxl drops shapes, a chart grouped with a shape, and
    validations in Excel 2010's extension list, such as a list read from
    another sheet). A feature is kept only by one of its kind and identity
    in the new version: a validation or conditional format of the same
    type and formula (the cells it reads aside, which rows and columns
    added or removed rewrite), a chart of the same type in the same place
    or on the same data, a shape or picture in the same place or by name,
    a name or table by its name. Another of its kind is no stand-in: a
    chart added beside a dropped one, a validation beside one dropped, and
    the banding `sheets fmt` writes over a conditional format, are the old
    one lost. A validation or conditional format whose cells hold another
    of its kind now (a bound raised, a dropdown's items, a threshold) was
    edited: it is in `features_changed`, with its rule before and after,
    not lost, and sets no exit 1. Only a rule of the same type is an edit
    of it (a whole-number validation made a decimal one, or one that
    allows any value, too), and a formula rule or custom validation only
    when it is the same formula but for its numbers, texts and
    comparison: another rule on the same cells (a list where a number was,
    data bars for a highlight, banding for a formula highlight) is the old
    one lost. A column deleted takes the
    validations on it, and the ones beside it move into its place. An
    Excel table renamed (the same columns, on the same sheet) is followed,
    not lost: the formulas that read it by name, which Excel rewrites, are
    no changes (`structure` says `table renamed`);
  - `unknown`: values sheets cannot compare, with why (a function the
    calculator lacks, INDIRECT or OFFSET, a link to another workbook, a
    value from outside the workbook such as RTD, WEBSERVICE, STOCKHISTORY
    or the CUBE functions, or a formula reading one of them): never called
    unexplained. Also a value that moved only because one version stores
    Excel's result and sheets computed the other's, where the sheets
    calculator disagrees with Excel (the difference is the calculator's),
    and a stored result that held though what it reads changed, which the
    calculator gets otherwise now: likely out of date (written without
    recalculating, or Excel calculating by hand);
  - `moved_rows`: rows moved as by a sort, which are not changes: each row
    whole, every column of its table (the columns between empty ones)
    together. A sort of a table with cells past an empty column that stay
    (a note, a list, a formula reading the table) is rows moved in its
    columns, as `Data!A2:D2, were A4:D4`, with `columns` in `--json`. A
    sort that left some of a table's columns behind (its labels, or a
    column next to it: Excel warns of data next to the selection), or two
    values swapped in a column, is inputs changed where they are, each the
    cause of what it moves. Rows cut and pasted where nothing was
    (`sheets move`, Excel's cut and paste) are rows moved too, with the rows
    they passed where they were, not rows added or removed. So is a block of
    cells cut from beside a table and pasted below it or into other columns
    (`sheets move book.xlsx Sales!H1:J20 Sales!A30`, listed as
    `Sales!A30:C49, were H1:J20`, with `columns` and `was_columns` in
    `--json`): cells that left one rectangle and filled another of its size,
    holding what they held, formulas as Excel rewrites them, while the
    cells around stayed (a few others may have changed, but not in the
    rows it left or went to, nor in its columns when it moved straight
    down or up: that is not told from a sort of some columns, and is
    changes). Moved straight down or up with nothing beside it, that is
    as well rows added or removed above, and is told as that;
  - then the check of book.xlsx, as above.

  A row deleted with a cell set in the row that moved up into its place is
  one row removed and one input changed: a row changed pairs with the row
  removed or added beside it that differs from it in the fewest cells.
  Formulas Excel rewrites for rows or columns added or removed (a column
  inserted between the columns a formula reads too, or a row inserted on
  the sheet it reads), cells cut and pasted, a sheet renamed, or a sheet
  deleted (`#REF!`), are not changes, nor is LibreOffice's `TRUE()` for
  `TRUE`. An array
  formula's result (a FILTER, SORT or SEQUENCE spilling) is where its
  formula puts it: one that lost a row or was sorted again moved its
  values in place, each with its cause, and rows inserted or deleted
  through it are those rows, its values as they were. Dates compare as
  dates, whether a workbook counts them from 1900 or from 1904 (a Mac's).
  Rows a sort moved are moved rows, though rows were inserted or deleted
  above or inside them. Text lists ten of each section, the check's
  findings too, with counts, and cuts long values and formulas at 120
  characters; `--json` lists everything, whole (each moved value with
  `cell`, `before`, `after` and `cause`). It exits 1 when a value is
  unexplained, a feature is lost, or the check finds errors. On a CSV file
  it compares values only. On 150,000 rows it takes about 4 to 6 seconds
  (15,000 rows changed, or all of them sorted), and about 9 when openpyxl
  left one version without results. It compares the cells there are, not
  the used range's rectangle: a sheet with A1:B1 and one value at
  XFD1048576 takes under a second.
- `--receipt receipt.md` (with `--against`) writes that report for a
  person, as Markdown to attach to a pull request or a message: what to
  look at first (values unexplained, features lost, formula changes
  flagged, the check's errors), a table of the counts, each section's
  first 25 entries (and the command that lists them all), the check of
  the file, then before and after pictures, drawn by `sheets render`, of
  up to three ranges that changed (each sheet before a second range of
  one, at most 30 rows by 12 columns each, the cells that changed named
  under each), written beside it as `receipt-1-before.png`,
  `receipt-1-after.png` and so on (pictures an earlier receipt left under
  those names are removed). A chart lost is pictured where it was. What no
  picture can show it says in words: a validation, a shape or picture, a
  conditional format, a table style, a name or a chart on a chart sheet
  lost; what `render` leaves out of a range; values sheets cannot know.
  The pictures show the results sheets computed where a file stores none,
  as the tables do. Each command's output ends with
  `receipt: receipt.md (4 pictures of 2 ranges)`; `--json` gives
  `receipt` with each range's `before`, `after` and their images. On a CSV
  file there are no pictures. On 150,000 rows it takes about 8 seconds.
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
- `diff` is git's diff driver for workbooks, so `git diff` shows what
  changed in a `.xlsx` file, not `Binary files differ`. Set it up once in
  a repository:

      git config diff.sheets.command 'sheets diff'
      echo '*.xlsx diff=sheets' >> .gitattributes

  `git diff` and `git diff --cached` use it; `git show` and `git log -p`
  with `--ext-diff`. Each file starts with its line,
  `sheets diff a/book.xlsx b/book.xlsx` (`new file`, `deleted` or
  `renamed` after it), then is as `sheets diff` prints two files.
