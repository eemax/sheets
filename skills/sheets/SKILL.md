---
name: sheets
description: Use this skill whenever a task involves an Excel workbook (.xlsx, .xlsm) or a CSV or TSV file - finding out what a spreadsheet holds, answering questions from its data, adding or changing rows, values or formulas, inserting or deleting rows, columns or sheets, sorting, copying, moving or find-and-replace, notes and hyperlinks, charts, making a sheet look professional or styling it as asked (number formats, highlights, dropdowns, frozen headers), converting between CSV and Excel, or checking a workbook for errors before handing it over. Use it even when the user only names the file or calls it the report or the budget, and use it instead of writing openpyxl or pandas code. It runs the sheets command-line tool, whose output is sized for the context window. Not for Google Sheets, Apple Numbers or legacy .xls files.
license: MIT
compatibility: Needs the sheets command-line tool on the PATH (go install github.com/eemax/sheets/cmd/sheets@latest). Works on local files, without network access.
metadata:
  author: eemax
  repository: https://github.com/eemax/sheets
---

# sheets

`sheets` reads, writes and formats Excel workbooks (.xlsx, .xlsm), and
reads and edits CSV and TSV files. Use it instead of writing openpyxl or
pandas code: its output is compact, it keeps what a workbook already holds
(formulas, formatting, charts, Excel tables, notes), it stores current
formula results when it saves, and every change can be undone. If the
`sheets` command is missing, install it with
`go install github.com/eemax/sheets/cmd/sheets@latest`.

## Look first

    sheets info book.xlsx                  # sheets, used ranges, tables, charts, names
    sheets peek book.xlsx Sales            # one line per column: type, range, values, formulas
    sheets read book.xlsx 'Sales!A1:F20'   # cells; --formulas shows formulas
    sheets find book.xlsx acme             # matching cells, with their column header
    sheets query book.xlsx 'SELECT Region, SUM(Revenue) FROM Sales GROUP BY Region'

Start with `info`, then `peek` the sheet, then `read` only the cells you
need: `peek` is the same size for ten rows or a million. To total, filter
or join, use `query` (SQLite over every table) rather than reading rows.
Each output ends with a `next:` line suggesting what to run next.

## Change

    printf 'Date\tAmount\n2026-07-01\t$1,200\n' | sheets write book.xlsx 'Sales!A1'
    sheets write book.xlsx Sales --append -i rows.json   # into the last table, above its total row
    sheets set book.xlsx Sales!F1 Margin                 # a new column's header,
    sheets set book.xlsx 'Sales[Margin]' '=E{r}/D{r}'    # then its data rows; {r} is the row
    sheets rows delete book.xlsx 'Sales!5:9'             # also: rows insert, cols, sheet
    sheets sort book.xlsx Sales Revenue:desc             # also: copy, move, clear, replace
    sheets fmt book.xlsx                                 # the house style
    sheets style book.xlsx 'Sales[Revenue]' format=currency decimals=0
    sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red

Put more than two or three changes in one batch. It is saved once, as
one undo step, and nothing is saved if a line fails:

    sheets do book.xlsx <<'EOF'
    write Sales!A1 -i sales.csv
    set Sales!F1 Margin
    set 'Sales[Margin]' '=E{r}/D{r}'
    fmt
    cond 'Sales[Margin]' '< 0' color=red
    EOF

## Check before you finish

1. Read each command's output. It lists formula errors with their cells,
   and formulas only Excel can compute (shown as `?`), with the reason.
2. Run `sheets lint book.xlsx`: errors, `#REF!`, formulas that break their
   column's pattern, totals that leave out rows, numbers stored as text.
3. `sheets diff book.xlsx` shows what the last command changed, and
   `sheets undo book.xlsx` takes it back.
4. After formatting a sheet or adding a chart, look at it:
   `sheets render book.xlsx Sales` writes a PNG.

## Gotchas

- Values are typed from their text: `2026-07-01` is a date, `12.5%` a
  percentage, `$1,200.50` currency, `1,234` a number, `TRUE` a boolean,
  `=...` a formula. Only ISO dates and decimal points count: `07/01/2026`
  and `1,5` stay text, so convert them first. Leading zeros (`007`) stay
  text; `'` or `--raw` forces text.
- Put formulas, conditions, table references and anything with spaces in
  single quotes: the shell expands `$`, matches `*`, `?` and `[...]` to
  file names, redirects at `<` and `>`, and in zsh expands a leading `=`.
  So `'=SUM($B$2:B9)'`, `'Sales[Revenue]'`, `'< 0'`, `'My Sheet!A1'`.
- Fill a formula down a table column by its name (`'Sales[Margin]'`): a
  guessed range such as `F2:F500` runs past the data and leaves errors in
  the empty rows.
- References follow every edit as in Excel. A formula that loses its cells
  becomes `#REF!`, and the output lists it.
- Results are stored for formulas the sheets calculator computes. LAMBDA,
  FILTER, XMATCH, OFFSET and array arithmetic such as
  `SUMPRODUCT((A2:A9="East")*B2:B9)` get theirs when Excel opens the file.
  When you need to check a number, prefer SUMIFS, COUNTIFS, or SUMPRODUCT
  of plain ranges.
- `--append` adds rows above a total row, whose sums then take them in:
  do not add the new rows to the totals yourself.
- `fmt` replaces the look of a whole sheet with the house style, keeping
  what `style`, `cond` and `view` set. Use it on sheets you build or when
  asked for a consistent look, not after a small edit to someone's
  formatted workbook. Write the data first, then `fmt`, then the styling
  asked for.
- In merged cells, only the first cell takes a value.
- A CSV file holds values only. For formatting, charts or more sheets,
  `sheets convert data.csv data.xlsx` first.
- On files of hundreds of thousands of rows, writing values and inserting
  or deleting rows stay quick. Styling, column and sheet edits load the
  whole workbook, seconds each: put them in one `do` batch.
- Every command takes `--json`, and every command that changes a file
  takes `--dry-run`. A command that fails exits with status 1.

## More detail

Before a task, read its part of the guide:

| Task | Read |
|---|---|
| Reading cells, formula results and patterns | [references/reading.md](references/reading.md) |
| SQL over tables, writing a result back | [references/query.md](references/query.md) |
| Naming cells, ranges and parts of tables (`Sales[#Totals]`) | [references/addresses.md](references/addresses.md) |
| Writing data and formulas, input formats, `--append` | [references/writing.md](references/writing.md) |
| Batches (`do`) | [references/do.md](references/do.md) |
| Inserting or deleting rows, columns and sheets | [references/structure.md](references/structure.md) |
| Clearing, copying, moving, sorting, find and replace | [references/ranges.md](references/ranges.md) |
| Notes and hyperlinks | [references/notes.md](references/notes.md) |
| CSV and TSV files, converting | [references/csv.md](references/csv.md) |
| Checking a workbook (`lint`, `diff`) | [references/check.md](references/check.md) |
| Charts | [references/charts.md](references/charts.md) |
| Looking at a sheet as an image (`render`) | [references/render.md](references/render.md) |
| Dry runs and undo | [references/safety.md](references/safety.md) |
| The house style (`fmt`), column types | [references/fmt.md](references/fmt.md) |
| Styling, highlights, dropdowns, frozen panes | [references/style.md](references/style.md) |
| Large files | [references/large.md](references/large.md) |
| Exit status | [references/exit.md](references/exit.md) |
