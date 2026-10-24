# sheets: guide for agents

`sheets` reads, writes and formats .xlsx files, and reads and edits CSV.
Its output is compact and sized for your context. Start with an overview and drill down; never dump
a whole sheet. Every output ends with a `next:` line suggesting a follow-up
command.

## Reading

    sheets info book.xlsx                 # sheets, used ranges, tables, charts, named ranges
    sheets peek book.xlsx [Sheet]         # one line per column: type, range, distinct values, formulas
    sheets read book.xlsx Sheet!A1:F20    # cells, with row numbers and column letters
    sheets find book.xlsx "acme"          # matching cells with their column header
    sheets query book.xlsx 'SELECT ...'   # SQL over the tables (see below)
    sheets render book.xlsx Sheet         # a PNG of the sheet, to look at (see below)

- `peek` output stays the same size whether a table has 10 rows or a
  million. Use it before `read`.
- Two-row headers (Q1 over Units and Revenue, Q2 over Units and Revenue)
  name columns `Q1 / Units`, `Q1 / Revenue`, ...; `peek` says
  `header rows 2-3`. Use those names in `--as`, `query` and `--append`.
- `info` and `peek` flag hidden rows and columns (`hidden rows 5:9`). Their
  cells are still read and counted; a filter in Excel may be hiding them.
- In `peek`, formulas are shown once per column as a pattern, with `{r}` for
  the current row: `=D{r}-E{r}` means every row computes D minus E. A column
  listing several patterns has inconsistent formulas, which is worth checking.
- `read` prints at most 50 rows (`-n` to change) and 30 columns
  (`--max-cols`), then gives the command for the next page.
  `--formulas` shows formulas instead of values.
  `--format tsv|csv|json` gives plain data; `--json` works on every command.
- Values print plainly: ISO dates (2026-07-01), percentages with % (12.5%),
  numbers without thousands separators, TRUE/FALSE, error codes like
  #DIV/0!.
- Formula results come from the file. When `sheets` saves, it recomputes
  every formula the change reaches and stores the result, so pandas,
  previews and other tools read current numbers, never stale ones. Table
  references (`Sales[Units]`, `[@Price]`) and LET are computed. `?` marks a
  formula the calculator cannot evaluate (LAMBDA, FILTER, XMATCH, OFFSET,
  array arithmetic such as `SUMPRODUCT((A2:A9="East")*B2:B9)`, and a few
  others), or one that reads such a formula: it has no stored result until
  Excel computes it on open, and the output of the change lists it with the
  reason. Prefer forms it computes (`SUMIFS`, `COUNTIFS`, `SUMPRODUCT` of
  plain ranges) so you can check the result. Circular references are
  reported and get no result.

## SQL queries

    sheets query book.xlsx --schema
    sheets query book.xlsx 'SELECT Region, SUM(Revenue) AS Revenue FROM Sales GROUP BY Region ORDER BY 2 DESC'
    sheets query book.xlsx 'SELECT _row, * FROM Sales WHERE Units > 100'
    sheets query book.xlsx 'SELECT ...' --into Summary        # write the result as a new table

- Every table `sheets` detects is a SQL table (SQLite dialect) named after
  its sheet (`Sales`; a second table on that sheet is `Sales_2`) or its
  Excel table name. Columns are the header labels; quote names with
  spaces: `"Order ID"`. `--schema` lists tables and columns with types.
- Numbers are numbers (percentages as fractions: 12.5% is 0.125). Dates
  are ISO text, so `date()`, `strftime()` and comparisons work. Errors and
  blanks are NULL. Total rows are left out. `_row` is the sheet row of a
  record, for going back to the cells.
- `--into Sheet` replaces that sheet with the result, `--into Sheet!C5`
  writes it at that cell. Columns named like a source column keep its
  number format.
- Prefer `query` over reading rows to aggregate, filter or join: the
  answer is a few lines whatever the size of the table.

## References

`Sheet!A1:C9`, `Sheet` (its used range), `A1:C9` (first sheet), whole
columns `Sheet!A:C`, whole rows `Sheet!5:20`. Sheet names with spaces need
no inner quotes: `'My Sheet!A1'` is fine (the outer quotes are for the
shell).

`style`, `cond` and `validate` also take table references in Excel's
syntax, named by a sheet holding one table or by an Excel table, so you
need no column letters: `Sales[Revenue]` (the column's data rows),
`Sales[[Units]:[Revenue]]`, `Sales[#Headers]`, `Sales[#Totals]`,
`Sales[#All]`, `Sales[[#Totals],[Units]]`.

## Writing

    printf 'Date\tAmount\n2026-07-01\t$1,200\n' | sheets write book.xlsx Sales!A1
    sheets write book.xlsx Sales --append -i rows.json
    sheets set book.xlsx Sales!F2 '=D2*E2'
    sheets set book.xlsx 'Sales!F2:F500' '=D2*E2'    # fill down: F3 gets =D3*E3
    sheets set book.xlsx 'Sales!F2:F500' '=D{r}*E{r}' # the same, with {r} = the row
    sheets set book.xlsx Summary!B1 'Q3 total' Summary!B2 '=SUM(Sales!F:F)'

- The file and sheet are created when missing.
- Input for `write` is TSV, CSV or JSON. JSON can be an array of arrays, or
  an array of objects whose keys become the header row. The format is
  detected automatically.
- Values are typed from their text: `2026-07-01` is a date,
  `2026-07-01 13:30` a datetime, `12.5%` a percentage, `$1,200.50` or
  `250 kr` currency, `1,234` a number, `TRUE` a boolean, `=...` a formula.
  Leading zeros (`007`) and digit strings over 15 digits stay text. Prefix
  with `'` or pass `--raw` to force text.
- Only ISO dates are recognised: `07/01/2026` stays text, because it means
  July 1 in the US and January 7 elsewhere. Convert dates to `YYYY-MM-DD`
  before writing.
- Numbers use a decimal point. `1,5` (a decimal comma) stays text; `1,500`
  is fifteen hundred. Convert European-style numbers before writing.
- `--append` adds rows to the sheet's last table: above its total row
  (whose sums then cover the new rows), or below it, moving any notes
  under it down. Charts of the table and its Excel table (if it is one)
  take in the new rows. If the first input row (or the JSON keys) names
  the table's headers, columns are matched by name, in any order.
- `--clear` empties the sheet first.
- In formulas, `{r}` stands for the row the value lands on. `=C{r}*D{r}`
  written to row 7 becomes `=C7*D7`. This is the notation `peek` prints,
  and it works with `--append`, where you don't know the row in advance.
  `{r-1}` and `{r+1}` work too.
- TSV input keeps backslashes as they are (so `C:\new` stays intact). For
  values that contain tabs or newlines, use CSV or JSON.
- Formulas are evaluated right away. The output shows a few values, every
  formula error (with the cell), and formulas the calculator cannot
  evaluate. Check it.
- All Excel functions work, including XLOOKUP, IFS, TEXTJOIN, LET and
  LAMBDA. Functions that return several values (FILTER, UNIQUE, SORT,
  SEQUENCE, ...) are stored as dynamic arrays and spill in Excel; keep the
  cells below and to the right free. Writing into the spill area of an
  existing dynamic array gives a warning (Excel would show #SPILL!).
- A cell hidden inside merged cells (B1 of a merged A1:D1) takes no value:
  Excel shows only the first cell. Writing one is refused with the merge
  named; empty values there are skipped.
- Write data first, then `fmt`, then the styling asked for (`style`,
  `cond`, `validate`, `view`).

## Many changes at once: do

    sheets do report.xlsx <<'EOF'
    write Sales!A1 -i sales.csv
    set Sales!F1 Margin Sales!F2:F200 '=E{r}/D{r}'
    fmt
    style 'Sales[#Headers]' fill=lightblue
    cond 'Sales[Margin]' '< 0' color=red
    EOF

- One command per line, as on the command line without `sheets` and the
  file. The file is saved once at the end, as one undo step; the first
  error stops the batch and nothing is saved. Each command sees the
  changes before it (`read` and `peek` in a batch show them).
- Data for `write` comes from `-i file`, `< file`, or a here-document in
  the script (`write Sales!A1 <<END`, the rows, `END`).
- Use it whenever a change takes more than two or three commands.

## Rows, columns and sheets

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

## Moving data: clear, copy, move, sort, replace

    sheets sort book.xlsx Sales Revenue:desc         # a table's data rows, by header names
    sheets sort book.xlsx Sales Region Units:desc    # several keys
    sheets copy book.xlsx Sales!E2 Sales!E3:E200     # fill a formula down
    sheets copy book.xlsx Model!B2:B20 Model!H2 --values
    sheets move book.xlsx Sales!H1:J20 Sales!A30     # cut and paste
    sheets clear book.xlsx Sales!A2:D100             # contents; --format, --notes, --links, --all
    sheets replace book.xlsx 'Acme Inc.' 'Acme Corp' # every sheet, or give a target
    sheets replace book.xlsx 'N/A' '' Sales --exact  # empties those cells

- These behave as in Excel. `copy` shifts relative references (A2 becomes
  A12 ten rows down); `--values` pastes results with their number formats.
  `move` keeps moved formulas reading the same cells, and formulas, names
  and charts anywhere follow the moved cells; references to the cells it
  pastes over become #REF! and are listed.
- `sort` moves whole rows of the target with their formats and notes;
  blanks always sort last. It refuses a sort that would silently change a
  formula's meaning (one reading another row by position, or a column left
  out of the sort) and says how to fix it. Running totals and other
  columns of alike formulas are fine.
- `replace` reads a changed value as typed input ("1,200" becomes a
  number); `--formulas` replaces in formula text instead; `--regex` takes
  groups as $1.

## Notes and links

    sheets note book.xlsx Sales!D7 'Includes the March refund'   # '' removes it
    sheets note book.xlsx --list
    sheets link book.xlsx Sales!A1 https://example.com/q3 --text 'Q3 report'
    sheets link book.xlsx Summary!B2 Sales!A1                    # a place in the workbook
    sheets link book.xlsx --list

`info` counts a sheet's notes and links. Notes and links move with their
cells. Excel 365 comment threads are listed but cannot be changed.

## CSV files

    sheets peek data.csv
    sheets sort data.csv data Amount:desc
    sheets convert data.csv data.xlsx      # then: sheets fmt data.xlsx
    sheets convert book.xlsx sales.csv --sheet Sales

- A .csv or .tsv file is one sheet named after the file. Reading commands,
  `write`, `set`, `sort`, `replace`, `clear`, `copy`, `move`, `rows` and
  `cols` work on it directly. Its delimiter (also `;` with decimal commas), encoding
  and line breaks are kept, and rows no edit touched stay byte for byte.
- A CSV file holds values only: formulas are saved as their results, and
  formatting, charts or more sheets need `convert` to .xlsx first.
- `convert` to CSV writes plain values (1200, 2026-03-01, 12%); `--bom`
  helps Excel read accented letters.

## Checking: lint and diff

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

## Charts

    sheets chart add book.xlsx Sales                        # first label column by the number columns
    sheets chart add book.xlsx Sales --x Region --y Revenue --type bar
    sheets chart add book.xlsx Sales --x Date --y Units,Returns --type line --at Summary!B2
    sheets chart list book.xlsx
    sheets chart delete book.xlsx Sales!H2

- The source is a table: a sheet, an Excel table name, `Sheet_2`, or a
  range whose first row is the header. Total rows are left out.
- Without `--type`, dates get a line, many or long labels a bar chart,
  anything else columns. Types: column, bar, line, area, pie, scatter;
  `--stacked` for column, bar and area.
- Charts take the house style: a palette that stays distinguishable with
  colour-vision deficiency, no legend for a single series, hairline
  gridlines, thin bars.
- Charts that would mislead are refused, with the reason: more than 8
  series, pies over 6 slices, columns in different units on one axis (plot
  them in separate charts).
- `render` the sheet afterwards to see the chart.

## Looking at a sheet: render

    sheets render book.xlsx Sales                 # writes book-Sales.png next to the file
    sheets render book.xlsx 'Sales!A1:H30' -o check.png
    sheets render book.xlsx Sales -o page.html    # an HTML page for people

- Draws values as formatted in the file, fonts, fills, borders, merged
  cells, banding and charts, with column letters and row numbers. Look at
  the PNG to check formatting or a chart before handing a file over.
- It approximates Excel: other fonts stand in for Calibri, and conditional
  formats other than banding and Excel table styles are not drawn (the
  output says so; `cond --list` shows them). `style --show` lists how
  cells look, as text.

## Safety: dry runs and undo

    sheets --dry-run rows delete book.xlsx Sales!5:9   # report, save nothing
    sheets undo book.xlsx                              # take back the last change
    sheets undo book.xlsx --list                       # what can be undone

- `--dry-run` works with every command that changes a file.
- Every change keeps the previous version of the file (the last 20 per
  file). `undo` restores it; run it again to go further back. It refuses
  if something else changed the file since; `--force` restores anyway.

## Formatting

    sheets fmt book.xlsx                  # house style
    sheets fmt book.xlsx -p finance       # adds blue inputs / black formulas / green links
    sheets fmt book.xlsx --as Revenue=currency --as 'Q3!D=percent'
    sheets fmt book.xlsx --check          # exit 1 if fmt would change anything

`fmt` removes existing formatting and applies one consistent style, then
lays over it what `style`, `cond` and `view` set (see below):
- a dark header row with filters, labels left-aligned clear of the filter
  buttons
- banded data rows
- thin borders
- a bold total row
- a frozen header row
- fitted column widths, with long text wrapped
- number formats: thousands separators, negatives in parentheses, ISO
  dates, two decimals for currency with cents

The output lists the type it chose for each column. If a type is wrong,
re-run with `--as COLUMN=TYPE`. COLUMN is a header or a column letter.
TYPE is one of integer, decimal, currency[:symbol], percent, date,
datetime, time or text. The choice is saved in the file's number formats,
so later runs keep it.

Excel tables (Insert › Table) keep their own table style. `fmt` only sets
their number formats and column widths, and says so in a note.

Two-row headers are styled as one header: the group labels (Q1, Q2) are
centred across the columns they group.

Detection rules:
- Tables are blocks of cells with a header row. Header labels are text, or
  years (Region | 2024 | 2025).
- A single text cell just above a table is its title.
- A last row starting with "Total" (or holding only SUM formulas) is the
  total row.
- Columns named like ids or years (Order ID, Year, Zip) get no thousands
  separator.
- A formula column takes the type of what it computes from: `=Units*Price`
  is currency when Price is, and `=Start+30` is a date. Divisions (ratios)
  and date differences (day counts) stay numbers.

## Styling on request: style, cond, validate, view

    sheets style book.xlsx 'Sales[#Headers]' fill=lightblue bold
    sheets style book.xlsx 'Sales[Revenue]' format=currency decimals=0
    sheets style book.xlsx 'Sales[#Totals]' border=top:thin border=bottom:double
    sheets style book.xlsx Sales!A1 size=16 bold
    sheets style book.xlsx Sales!A1:F9 --show                 # how cells look, in these terms
    sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red
    sheets cond book.xlsx Sales!A2:F200 '=$E2="Late"' fill=lightyellow   # whole rows
    sheets cond book.xlsx 'Sales[Units]' bars                 # also: scale, icons
    sheets validate book.xlsx 'Sales[Region]' list North,South,East,West
    sheets validate book.xlsx 'Sales[Units]' whole '>=' 0 --error 'Units cannot be negative'
    sheets view book.xlsx Sales --freeze B2 --tab green --hide F:G

- `fmt` keeps what `style`, `cond` and `view` set: they record it in the
  file (as hidden names that follow inserted and deleted rows), and every
  `fmt` lays it over the house style again. `style --list` shows the
  record, `--unrecord N` removes an entry, `fmt --reset` drops it all.
- Properties: `bold italic underline strike wrap` (`nobold`, ...),
  `font= size= color= fill=` (`#RRGGBB` or a name: lightblue, lightgreen,
  lightyellow, lightred, lightgray, darkblue, red, green, blue, ...),
  `format=TYPE [decimals=N]` (the `--as` types) or an Excel format code,
  `align=left|center|right|center-across valign=top|middle|bottom indent=
  rotate=`, `border=SIDES[:STYLE][:COLOR]` (all, outline, inner, top,
  bottom; thin, medium, thick, double), `width=N|auto height=N`, `merge`,
  `reset`. A fill without `color=` gets black or white text, whichever
  reads.
- A number format on a column (`Sales[Revenue]`) reaches its total row.
  Rows added with `write --append` take the look, conditional formats and
  validation of the rows above.
- `cond` conditions: `'< 0'`, `'>= 100'`, `'= East'`, `'between 1 and
  10'`, `'top 10'`, `'bottom 5%'`, `'above average'`, `duplicates`,
  `unique`, `blank`, `errors`, `'contains x'`, `'begins with x'`, a formula
  for the range's first cell (`'=$E2="Late"'`: the $ keeps column E, so
  every cell of a row tests that row's E), `scale [red-yellow-green]`,
  `bars`, `icons`. The look defaults to Excel's light red. Quote anything
  with < or >.
- `validate` rules: `list A,B,C`, `list =Lists!$A$2:$A$9`,
  `whole|number|date|time|length OP X` or `... between X and Y`,
  `formula =...`; `--prompt`, `--error`, `--warn`.
- `cond` and `validate` take `--list` (what a sheet has) and `--clear`
  (remove what lies within the target). `view` with no flags reports the
  sheet: frozen cell, gridlines, zoom, tab, filter, hidden rows/columns.

## Large files

Reading a large workbook (over 1 MB) the first time keeps a parsed copy in
the user cache, so later commands on the same file start in a fraction of
the time. The copy is used only for exactly the same file contents.
`SHEETS_CACHE=off` turns this off; `SHEETS_CACHE_DIR` moves the cache.

Writing values, inserting and deleting rows, and appending to a table
rewrite only the sheets they change, so they stay quick on large files:
about a second for one cell of a 200,000-row sheet (three the first time
the file is read). Other edits (styling,
columns, sheets, copy, move, sort), and listing styling (`cond --list`,
`style --show`, `view`), load the whole workbook, which takes several
seconds at that size. Put a large file's styling into one `do` batch.

Formula results missing from a file (written by a tool that does not
calculate, such as openpyxl) are computed on demand. On a large sheet the
first such value takes a few seconds. Files saved by Excel carry their
results.

## Exit status

0 on success. 1 on errors (a message goes to stderr, or `{"error": ...}`
with `--json`), when `fmt --check` finds changes, when `lint` finds
problems at `--fail-on` level, and when `diff --exit-code` finds
differences.
