# sheets: guide for agents

`sheets` reads, writes and formats .xlsx files. Its output is compact and
sized for your context. Start with an overview and drill down; never dump
a whole sheet. Every output ends with a `next:` line suggesting a follow-up
command.

## Reading

    sheets info book.xlsx                 # sheets, used ranges, tables, named ranges
    sheets peek book.xlsx [Sheet]         # one line per column: type, range, distinct values, formulas
    sheets read book.xlsx Sheet!A1:F20    # cells, with row numbers and column letters
    sheets find book.xlsx "acme"          # matching cells with their column header

- `peek` output stays the same size whether a table has 10 rows or a
  million. Use it before `read`.
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
- Formula results come from the file. After `sheets` has written to a file,
  Excel has not recalculated it yet, so `sheets` computes the results
  itself. `?` marks a formula its calculator cannot evaluate (LET, LAMBDA,
  FILTER and a few others); Excel computes those on open.

## References

`Sheet!A1:C9`, `Sheet` (its used range), `A1:C9` (first sheet), whole
columns `Sheet!A:C`, whole rows `Sheet!5:20`. Sheet names with spaces need
no inner quotes: `'My Sheet!A1'` is fine (the outer quotes are for the
shell).

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
- `--append` writes below the last used row. If the first input row (or
  the JSON keys) names the sheet's headers, columns are matched by name, in
  any order.
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
- Write data first, then format.

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
- Renaming a sheet updates every reference to it, including in charts.
- `sheet delete` refuses while formulas elsewhere still refer to the sheet,
  and lists them. `--force` deletes anyway; those references become #REF!.
- Chart data ranges do not move when rows or columns are inserted or
  deleted; the output warns when the workbook has charts.

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

`fmt` removes existing formatting and applies one consistent style:
- a dark header row with filters
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

## Exit status

0 on success. 1 on errors (a message goes to stderr, or `{"error": ...}`
with `--json`) and when `fmt --check` finds changes.
