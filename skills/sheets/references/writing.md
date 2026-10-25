# Writing

    printf 'Date\tAmount\n2026-07-01\t$1,200\n' | sheets write book.xlsx Sales!A1
    sheets write book.xlsx Sales --append -i rows.json
    sheets set book.xlsx Sales!F2 '=D2*E2'
    sheets set book.xlsx 'Sales!F2:F500' '=D2*E2'    # fill down: F3 gets =D3*E3
    sheets set book.xlsx 'Sales!F2:F500' '=D{r}*E{r}' # the same, with {r} = the row
    sheets set book.xlsx 'Sales[Revenue]' '=D{r}*E{r}' # a table column's data rows
    sheets set book.xlsx Summary!B1 'Q3 total' Summary!B2 '=SUM(Sales!F:F)'

- The file and sheet are created when missing.
- `set` takes a table reference (`Sales[Revenue]`, see the addresses
  topic) for exactly a column's data rows, its total row left alone. The
  column needs its header first. A range past the data fills empty rows,
  where a formula reading them gives errors.
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
  take in the new rows, and so do formulas anywhere covering one of its
  columns whole (`=SUM(Sales!$D$2:$D$61)`), as with an Excel table. If the
  first input row (or the JSON keys) names the table's headers, columns
  are matched by name, in any order.
- Appended rows get the formulas of the columns that compute, as Excel
  fills them: an Excel table's calculated column, or a column whose last
  rows hold one formula filled down. Give only the input columns; a value
  you give for a formula column is kept. The output names the columns it
  filled.
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
  SEQUENCE, ...) are stored as dynamic arrays and spill: the file holds the
  whole result, as Excel saves it, so pandas and other readers see every
  value. Keep the cells below and to the right free: a result that would
  cover values is #SPILL!, as in Excel, and writing into a spill area gives
  a warning.
- Formulas computing over a range where one value goes
  (`=SUM(IF(A2:A9>5,B2:B9))`, `=MAX(B2:B9-C2:C9)`, `=A2:A9*2`) are stored
  as Excel 365 stores them, as dynamic array formulas, so Excel computes
  them over the whole range; no Ctrl+Shift+Enter needed. `=A2:A9*2` spills.
  A whole column in a result (`=B:B*2`) still means the formula's own row,
  as written for older Excel.
- A cell hidden inside merged cells (B1 of a merged A1:D1) takes no value:
  Excel shows only the first cell. Writing one is refused with the merge
  named; empty values there are skipped.
- Write data first, then `fmt`, then the styling asked for (`style`,
  `cond`, `validate`, `view`).
