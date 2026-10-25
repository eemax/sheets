# Reading

    sheets info book.xlsx                 # sheets, used ranges, tables, charts, named ranges
    sheets peek book.xlsx [Sheet]         # one line per column: type, range, distinct values, formulas
    sheets read book.xlsx Sheet!A1:F20    # cells, with row numbers and column letters
    sheets find book.xlsx "acme"          # matching cells with their column header
    sheets query book.xlsx 'SELECT ...'   # SQL over the tables
    sheets render book.xlsx Sheet         # a PNG of the sheet, to look at

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
  #DIV/0!. Numbers show up to 15 significant digits, without the noise of
  binary arithmetic (24836587.21, not 24836587.2099999); the file keeps
  every digit.
- Formula results come from the file. When `sheets` saves, it recomputes
  every formula the change reaches with its own calculator and stores the
  result, so pandas, previews and other tools read current numbers, never
  stale ones. The calculator follows Excel: tested on 68,000 formulas of
  800 workbooks Excel saved, it gives Excel's result for 99.8% of those it
  computes. It does about 250 functions, among them XLOOKUP, FILTER, SORT,
  UNIQUE, LET, the financial ones (NPV, IRR, PMT, XIRR), table references
  (`Sales[Units]`, `[@Price]`) and array arithmetic
  (`SUMPRODUCT((A2:A9="East")*B2:B9)`).
- `?` marks a formula with no result yet, with the reason: a link to
  another workbook, a function the calculator lacks (LAMBDA, CELL,
  statistical distributions such as NORM.DIST, bond pricing), or a
  formula reading one. Excel computes it on open. A result that depends
  on Excel's language is not guessed either: text such as `1,200`,
  `1.5`, `03/05/2026` or `May 2026` reads as a number or date differently
  in other languages, so a formula computing with it is left to Excel.
  Write numbers and ISO dates as values, and the calculator computes.
  Circular references are reported and get no result.
- A report laid out in sections under one header (a heading row such as
  `West`, its lines, a subtotal; then `North`, ...) reads as one table:
  `info` lists the sections, and `query` and `lint` take in all of them.
