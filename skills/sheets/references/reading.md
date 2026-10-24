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
