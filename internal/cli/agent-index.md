# sheets: guide for agents

`sheets` reads, writes and formats .xlsx files, and reads and edits CSV.
Its output is compact and sized for your context. Start with an overview
and drill down; never dump a whole sheet. Every output ends with a `next:`
line suggesting a follow-up command.

    sheets info book.xlsx                 # sheets, used ranges, tables, charts, names
    sheets peek book.xlsx [Sheet]         # one line per column: type, range, values, formulas
    sheets read book.xlsx Sheet!A1:F20    # cells; --formulas shows formulas
    sheets find book.xlsx "acme"          # matching cells with their column header
    sheets query book.xlsx 'SELECT Region, SUM(Revenue) FROM Sales GROUP BY Region'
    printf 'Date\tAmount\n2026-07-01\t$1,200\n' | sheets write book.xlsx Sales!A1
    sheets write book.xlsx Sales --append -i rows.json
    sheets set book.xlsx 'Sales!F2:F500' '=D{r}*E{r}'    # {r}: each cell's row
    sheets rows delete book.xlsx Sales!5:9
    sheets sort book.xlsx Sales Revenue:desc
    sheets fmt book.xlsx                  # the house style
    sheets style book.xlsx 'Sales[Revenue]' format=currency decimals=0
    sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red
    sheets lint book.xlsx                 # errors and broken patterns, before handing over
    sheets diff book.xlsx                 # what the last command changed
    sheets undo book.xlsx

- Use `peek` before `read`, and `query` (SQLite over every table) to
  aggregate, filter or join.
- Values are typed from their text: `2026-07-01` is a date, `12.5%` a
  percentage, `$1,200.50` currency, `1,234` a number, `TRUE` a boolean,
  `=...` a formula. Only ISO dates and decimal points are recognised:
  convert `07/01/2026` and `1,5` first. Prefix `'` or pass `--raw` for
  text.
- Formulas are computed when written and the output lists their errors:
  check it. `?` marks a result only Excel can compute (LAMBDA, FILTER, ...).
- References follow every edit as in Excel; ones that lose their cells
  become #REF! and are listed.
- Write data first, then `fmt`, then the styling asked for (`style`,
  `cond`, `validate`, `view`).
- More than two or three changes: one `sheets do book.xlsx` batch (a
  command per line, saved once). `--dry-run` reports a change without
  saving.

More on a topic: `sheets agent-docs TOPIC`, one or more of: reading,
query, references, writing, do, structure (rows, columns, sheets), ranges
(clear, copy, move, sort, replace), notes (and links), csv, check (lint,
diff), charts, render, safety (dry runs, undo), fmt, style (style, cond,
validate, view), large (large files), exit. `sheets agent-docs all` prints
the whole guide.
