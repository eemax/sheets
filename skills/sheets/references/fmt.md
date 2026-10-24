# Formatting

    sheets fmt book.xlsx                  # house style
    sheets fmt book.xlsx -p finance       # adds blue inputs / black formulas / green links
    sheets fmt book.xlsx --as Revenue=currency --as 'Q3!D=percent'
    sheets fmt book.xlsx --check          # exit 1 if fmt would change anything

`fmt` removes existing formatting and applies one consistent style, then
lays over it what `style`, `cond` and `view` set:
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
