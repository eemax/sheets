# SQL queries

    sheets query book.xlsx --schema
    sheets query book.xlsx 'SELECT Region, SUM(Revenue) AS Revenue FROM Sales GROUP BY Region ORDER BY 2 DESC'
    sheets query book.xlsx 'SELECT _row, * FROM Sales WHERE Units > 100'
    sheets query book.xlsx 'SELECT ...' --into Summary        # write the result as a new table

- Every table `sheets` detects is a SQL table (SQLite dialect) named after
  its sheet (`Sales`; a second table on that sheet is `Sales_2`) or its
  Excel table name. Columns are the header labels; quote names with
  spaces: `"Order ID"`. `--schema` lists tables and columns with types.
- A report in sections is one table of every section's lines, with a
  `_section` column holding each line's heading (`West`), so
  `WHERE _section = 'West'` or `GROUP BY _section` work. Subtotals and the
  grand total are left out.
- Numbers are numbers (percentages as fractions: 12.5% is 0.125). Dates
  are ISO text, so `date()`, `strftime()` and comparisons work. Errors and
  blanks are NULL. Total rows are left out. `_row` is the sheet row of a
  record, for going back to the cells.
- `--into Sheet` replaces that sheet with the result, `--into Sheet!C5`
  writes it at that cell. Columns named like a source column keep its
  number format.
- Prefer `query` over reading rows to aggregate, filter or join: the
  answer is a few lines whatever the size of the table.
