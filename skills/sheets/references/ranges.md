# Moving data: clear, copy, move, sort, replace

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
