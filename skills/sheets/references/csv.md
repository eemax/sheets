# CSV files

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
- So `set` on a CSV computes from the values as they were, and a column
  can change in place: `sheets set prices.csv 'D2:D41' '=ROUND(D{r}*1.05,2)'`
  raises every price by 5%. New values take the style of their column
  (`1.311,45` in a `1.249,00` column).
- `convert` to CSV writes plain values (1200, 2026-03-01, 12%); `--bom`
  helps Excel read accented letters.
