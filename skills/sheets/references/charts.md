# Charts

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
