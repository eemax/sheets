# Looking at a sheet: render

    sheets render book.xlsx Sales                 # writes book-Sales.png next to the file
    sheets render book.xlsx 'Sales!A1:H30' -o check.png
    sheets render book.xlsx Sales -o page.html    # an HTML page for people

- Draws values as formatted in the file, fonts, fills, borders, merged
  cells, banding and charts, with column letters and row numbers. Look at
  the PNG to check formatting or a chart before handing a file over.
- It approximates Excel: other fonts stand in for Calibri, and conditional
  formats other than banding and Excel table styles are not drawn (the
  output says so; `cond --list` shows them). `style --show` lists how
  cells look, as text.
