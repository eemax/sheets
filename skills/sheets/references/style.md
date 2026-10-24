# Styling on request: style, cond, validate, view

    sheets style book.xlsx 'Sales[#Headers]' fill=lightblue bold
    sheets style book.xlsx 'Sales[Revenue]' format=currency decimals=0
    sheets style book.xlsx 'Sales[#Totals]' border=top:thin border=bottom:double
    sheets style book.xlsx Sales!A1 size=16 bold
    sheets style book.xlsx Sales!A1:F9 --show                 # how cells look, in these terms
    sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red
    sheets cond book.xlsx Sales!A2:F200 '=$E2="Late"' fill=lightyellow   # whole rows
    sheets cond book.xlsx 'Sales[Units]' bars                 # also: scale, icons
    sheets validate book.xlsx 'Sales[Region]' list North,South,East,West
    sheets validate book.xlsx 'Sales[Units]' whole '>=' 0 --error 'Units cannot be negative'
    sheets view book.xlsx Sales --freeze B2 --tab green --hide F:G

- `fmt` keeps what `style`, `cond` and `view` set: they record it in the
  file (as hidden names that follow inserted and deleted rows), and every
  `fmt` lays it over the house style again. `style --list` shows the
  record, `--unrecord N` removes an entry, `fmt --reset` drops it all.
- Properties: `bold italic underline strike wrap` (`nobold`, ...),
  `font= size= color= fill=` (`#RRGGBB` or a name: lightblue, lightgreen,
  lightyellow, lightred, lightgray, darkblue, red, green, blue, ...),
  `format=TYPE [decimals=N]` (the `--as` types) or an Excel format code,
  `align=left|center|right|center-across`,
  `valign=top|middle|bottom`, `indent=`, `rotate=`,
  `border=SIDES[:STYLE][:COLOR]` (all, outline, inner, top, bottom; thin,
  medium, thick, double), `width=N|auto height=N`, `merge`, `reset`. A
  fill without `color=` gets black or white text, whichever reads.
- A number format on a column (`Sales[Revenue]`) reaches its total row.
  Rows added with `write --append` take the look, conditional formats and
  validation of the rows above.
- `cond` conditions: `'< 0'`, `'>= 100'`, `'= East'`,
  `'between 1 and 10'`, `'top 10'`, `'bottom 5%'`, `'above average'`,
  `duplicates`, `unique`, `blank`, `errors`, `'contains x'`,
  `'begins with x'`, a formula
  for the range's first cell (`'=$E2="Late"'`: the $ keeps column E, so
  every cell of a row tests that row's E), `scale [red-yellow-green]`,
  `bars`, `icons`. The look defaults to Excel's light red. Quote anything
  with < or >.
- `validate` rules: `list A,B,C`, `list =Lists!$A$2:$A$9`,
  `whole|number|date|time|length OP X` or `... between X and Y`,
  `formula =...`; `--prompt`, `--error`, `--warn`.
- `cond` and `validate` take `--list` (what a sheet has) and `--clear`
  (remove what lies within the target). `view` with no flags reports the
  sheet: frozen cell, gridlines, zoom, tab, filter, hidden rows/columns.
