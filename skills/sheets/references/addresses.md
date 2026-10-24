# References

`Sheet!A1:C9`, `Sheet` (its used range), `A1:C9` (first sheet), whole
columns `Sheet!A:C`, whole rows `Sheet!5:20`. Sheet names with spaces need
no inner quotes: `'My Sheet!A1'` is fine (the outer quotes are for the
shell).

`set`, `style`, `cond` and `validate` also take table references in Excel's
syntax, named by a sheet holding one table or by an Excel table, so you
need no column letters: `Sales[Revenue]` (the column's data rows),
`Sales[[Units]:[Revenue]]`, `Sales[#Headers]`, `Sales[#Totals]`,
`Sales[#All]`, `Sales[[#Totals],[Units]]`.
