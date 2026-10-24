# Large files

Reading a large workbook (over 1 MB) the first time keeps a parsed copy in
the user cache, so later commands on the same file start in a fraction of
the time. The copy is used only for exactly the same file contents.
`SHEETS_CACHE=off` turns this off; `SHEETS_CACHE_DIR` moves the cache.

Writing values, inserting and deleting rows, and appending to a table
rewrite only the sheets they change, so they stay quick on large files:
about a second for one cell of a 200,000-row sheet (three the first time
the file is read). Other edits (styling,
columns, sheets, copy, move, sort), and listing styling (`cond --list`,
`style --show`, `view`), load the whole workbook, which takes several
seconds at that size. Put a large file's styling into one `do` batch.

Formula results missing from a file (written by a tool that does not
calculate, such as openpyxl) are computed on demand. On a large sheet the
first such value takes a few seconds. Files saved by Excel carry their
results.
