# Many changes at once: do

    sheets do report.xlsx <<'EOF'
    write Sales!A1 -i sales.csv
    set Sales!F1 Margin
    set 'Sales[Margin]' '=E{r}/D{r}'
    fmt
    style 'Sales[#Headers]' fill=lightblue
    cond 'Sales[Margin]' '< 0' color=red
    EOF

- One command per line, as on the command line without `sheets` and the
  file. The file is saved once at the end, as one undo step; the first
  error stops the batch and nothing is saved. Each command sees the
  changes before it (`read` and `peek` in a batch show them).
- Data for `write` comes from `-i file`, `< file`, or a here-document in
  the script (`write Sales!A1 <<END`, the rows, `END`).
- Use it whenever a change takes more than two or three commands.
