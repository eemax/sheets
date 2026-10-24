# Safety: dry runs and undo

    sheets --dry-run rows delete book.xlsx Sales!5:9   # report, save nothing
    sheets undo book.xlsx                              # take back the last change
    sheets undo book.xlsx --list                       # what can be undone

- `--dry-run` works with every command that changes a file.
- Every change keeps the previous version of the file (the last 20 per
  file). `undo` restores it; run it again to go further back. It refuses
  if something else changed the file since; `--force` restores anyway.
