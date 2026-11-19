# The change report's cost on 150,000 rows, wave 2 (2026-11-19)

Every changing command ends with its short change report (wave 2, piece
change-feedback). Its bar: an edit takes at most 1.3 times what the last
pushed binary takes. Measured on this MacBook (M4) by piece report-causes
(run 20261118-sheets-wave2), which also changed what the report compares
(features changed, features that went with their cells, undo's pairing).

- **Base:** sheets' origin/main at 99ba275, the last pushed binary (no
  change report).
- **Measured:** b8aa73d, the report-causes branch. The commits after it
  change no code.
- **Workbook:** 150,000 rows (Region, Units, Price, Revenue = Units ×
  Price), written by the base binary, with Summary!B1 =SUM(Data!D:D),
  B2 =AVERAGE(Data!B2:B150001) and B3 =1/Data!B2.
- **Method:** each run copies the workbook afresh. The two editors are the
  streaming one (`SHEETS_STREAM=always`, the default) and the full one
  (`never`). Three runs each, interleaved base and new by command; the
  median is given.
- **Raw data:** the runs, with each binary's output, are in
  `~/sheets-evals/wave2/report-causes/` (`timings.txt`, `run-all.sh`,
  `timeit.sh`).

| command | editor | base (s) | new (s) | ratio | report lines |
|---|---|---:|---:|---:|---:|
| `set Data!B5 7` | streaming | 0.764 | 0.821 | 1.07 | 3 |
| `set Data!B5 7` | full | 4.463 | 4.520 | 1.01 | 3 |
| `rows insert Data!10` | streaming | 1.480 | 1.527 | 1.03 | 2 |
| `rows insert Data!10` | full | 5.015 | 5.093 | 1.02 | 2 |
| `sort Data Units` | streaming | 7.820 | 7.845 | 1.00 | 4 |
| `sort Data Units` | full | 7.803 | 7.907 | 1.01 | 4 |

Every command is within the bar: 1.00 to 1.07 times the base. The report
costs little for three reasons:

- it is built from the results the save recomputed anyway;
- the features are read off the file before and after, side by side;
- the lint is compared only up to 512 KB of sheet XML, so on this
  workbook the report says `lint not compared`.

The change-feedback build measured the same commands at 88d937b, 1.00 to
1.10 times the base (`~/sheets-evals/wave2/d2/README.md`).
