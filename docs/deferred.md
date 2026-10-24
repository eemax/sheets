# Deferred

Work we chose not to do yet, with the reason and what it would take. Pick
from here when there is time; the reasons still apply unless noted.

## Decided against for now

- **A patched copy of excelize.** The user's call (2026-10-23): "no patched
  copy of excelize at this time, but we might consider it later". It would
  fix, at the source, what we now work around:
  - deleting rows goes through excelize one row at a time;
  - `UnsetConditionalFormat` and `DeleteDataValidation` take cells out of
    ranges one at a time, which is slow on large ranges;
  - `GetConditionalFormats` loses rules that share a range;
  - conditional-format formulas are not adjusted when rows or columns move
    (we rewrite them in a patch on save);
  - `SetConditionalFormat` can repeat a priority and puts new rules below
    existing ones (we renumber them on save).
- **Better-looking charts.** Parked by the user (2026-10-23). The plan when
  it comes up: write the chart XML ourselves rather than use excelize's
  chart API, so the house dataviz style (thin marks, recessive axes, direct
  labels) is fully ours.

## Larger pieces of work

- **Our own calculation engine.** Results excelize cannot compute are left
  for Excel (shown as `?`). An engine of our own that evaluates the common
  functions straight from parsed data would remove the excelize round trip
  on large files. Worth doing together with streaming, where the speed
  matters. Until then, LAMBDA and other uncomputed functions rely on Excel
  or on LibreOffice when it is installed.
- **Drawing conditional formats in `render`.** `render` draws cell styles,
  banding and Excel table styles, but not highlight rules, colour scales,
  data bars or icons. Cell-value rules, scales and bars could be drawn from
  the stored values; formula rules need the calculator inside the
  renderer, which is why it was deferred. Agents use `cond --list` to
  check instead.
- **OpenDocument (.ods).** The format seam (`internal/grid`, `internal/book`)
  is there; .ods needs a reader and a writer.
- **Custom house styles.** The presets (`default`, `finance`) are built in.
  They are shaped so a config file (TOML) could override them per project.

## Smaller gaps

- Columns cannot be inserted into or deleted from an Excel table (the
  command refuses). It would need the table's columns, totals formulas and
  structured references rewritten.
- Pivot table sources follow sheet renames only, not row or column edits.
- `sheet copy` refuses a sheet holding an Excel table. Table names must be
  unique, so the copy needs a new name (Excel picks one) and every
  structured reference on the copied sheet must be renamed with it.
- `fmt` leaves merged cells alone, and Excel tables keep their own table
  style; `fmt` only sets their number formats and widths.
- `render` stands other fonts in for Calibri, so widths and wrapping are
  approximate.
- openpyxl drops the Excel 2010 extension that gives data bars their solid
  fill and negative colours; the bars survive as plain gradient bars. This
  is openpyxl's behaviour, noted in case a user reports it.

## Evaluation

- The `claude` CLI was not logged in on this machine, so the A/B runs used
  Sonnet subagents reading each task's prompt. `evals/run.py run` drives the
  CLI directly once it is logged in, and records cost too.
- The plan was about 15 tasks; there are 12. Missing: a messy human-made
  workbook to clean up, a range-operations task (sort, replace, move), and
  a large-file edit.
- Each phase should be followed by a rerun of both arms.

## Waiting on the user

- Opening the `excel-check/` workbooks in real Excel (built by
  `scripts/excel-check.sh`, checklist in `excel-check/CHECKLIST.md`). The
  test suite cannot catch a file Excel would repair.
