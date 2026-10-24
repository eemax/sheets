# Deferred

Work we chose not to do yet, with the reason and what it would take. Pick
from here when there is time; the reasons still apply unless noted.

## Decided against for now

- **A patched copy of excelize.** The user's call (2026-10-23): "no patched
  copy of excelize at this time, but we might consider it later". It would
  fix, at the source, what we now work around:
  - deleting rows goes through excelize one row at a time (streaming
    avoids it for value and row edits; a session that also styles or
    edits columns still pays it);
  - excelize moves conditional formats and validations wrongly when rows
    or columns go, and removes them only cell by cell (seconds on a large
    range). We keep conditional formats out of excelize altogether and
    take validations off around each edit;
  - `GetConditionalFormats` loses rules that share a range;
  - `SetConditionalFormat` can repeat a priority and puts new rules below
    existing ones (we renumber them on save);
  - a value written into a merge's hidden cell lands on its first cell (we
    refuse it before excelize sees it).
- **Better-looking charts.** Parked by the user (2026-10-23). The plan when
  it comes up: write the chart XML ourselves rather than use excelize's
  chart API, so the house dataviz style (thin marks, recessive axes, direct
  labels) is fully ours.

## Larger pieces of work

- **Our own calculation engine.** Results excelize cannot compute are left
  for Excel (shown as `?`). An engine of our own that evaluates the common
  functions straight from parsed data would remove the excelize round trip
  on large files. sheets already computes SUM, COUNT, COUNTA, AVERAGE,
  MIN and MAX of ranges itself; the rest goes through excelize. Until
  then, LAMBDA and other uncomputed functions rely on Excel.
- **LibreOffice as a second calculator, and .xls/.ods through it (step
  7).** Not built: LibreOffice is not installed on this machine, and code
  driving it could not be checked against the real program. Installing it
  (`brew install --cask libreoffice`, about 1 GB) is the user's call. The
  design, ready to build:
  - Find `soffice` (`SHEETS_SOFFICE`, the PATH,
    `/Applications/LibreOffice.app/Contents/MacOS/soffice`); without it
    nothing changes.
  - When a save leaves formulas without a result because the calculator
    cannot evaluate them (not for circular references or the time limit),
    render the package with the results we have and convert it with
    `soffice --headless --convert-to xlsx`, in a private profile
    (`-env:UserInstallation=file://…`) whose `registrymodifications.xcu`
    sets Calc's `OOXMLRecalcMode` to always recalculate.
  - Read the converted file's stored values with our reader. Take
    LibreOffice's results only if it agrees with every result we computed
    ourselves (the cross-check the judge already makes against Excel's
    stored results), never for dynamic arrays, and never an error that
    means an unknown function (`#NAME?`). Store them before the stamp is
    computed, so it vouches for them.
  - Legacy .xls (and .ods): convert to .xlsx on the way in with the same
    call. Reading works on the copy; edits write a sibling .xlsx rather
    than the original format.
- **Drawing conditional formats in `render`.** `render` draws cell styles,
  banding and Excel table styles, but not highlight rules, colour scales,
  data bars or icons. Cell-value rules, scales and bars could be drawn from
  the stored values; formula rules need the calculator inside the
  renderer, which is why it was deferred. Agents use `cond --list` to
  check instead.
- **OpenDocument (.ods).** The format seam (`internal/grid`, `internal/book`)
  is there, and CSV now uses it; .ods needs a reader and a writer.
- **Custom house styles.** The presets (`default`, `finance`) are built in.
  They are shaped so a config file (TOML) could override them per project.

## Smaller gaps

- Columns cannot be inserted into or deleted from an Excel table (the
  command refuses). It would need the table's columns, totals formulas and
  structured references rewritten.
- Pivot table sources follow sheet renames only, not row or column edits.
- `write` and `set` keep a typed error code (`#N/A`) as text; Excel makes
  it an error value. Errors already in a file stay errors when copied,
  moved or sorted.
- `rows insert` and `cols insert` leave the new cells unformatted (no
  borders or number formats inside a formatted table; conditional formats
  and banding do reach them). Excel formats them like the row above or the
  column to the left. `fmt` restyles them, and `write --append` already
  copies the look of the rows above. Doing it on insert means both editors
  writing the styles of the line next to the new ones, the streaming one
  while it writes the new rows.
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

## Range operations and CSV (step 5)

- `copy` and `move` carry conditional formats that `sheets cond` made
  (from their record); other conditional formats stay where they are, and
  the command says so. A move within a sheet takes along conditional
  formats lying wholly inside the moved cells; it cannot split one.
- Moving cells into, out of or within an Excel table is refused (`sort`
  works within one). Excel would rewrite the table's structured
  references as cells move in or out.
- `sort` orders text as Excel does for case and accents (é next to e),
  but not Excel's "word sort", which ignores hyphens and apostrophes.
  Sorting by colour, by custom lists, or left to right is not offered.
- A note that moves gets a fresh note box: a custom size, or a note set to
  always show, is reset. A hyperlink over a range moves as a link on its
  first cell.
- CSV files are edited through an .xlsx workbook in memory, so editing a
  large CSV costs what editing an .xlsx of its size does (see streaming).
- `render` draws .xlsx only; a CSV file has no formatting to draw.

## Streaming edits (step 6)

- Streaming covers writing values, inserting and deleting rows and
  appending to a table. Everything else hands the session to the full
  editor, which loads the workbook into excelize. Column edits are the
  next candidate: every formula, merge, rule and table column would move
  sideways, as rows already do.
- Row edits on a sheet holding dynamic arrays (spills) and writes into a
  spill area also go to the full editor, which knows the spill metadata.
- The target was about 2 s and 300 MB for a one-cell edit on a
  200,000-row file. It takes 1.2 s with the parse cache warm and 3 s
  cold, in 0.7 GB: the parsed grid of the whole workbook is held in
  memory. A reader that keeps only the edited sheet's formulas and the
  cells formulas read would bring memory down.
- Excel 2010 data validations (`x14:dataValidation`, written for lists on
  another sheet by Excel 2010) keep their ranges when rows move, in both
  editors. Modern Excel writes ordinary validations, which do move.
- When cut and pasted cells land on a conditional format's range, Excel
  takes the landing cells out of that range; `move` leaves it whole.

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
