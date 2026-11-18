# Wave 2, piece B: does sheets keep a living workbook whole? Pre-registration

Written 2026-11-18, before any agent has run on tasks 19 and 20. This
section is not edited after the first run; the results go in a section
below it, with the commit the runs measured. Nothing here is a result.

## The claim

An agent with `sheets` edits a feature-rich workbook, over many edits,
with fewer defects than the best agent that writes Python, at about the
same cost, and when it fails, does not say it succeeded.

## What is measured

- **Tasks** (`evals/tasks`, described in `evals/README.md`):
  - `19-living-workbook`: one prompt, 20 ordered edits to a five-sheet
    workbook (two Excel tables, named cells, a chart, conditional formats,
    validations, cross-sheet formulas, four dynamic arrays). 43 task
    items, 3 stored items.
  - `20-feature-loss`: six edits (one a row inserted mid-table) to a
    budget with a chart, a shape, four kinds of validation, five kinds of
    conditional format, a note and frozen panes. 24 task items, 2 stored
    items.
- **Grading** is independent of sheets. Formula values come only from
  LibreOffice (`evalkit.lo_values`, LibreOffice 26.8.3.1); expected values
  are computed in Python from the inputs; features are read from the
  package's XML or with openpyxl. Each task passed `evals/run.py validate`
  (its right solution passes, its wrong one fails), its untouched inputs
  fail it, and a value planted by hand in a right solution's output is
  caught.
- **Arms**: `sheets`, `python`, `skill-lo` (Anthropic's xlsx skill, Apache
  2.0, with LibreOffice on the PATH). The Python arms are `python` and
  `skill-lo`.
- **Runs**: three per arm and task, so 18 runs, each in its own output
  folder under `~/sheets-evals/wave2/` (`b-runs/r1`, `r2`, `r3`), with
  `claude-sonnet-5-5`, `--timeout 40` and `--budget 10`:
  `evals/run.py run OUT --tasks 19 20 --arms sheets python skill-lo --model claude-sonnet-5-5 --timeout 40 --budget 10`.
- **Build under test**: sheets as of piece A (trust-engine and
  kinds-spills), without D's change report. The commit is recorded with
  the results. A rerun with D is for the next wave.

## Counts

- A run's **defects** are its failed `task` check items. A run that left
  no workbook to grade, or whose agent was stopped before it saved one,
  has every task item of its task as a defect.
- An arm's **defect rate** is its defects divided by task items times
  runs: both tasks pooled, `(43 + 24) * 3 = 201` per arm. Per-task rates
  and the raw counts are reported beside it.
- **Stored** items (does the file store the results, as pandas and
  previews read them) are reported in their own column, outside the bar,
  because piece C tests stored results.
- **Cost** is what the claude CLI reports in USD per run (`cost_usd` in
  `meta.json`), summed over the arm's six runs.
- **Silent defects** are the report's `silent` column: failed runs whose
  final answer claims success, by the rule in `evals/README.md`. The
  count is reported beside the defect rate; it is not part of the bar.

## The bar

sheets passes if, over the 18 runs:

1. its defect rate is at most half the defect rate of the best Python arm
   (the lower of `python` and `skill-lo`; on a tie the cheaper), and
2. its cost is at most 1.3 times that arm's cost.

A best Python arm with no defects makes the first condition mean none for
sheets too. The result is mechanical: a verdict, not an argument. It is
reported to Max and written in the results file; it never cuts a piece.

## The spot-check

Before the results are read, 10 of the 18 run folders are drawn with
`random.Random(20261118).sample(sorted(folders), 10)`, where `folders` are
the `TASK/ARM` runs of the three output folders as `r1/19-living-workbook/sheets`
and so on. For each, the transcript is read against the failed items:
is each failure a real defect, and is each silent verdict right? A check
found wrong by it is corrected, every run is regraded, and the
correction and its reason are listed in the results. The count of items
read and judged real is reported.

## Known before the runs

- Checks overlap: one slip can fail an edit's item and its sheet's item.
  This counts the same for every arm.
- A tool that splits a validation or a conditional format around an
  inserted row passes if every row stays covered. A new row must look
  like its neighbours (number formats), and a table must take its new
  column and rows.
- Task 19's four dynamic arrays are graded on being dynamic arrays
  (`cm` attribute and their spill range) and on their values. At the
  commit this was written on, sheets saves each as a one-cell array that
  Excel would show as `#SPILL!`; `solutions/right.py` repairs that by
  hand. If piece A has not fixed it by the runs, that item fails for
  sheets, and the results say so.
- Three runs per cell is a small sample and no significance test is run.
  The verdict is still the bar's, applied as written to the counts: there
  is no second threshold on the size of a difference. The raw counts are
  reported beside the rates so the reader can weigh them.
