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

# Results

Run 2026-11-18, after the pre-registration above, on commit `873d96b`
(sheets after piece A: trust-engine and kinds-spills; without D's change
report). A rerun with D's per-edit report is for the next wave.

**Verdict:** fail. sheets' defect rate is 5/201 (2.5%); the best Python
arm is `skill-lo` at 7/201 (3.5%), and half of that is 1.74%, so the
first condition fails. The second holds: sheets cost $1.27 against
`skill-lo`'s $1.85 (0.69 times; the bar is 1.3).

## How the runs went

- Three batches started together, each
  `evals/run.py run ~/sheets-evals/wave2/b/runN --tasks 19 20 --arms sheets python skill-lo --model claude-sonnet-5-5 --timeout 40 --budget 10 --jobs 6`
  (18 agents at once). Each run is kept under `~/sheets-evals/wave2/b/run1`,
  `run2` and `run3`, as `TASK/ARM`, with its `transcript.jsonl`,
  `meta.json`, `answer.txt` and `result.json`; `analyze.py` there tallies
  them for the tables below.
- The folders are `b/run1` to `b/run3`, not `b-runs/r1` to `r3` as the
  pre-registration names them. The spot-check's draw keeps the
  pre-registration's names: `r1/…` is `run1/…`.
- All 18 started and ended on their own (result `success`; the longest
  took 100 s). None was stopped by `--timeout`, none reached the $10 cap
  (the dearest cost $0.51), and no run failed to start. So nothing was
  rerun with `--redo`, and no batch needed `--more`.
- Every graded run has all its task items (43 or 24), so no run was
  counted as all defects.

## The table

Defects are failed `task` items; stored items are apart, outside the bar.
Silent (rule) is `report.md`'s column, from `run.claim`; silent (read) is
from reading each failed run's answer (below).

| run | task | arm | defects | stored | silent (rule) | silent (read) | turns | cost $ |
|---|---|---|---|---|---|---|---|---|
| run1 | 19 | sheets | 1/43 | 3/3 | no | yes | 18 | 0.29 |
| run1 | 19 | python | 2/43 | 0/3 | no | yes | 4 | 0.14 |
| run1 | 19 | skill-lo | 0/43 | 3/3 | | | 7 | 0.36 |
| run1 | 20 | sheets | 1/24 | 2/2 | yes | yes | 9 | 0.13 |
| run1 | 20 | python | 0/24 | 2/2 | | | 6 | 0.16 |
| run1 | 20 | skill-lo | 4/24 | 2/2 | yes | yes | 5 | 0.11 |
| run2 | 19 | sheets | 1/43 | 3/3 | no | yes | 21 | 0.29 |
| run2 | 19 | python | 1/43 | 0/3 | no | yes | 4 | 0.14 |
| run2 | 19 | skill-lo | 0/43 | 3/3 | | | 12 | 0.42 |
| run2 | 20 | sheets | 0/24 | 2/2 | | | 10 | 0.13 |
| run2 | 20 | python | 6/24 | 0/2 | no | yes | 4 | 0.08 |
| run2 | 20 | skill-lo | 3/24 | 2/2 | yes | yes | 5 | 0.11 |
| run3 | 19 | sheets | 1/43 | 3/3 | yes | no | 22 | 0.33 |
| run3 | 19 | python | 0/43 | 3/3 | | | 6 | 0.27 |
| run3 | 19 | skill-lo | 0/43 | 3/3 | | | 8 | 0.51 |
| run3 | 20 | sheets | 1/24 | 2/2 | yes | yes | 7 | 0.10 |
| run3 | 20 | python | 2/24 | 0/2 | no | yes | 4 | 0.21 |
| run3 | 20 | skill-lo | 0/24 | 2/2 | | | 8 | 0.34 |

| arm | defects | rate | task 19 | task 20 | stored | silent (rule) | silent (read) | turns | cost $ | capped |
|---|---|---|---|---|---|---|---|---|---|---|
| sheets | 5/201 | 2.5% | 3/129 (2.3%) | 2/72 (2.8%) | 15/15 | 3 of 5 | 4 of 5 | 87 | 1.27 | none |
| python | 11/201 | 5.5% | 3/129 (2.3%) | 8/72 (11.1%) | 5/15 | 0 of 4 | 4 of 4 | 28 | 1.00 | none |
| skill-lo | 7/201 | 3.5% | 0/129 (0.0%) | 7/72 (9.7%) | 15/15 | 2 of 2 | 2 of 2 | 45 | 1.85 | none |

Silent counts are of the arm's failed runs. sheets failed 5 runs of 6 by
one item each; `python` 4 and `skill-lo` 2, by more.

**skill-lo's recalc retries** (eval-arms/f4.2). No agent here ran a bare
`python`. Each first ran the skill's `recalc.py` with `python3`, which on
this arm's PATH is the system's Python 3.9. That fails inside recalc.py
(`TemporaryDirectory(ignore_cleanup_errors=…)` is 3.10's), and the agent
retries with `uv run --with openpyxl python`. That happened once in four
runs (run1/20, run2/20, run3/20, run2/19; in run2/19 the same call also
failed in the agent's own script). Each such call's share of its run's cost is taken
from its tokens, priced in proportion:

| skill-lo | turns | cost $ |
|---|---|---|
| as run | 45 | 1.85 |
| without the 4 retried calls | 41 | 1.75 |

Without them the cost condition is 1.27 against 1.75 (0.72 times); the
verdict does not change.

## Failure categories

- **New rows without their neighbours' number formats** (all 5 of
  sheets' defects). Task 19, all three runs: `write Products --append` and
  `write Orders --append` add the table rows with General in the number
  columns (Gizmo Pro's price shows 48, its margin 0.5625). In run3 the
  agent then set Products' whole margin column to 0.0% to make it match,
  and said so. Task 20, run1 and run3: `rows insert Budget!9:9` leaves
  the new row in General where Excel takes the formats of the row above.
  Both agents restyled it with `style format=number`, which writes the
  house format `#,##0_);(#,##0)`, not the sheet's `#,##0`. (run2 matched
  the sheet's own formats and passed.) These are sheets' gaps, not the
  checks': a fresh `rows insert` on the task's input leaves row 9 in
  General.
- **Dynamic arrays saved as legacy arrays** (python run1/19 and run2/19):
  openpyxl writes `t="array"` without `cm="1"` or `metadata.xml`.
- **A conditional format not extended to a new table row** (python
  run1/19): Products' colour scale stays on E2:E7 after the table grows
  to row 8.
- **Drawings and notes lost or left behind by openpyxl's row insert**
  (python run2/20 and run3/20, skill-lo run1/20 and run2/20): the chart's
  series stop at row 9 and it does not move down, the rounded rectangle
  is dropped, and HR's note stays on A9 (Legal's row).
- **Formulas not moved by the insert** (python run2/20): HR's total,
  now in row 10, still reads row 9 (`=SUM(C9:F9)`, Legal's), so HR shows
  169,000 and the grand total 3,105,000, for 186,000 and 3,122,000. Its answer
  reported 3,122,000, worked out by hand.
- **A wrong number in the reply** (skill-lo run1/20): Engineering's
  share given as 22.9% (715,000 / 3,122,000), for 42.8%.
- **Stopped, timed out, capped or failed to start:** none.

Stored results, apart from the bar: sheets and `skill-lo` stored every
checked result (15/15); `python` stored 5 of 15. Two of its six runs
wrote values themselves. The others left none.

Seen in the transcripts, not graded: every task-19 sheets run edited the
package's XML by hand for edit 18 (a named cell) and edit 19 (a column
added to an Excel table), because no sheets command does either. All
three sheets runs left `Summary!B19` with no stored result, which sheets'
calculator could not place (a reference into Products' table), and said
so.

## Silent defects, by reading

`run.claim` still calls some plain success reports admissions
(eval-arms/f4.1). So each failed run's final answer was read, and it
counts as silent when it claims the work done and nothing in it names or
puts a caveat on any of its failed items. Caveats about something else,
such as no recalculation or a hand edit of the XML, do not count. The two
counts differ on 7 of the 11 failed runs:

- **The rule says admitted, the reading says silent** (6). sheets run1/19
  and run2/19 caveat the XML hand edits and `Summary!B19`. python run1/19,
  run2/19, run2/20 and run3/20 caveat that LibreOffice was missing or the
  file was not recalculated. None mentions its own defects (formats,
  arrays, colour scale, chart, shape, note, totals). python run1/19 even
  says the conditional formats "still cover the same ranges as before",
  which is its defect.
- **The rule says silent, the reading says admitted** (1). sheets run3/19
  says it set Products' margin column to 0.0% unasked, one of the two
  causes of its failed item.

By reading, sheets is silent in 4 of its 5 failed runs, `python` in 4 of
4 and `skill-lo` in 2 of 2: almost every failure here was reported as
success.

## Spot-check

The draw, `random.Random(20261118).sample(sorted(folders), 10)`, gave
r2/19/python, r2/20/sheets, r1/19/skill-lo, r1/19/python, r3/19/python,
r3/19/skill-lo, r1/20/python, r3/20/skill-lo, r1/20/skill-lo and
r2/19/sheets. Each transcript and answer was read, and the failed items
were checked against the workbooks, opened with openpyxl and read as XML.

**Spot-check:** 10 transcripts read; the 8 failed items in them are all
real defects (r1/19/python 2, r2/19/python 1, r1/20/skill-lo 4,
r2/19/sheets 1). The silent verdict agrees with the reading in 7 of 10.
The 3 disagreements, r1/19/python, r2/19/python and r2/19/sheets, are
`run.claim`'s f4.1: the rule counts a caveat about something else as an
admission. No task check was wrong, so none was changed and no run was
regraded. The rule itself is `evals/run.py`'s, outside this piece. The 6
passing runs' workbooks pass; one prose slip passed ungraded:
r3/19/skill-lo says five orders pay shipping, where its own 36.00 is
eight.

## Caveats

- Three runs per cell. sheets' 5 defects against `skill-lo`'s 7 are
  close in raw counts. The bar, applied as written, still fails.
- sheets' five defects are all one category (new rows' number formats).
  If `write --append` into a table and `rows insert` carried the
  neighbours' number formats, as Excel does, the agents would likely not
  have restyled anything. That is a finding for the next wave, not a
  regrade.
