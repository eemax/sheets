# A/B evaluation

Does `sheets` make agents better at spreadsheet work? Each task is done by
fresh agents twice: once with the `sheets` CLI and its guide, once with
Python (openpyxl, pandas), the way agents work without it. Checks grade
the results; the report compares pass rates, turns, tokens and cost.

```bash
evals/run.py run OUT                 # everything, with `claude -p` (needs the CLI logged in)
evals/run.py run OUT --tasks 03 09   # some tasks
evals/run.py prepare OUT             # inputs and prompts only, to run agents another way
evals/run.py check OUT               # grade OUT/*/*/work, write OUT/report.md
```

`run` builds `sheets` into `OUT/bin`, puts it on the `sheets` arm's PATH
(and keeps it off the other arm's), and runs each agent in its own folder
with a spending cap (`--budget`, USD per run) and a model (`--model`,
default `sonnet`). Set `EVAL_BASELINE_SKILL` to a file of instructions to
give the Python arm the spreadsheet skill agents normally have.

## Tasks

| task | what it tests |
|---|---|
| 01-understand-report | reading a report with a title, a formula column with no stored results, a total row and a note |
| 02-large-question | answering a question over 30,000 rows |
| 03-append-totals | adding rows above a total row, keeping the totals and the chart right |
| 04-insert-row | inserting a row in a model whose other sheet reads single cells |
| 05-report-from-csv | building a formatted report with totals and a chart from a CSV |
| 06-format-request | applying specific formatting a user asks for |
| 07-fix-errors | finding and fixing planted mistakes |
| 08-rename-sheet | renaming a sheet that formulas and a name refer to |
| 09-table-append | adding a row to an Excel table with a totals row |
| 10-downstream-values | changing an input of a workbook Excel saved, with results stored |
| 11-style-requests | five styling requests at once: currency with the total, whole-row highlights by a rule, a dropdown, a header, frozen panes |
| 12-build-report | a two-sheet report from two CSVs in many steps: formulas across sheets, money formats, a red-when-negative rule, frozen headers |

Each task folder has `prompt.md` (the request, the same for both arms),
`setup.py` (writes the inputs, and the true answers to a file the agent
does not see) and `check.py` (grades the result with `evals/lib/evalkit.py`).

Checks judge values by computing formulas (with `sheets read`), so a
correct formula counts whoever wrote it. A separate kind of check, *stored
results*, asks whether the file stores formula results, as pandas and
previews read them; it is reported apart from the task's pass or fail.

## Adding a task

Write the three files; then check the check both ways before trusting it:
untouched inputs must fail (`prepare`, then `check`), and a correct
solution, by either tool, must pass.
