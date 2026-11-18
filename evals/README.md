# A/B evaluation

Does `sheets` make agents better at spreadsheet work? Each task is done by
fresh agents in up to six arms:

- `sheets`: the CLI and its agent skill's instructions (what `sheets agent-docs` prints).
- `python`: openpyxl and pandas, the way agents work without it.
- `python-hint`: `python`, plus one line on writing newer functions with
  their `_xlfn.` prefix (`arms/python-hint.md`).
- `skill`: Python with Anthropic's xlsx skill, committed in
  `arms/xlsx-skill` (Apache 2.0; `NOTICE` names the upstream commit), with
  no LibreOffice. `EVAL_BASELINE_SKILL` names another skill's SKILL.md
  instead.
- `skill-lo`: the same skill in its own setting: LibreOffice is there
  (`soffice` on the PATH, no stand-ins), so its `scripts/recalc.py` can
  store formula results (`arms/skill-lo.md`).
- `mcp`: the tools of haris-musa/excel-mcp-server, run with `uvx`.

Both skill arms get the skill's folder (copied to `OUT/skills`) and its
instructions, as Claude Code loads a skill.

Checks grade the results; the report compares pass rates, turns, tokens,
cost and time.

```bash
evals/run.py run OUT                                  # sheets and python, with `claude -p` (needs the CLI logged in)
evals/run.py run OUT --arms sheets python skill mcp --jobs 8
evals/run.py run OUT --tasks 03 09 --model haiku      # some tasks, another model
evals/run.py prepare OUT                              # inputs and prompts only, to run agents another way
evals/run.py check OUT                                # grade OUT/*/*/work, write OUT/report.md
evals/run.py run OUT --tasks 13 --arms sheets --timeout 40  # a longer limit per agent
evals/run.py run OUT --arms sheets python --more      # only the runs not finished yet
evals/run.py run OUT --redo 13/sheets 02/python       # rerun these two
evals/run.py validate OUT --tasks 10                  # check the checks with known solutions
```

`run` builds `sheets` into `OUT/bin`, puts it on the `sheets` arm's PATH
(and keeps it off the others'), and runs each agent in its own folder with
a spending cap (`--budget`, USD per run) and a model (`--model`, default
`sonnet`). An agent still running after 20 minutes (`--timeout MIN`) is
stopped, with its tools, and counts as failed (an MCP server once hung for
three hours).
Agents run without user or project settings, skills or MCP
servers, so each arm gets only what it is given. Each agent's transcript
is kept in `transcript.jsonl` beside its work, and `meta.json` counts its
tool calls (shell commands by program).

A run is finished once its agent has written `meta.json` (a stopped or
failed agent writes one too). `run --more` prepares and runs only the
task/arm folders of `--tasks` and `--arms` without one, and leaves
finished runs untouched: after an interruption, or to add an arm.
`run --redo TASK/ARM …` (a task prefix naming one task, then an arm)
re-prepares and reruns only those runs, keeping the old `meta.json` as
`meta.prev.json`. Either way `check` then grades every run of the tasks,
so the report holds the kept results beside the new ones.

### The silent column

`report.md` counts *silent* failures per arm: a failed run whose final
answer claims success, the defect a user would not notice. `run.claim`
sorts each final answer (`answer.txt`) by this rule:

1. An empty answer, or a run that ended in an error (stopped, out of
   budget), makes no claim.
2. An answer admits failure when it says the work could not be done or
   checked (*could not*, *unable to*, *failed to*, *unfortunately*, *not
   sure*, …) or carries a caveat about the result (*may not update*,
   *until it is opened*, *open the file in Excel*,
   *not recalculated*, *caveat*, …).
3. Otherwise it claims success when it reports work done or gives an
   answer: a success word (*done*, *changed*, *updated*, *saved*,
   *added*, *now*, *total*, …) or a number (a cell reference or
   sheet name such as *B2* or *Sheet1* is not one).
4. Otherwise it makes no claim.

Admitting wins over claiming: an answer that does both is not silent. A
run is silent when it failed (the task's checks) and its answer claims
success. The rule's examples are doctests
(`python3 -m doctest evals/run.py`); it errs on the side of not calling a
run silent.

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
| 13-model-whatif | a downside case through a chained five-year plan, and its 2030 net income, cumulative net income and NPV |
| 14-keep-features | editing a dashboard with sparklines, a text box, Excel 2010 data bars, a logo and a chart, keeping them |
| 15-messy-report | questions over an exported report: two header rows, subtotals, a hidden row, a figure stored as text |
| 16-debug-budget | finding five planted errors in a budget (a short total, a broken row, text for a number, a wrong reference, double counting) |
| 17-large-dedupe | removing a repeated import from 150,000 rows and reporting the corrected totals |
| 18-csv-dialect | a price change in a German ERP export (semicolons, decimal commas, Windows-1252, CRLF), kept in its format |

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

To keep that check, add `solutions/right.py` (a correct solution) and
`solutions/wrong.py` (a plausible mistake the check must catch).
`validate OUT --tasks …` runs each task's setup, then each solution as
`uv run --with openpyxl --with xlsxwriter python SOLUTION WORK_DIR` on a
fresh copy of the inputs (in `OUT/<task>/solution-right` and `solution-wrong`),
with this checkout's `sheets` first on the PATH and in `SHEETS_BIN`. The
solution's stdout is saved as `answer.txt` and graded as an agent's reply.
It prints a line per task and solution, and exits 0 only when every right
solution passes and every wrong one fails (a solution that exits non-zero
counts as a failure of `validate`). Tasks without `solutions/` are
skipped; `10-downstream-values` has the first pair.

## Does the skill trigger?

`triggers.json` holds 20 requests, half of which should make an agent load
the `sheets` skill and half of which are near misses that should not (a
CSV parser to write, Google Sheets, Apple Numbers, a legacy .xls). They
test the skill's description the way
[agentskills.io](https://agentskills.io/skill-creation/optimizing-descriptions)
describes: run each request a few times with the skill installed, count how
often the agent loads it, and change the description when a request lands
on the wrong side.

```bash
evals/triggers.py OUT --reps 2                    # the skill alone
evals/triggers.py OUT --reps 2 --with SKILL_DIR   # beside another skill, such as Anthropic's xlsx
```

Each request runs in a fresh project folder with the skill installed; the
agent is stopped at its first skill, or after four tool calls without
one, so a run costs a few cents. `OUT/triggers.md` gives, per request, how
often each skill was loaded first.


## SpreadsheetBench Verified

`ssb.py` runs the same arms on [SpreadsheetBench Verified](https://huggingface.co/datasets/KAKA22/SpreadsheetBench)
(400 tasks taken from real spreadsheet questions, CC BY-SA 4.0), graded
the way the benchmark's own `evaluation.py` grades:

    curl -L -o ssb.tar.gz https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/main/spreadsheetbench_verified_400.tar.gz
    mkdir ssb && tar -xzf ssb.tar.gz -C ssb
    evals/ssb.py run OUT --data ssb/spreadsheetbench_verified_400 --sample 40 --arms sheets python skill mcp --jobs 6

- Each agent gets the task's starting workbook as `workbook.xlsx`, the
  user's question, and where the answer belongs (`answer_position`), and
  changes the workbook in place.
- Grading recalculates the result in LibreOffice (`soffice` on the PATH),
  with a profile that recomputes every formula on load, then compares the
  answer's cells with the golden file's values: numbers rounded to two
  decimals, dates as day numbers, blanks equal to empty text.
- 23 tasks are left out: their own golden file fails that grading, because
  LibreOffice computes them otherwise than Excel or cannot open them
  (`ssb.UNGRADABLE`).
- The sample is drawn with a fixed seed (`--seed`), in proportion to the
  instruction types, so arms and models are graded on the same tasks.
- The agents see no LibreOffice: `soffice` is a stand-in that reports it
  missing, as on most machines.
- Each conversion takes a LibreOffice profile of its own from a pool, one
  per job: two conversions sharing one fail without saying so.
