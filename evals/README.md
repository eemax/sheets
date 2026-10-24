# A/B evaluation

Does `sheets` make agents better at spreadsheet work? Each task is done by
fresh agents in up to four arms:

- `sheets`: the CLI and its agent skill's instructions (what `sheets agent-docs` prints).
- `python`: openpyxl and pandas, the way agents work without it.
- `skill`: Python with a spreadsheet skill. `EVAL_BASELINE_SKILL` names its SKILL.md, for example Anthropic's xlsx skill.
- `mcp`: the tools of haris-musa/excel-mcp-server, run with `uvx`.

Checks grade the results; the report compares pass rates, turns, tokens,
cost and time.

```bash
evals/run.py run OUT                                  # sheets and python, with `claude -p` (needs the CLI logged in)
evals/run.py run OUT --arms sheets python skill mcp --jobs 8
evals/run.py run OUT --tasks 03 09 --model haiku      # some tasks, another model
evals/run.py prepare OUT                              # inputs and prompts only, to run agents another way
evals/run.py check OUT                                # grade OUT/*/*/work, write OUT/report.md
```

`run` builds `sheets` into `OUT/bin`, puts it on the `sheets` arm's PATH
(and keeps it off the others'), and runs each agent in its own folder with
a spending cap (`--budget`, USD per run) and a model (`--model`, default
`sonnet`). Agents run without user or project settings, skills or MCP
servers, so each arm gets only what it is given. Each agent's transcript
is kept in `transcript.jsonl` beside its work, and `meta.json` counts its
tool calls (shell commands by program).

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

