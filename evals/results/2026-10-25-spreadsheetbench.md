# SpreadsheetBench Verified, 2026-10-25

## Verdict

On 90 real tasks from SpreadsheetBench Verified, Sonnet with sheets passed
81, more than any other arm. It beat plain Python (71) and Python with
Anthropic's xlsx skill (72), both by a margin that holds up statistically.
It tied the Excel MCP server (80) at a little over half the cost and 40%
of the turns. Five tasks fail in every arm because the task itself is
flawed. Of the 85 sound tasks, sheets failed four: two through tool bugs,
both now fixed, and two through the agent's own choices.

Python's and the skill's largest single loss comes from openpyxl, not the
agents: it writes newer functions (AGGREGATE, TEXTJOIN, MAXIFS) without
the `_xlfn.` prefix the file format requires. That costs 9 tasks in the
python arm and 7 in the skill arm. `sheets lint` now finds this in
files a script wrote.

## How it was tested

- **Tasks:** SpreadsheetBench Verified holds 400 tasks; the 23 whose own
  golden file fails the grader are left out (`UNGRADABLE` in
  `evals/ssb.py`). The sample is 90 tasks, stratified by kind (54
  cell-level, 36 sheet-level), drawn with seed 2026.
- **Runs:** one `claude -p` run per task and arm, on Sonnet, with a $3
  budget per run, isolated from user settings, skills and MCP servers.
  These are the four arms of `evals/run.py`:
  - **sheets:** the CLI, with `sheets agent-docs` in the prompt.
  - **python:** openpyxl and pandas through `uv run`.
  - **skill:** the same, plus Anthropic's xlsx skill.
  - **mcp:** haris-musa/excel-mcp-server 2.2.2.

  LibreOffice is installed for grading, but it is hidden from the agents,
  as it is missing on most machines.
- **Grading:** done as SpreadsheetBench grades. LibreOffice recalculates
  the result, then openpyxl compares the answer range with the golden
  file: numbers rounded to two places, datetimes as day serials, and
  empty equal to "". Run it with `python3 -I evals/ssb.py run|check`.

## Results

| arm | passed (90) | sound tasks (85) | cell-level | sheet-level | turns | cost $ | median s | tokens |
|---|---|---|---|---|---|---|---|---|
| sheets | **81** | **81** | 49/54 | 32/36 | 363 | 6.30 | 17 | 15.1M |
| mcp | 80 | 80 | 47/54 | 33/36 | 893 | 11.17 | 24 | 31.1M |
| skill | 72 | 72 | 40/54 | 32/36 | 401 | 7.37 | 22 | 17.3M |
| python | 71 | 71 | 41/54 | 30/36 | 300 | 4.91 | 17 | 11.8M |

Paired on the 85 sound tasks:

| comparison | sheets alone passed | the other alone passed | sign test |
|---|---|---|---|
| sheets vs python | 14 | 4 | p = 0.03 |
| sheets vs skill | 12 | 3 | p = 0.04 |
| sheets vs mcp | 4 | 3 | p = 1.0 |

## Why tasks failed

| cause | sheets | python | skill | mcp |
|---|---|---|---|---|
| flawed task (fails in every arm) | 5 | 5 | 5 | 5 |
| function written without `_xlfn.` (openpyxl) | – | 9 | 7 | – |
| range formula stored for older Excel | – | 1 | – | – |
| tool bug or hang | 2 | – | – | 1 |
| the agent's logic or choice | 2 | 4 | 6 | 4 |
| **failed** | **9** | **19** | **18** | **10** |

**Flawed tasks.** Each of these fails in every arm, for a reason in the
task itself:

- **13284:** the task says names must "exactly match", but the golden
  answer matches "Braz de Pina" to "Braz De Pina".
- **146-49:** the task does not say which column of a word pair holds
  which word; all four arms chose the same order.
- **45300:** the golden answer keeps a trailing space.
- **54667:** the golden answer returns 0:00 for rows whose two name cells
  are both blank, treating them as "the same".
- **80-42:** the sheet already holds the consolidated rows, and the
  golden answer appends a second copy.

**Functions without `_xlfn.`.** Excel stores functions added after 2007
with a prefix (`_xlfn.AGGREGATE`). openpyxl writes formulas as typed;
AGGREGATE is not in its list of known functions, and its documentation
says such formulas must carry the prefix to work. LibreOffice gives
`#VALUE!` or `#NAME?` for them.

- **Affected tasks:** the INDEX/AGGREGATE(15,6,…) lookup in tasks 10452,
  17-35, 1818, 41410, 50971, 56378 and 58499; TEXTJOIN in 50768; MAXIFS
  in 516-46.
- **Cause confirmed:** the same AGGREGATE formula computes in LibreOffice
  once it has the prefix, whether stored as a plain formula or an array
  formula.
- **sheets and the MCP server:** both write the prefixes themselves.

An earlier note (chunk 1) called the AGGREGATE failures a grader
artifact. They are not: the files themselves are wrong.

sheets now catches this in files other tools write.

- **Reading:** it gives such a formula, and whatever reads it, an
  unknown value (`?`), where it used to compute a value Excel does not
  show.
- **Lint:** the new `unprefixed-function` rule reports them, with the
  fix (`sheets set` on the cell adds the prefix).
- **Checked:** it flags exactly the 9 python and 7 skill outputs above,
  and none of the other 344 outputs or the corpus's 800 files.

**Range formula for older Excel.** In 51090, python stored
`SUM(SUMIFS(…, $G3:$J3))` as a plain formula. LibreOffice then
intersects `$G3:$J3` with the formula's own column and gives `#VALUE!`;
Excel 365 reads it the same way, as written for older Excel. sheets stores
such a formula as a dynamic array, which computes over the whole range.
The skill arm used SUMPRODUCT, which works either way.

**sheets' four.**

- **191-40 (tool bug, fixed).** Unsharing filled-down formulas made
  excelize mark every rewritten formula's result as text. Unsharing
  happens on clear, copy, move, sort and script edits.
  - **What went wrong:** in this task, the agent's batch of nine block
    sorts read "3.7" as text after the first sort and sorted one block
    as text.
  - **Wider impact:** every reader of the file (pandas, openpyxl) would
    also have read about 150 numbers in this workbook as text. Excel and
    LibreOffice hide this by recalculating on open, which is why no eval
    caught it until now. 78 of the 405 task inputs use shared formulas.
  - **Fix:** the results of unshared formulas are stored again with
    their own types (`TestUnshareKeepsResults`). Replaying the agent's
    commands now gives the golden answer.
- **50051 (tool bug, fixed during the run).** A comparison against a
  whole column taken with `INDEX(…,,col)` was not stored as an array
  formula. Replaying the agent's commands passes with the fix.
- **203-15:** the agent named a column "Transport Allowance", as the task
  text does; the data and the golden answer say "Transport Allowances".
- **55977:** the agent wrote "Monday" where "Mon" was expected.

**The MCP server's hang.** On 50193 the excel-mcp-server process sat at
98% CPU for three hours; the run was stopped and counts as a failure.
The 248 minutes in the table include it; without it the arm took 58.
`evals/run.py` now stops an agent and its tools after 20 minutes.

## Grader fixes made during this run

- **Shared profiles:** `check` gave each LibreOffice conversion a profile
  by position, but the thread pool runs jobs in any order. Two
  conversions could then share a profile, and one would silently fail,
  so results changed from run to run. Now each conversion takes a
  profile of its own from a pool and retries once. Two full gradings now
  agree exactly.
- **Timeouts:** a conversion that times out takes LibreOffice's own
  process down with it, so the profile is not left locked.

## Limits

- **Sample size:** one run per task, on one model. The sheets–mcp gap is
  within noise; the gaps to python and the skill are not, but rest on 18
  and 15 discordant tasks.
- **Grading through LibreOffice:** LibreOffice is not Excel. Every arm
  was graded the same way, and the prefix failures were checked by hand.
- **The skill arm without LibreOffice:** the skill's instructions check
  results with LibreOffice's `recalc.py`, which these agents could not
  run. With it, they might have seen the `#VALUE!` results and fixed
  them.
