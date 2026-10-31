# Roadmap

Where sheets is going, and why, is in [vision.md](vision.md). How the work
is run (workers, audits, merging) is in [orchestration.md](orchestration.md).
This file is the plan: a status board, then Horizon 1 as work packages.

## Board

| Package | What | Wave | Status |
|---|---|---|---|
| WP-01 | `sheets check` | 1 | planned |
| WP-07a | Excel check, round 2: what sheets writes today | 1 | planned |
| WP-08 | Evals that measure the claims | 1 | planned |
| WP-02 | Experiment: does `check` cut Python's silent defects? | 2 | planned |
| WP-05 | `explain` and `trace` | 2 | planned |
| WP-06 | `what-if` | 2 | planned |
| WP-04 | Expectations the workbook carries | 3 | planned |
| WP-03 | Values that moved where no edit reached | 3 | planned |
| WP-09 | Receipts | 4 | planned |
| WP-10 | Skill and docs for the new position | 4 | planned |
| WP-07b | Excel check: the new features; then Max opens the pack in Excel | 4 | planned |
| WP-11 | An Excel oracle: real Excel computing, through Microsoft Graph | — | waiting on Max |

**Gate after wave 2.** WP-02's verdict decides whether the position holds
(see the kill criterion in the vision). Max reviews it before wave 3.

Statuses: `planned`, `working`, `auditing`, `fixing`, `done`, `parked`.

## Horizon 1: work packages

Each package is one worker session, then an independent audit. Its
acceptance criteria are what the auditor checks. Every package also meets
the definition of done in [orchestration.md](orchestration.md).

### WP-01 `sheets check`

- **Goal:** one command makes any workbook ready to hand over, whatever
  wrote it.
- **Why:** this is the core of "Python writes, sheets proves", and the
  wedge into workflows agents already use.
- **Behaviour.** This shape is a recommendation; the worker may refine it
  within the principles.
  - `sheets check book.xlsx` reports, and writes nothing. It lists:
    - formulas with no stored result, or a stale one;
    - values sheets cannot know, with the reasons;
    - formula errors;
    - functions stored without their prefix;
    - lint's findings.

    It exits 1 on errors, as `lint` does.
  - `sheets check --fix book.xlsx` stores every result sheets can compute
    and adds missing `_xlfn.` prefixes, as one undo step. It says what it
    changed.
    - **Results it computes:** missing and stale ones.
    - **Excel's results:** a result Excel stored and still vouches for is
      never replaced; a disagreement is a note.
    - **Repeated runs:** a second run changes nothing.
    - **Everything else:** stays as it was.
  - `--json` and `--dry-run`, as other commands have.
  - On a CSV file, it runs lint only.
- **Acceptance:**
  - **Files other tools wrote:** workbooks written by openpyxl and by
    XlsxWriter, after `--fix`:
    - openpyxl (`data_only`) and pandas read the computed values;
    - LibreOffice's recalculation agrees on every cell;
    - functions without their prefix get it.
  - **The corpus:** across the 800 workbooks in
    `~/sheets-evals/data/ssb`, `--fix`:
    - corrupts none (sheets and LibreOffice open every result);
    - changes no result Excel stored and vouched for;
    - is followed by a second run that changes nothing.
  - **A large file:** a 150,000-row workbook finishes; record the time.
  - **Tests:** txtar and unit tests, passing in streaming and full modes.
  - **Docs:** the check guide, and the skill regenerated.
- **Not in scope:** expectations (WP-04), receipts (WP-09), values that
  moved (WP-03).

### WP-02 Experiment: does `check` cut Python's silent defects?

- **Goal:** test the position against the kill criterion in the vision.
- **Sample:** the 90 SpreadsheetBench tasks from 2026-10-25. That is the
  union of `ssb.py --sample 40` and `--sample 80`, both with seed 2026.
- **Arms:**
  - **python:** run again.
  - **python + check afterwards:** `sheets check --fix` applied to each
    python output before grading.
  - **python + check in the prompt:** the agent is told to finish with
    `sheets check --fix`.
  - **skill + check afterwards.**

  Use the 2026-10-25 sheets arm as the reference, and say so.
- **Silent defects:** a failed run whose final answer reports success
  (see WP-08).
- **Acceptance:**
  - **Results:** a results document in `evals/results`, with paired
    statistics, failure categories and an explicit answer to the kill
    criterion.
  - **Harness:** the harness changes, committed.
  - **Run outputs:** kept under `~/sheets-evals/runs/`.
- **Depends on:** WP-01; WP-08's silent-defect column, if it is ready.

### WP-03 Values that moved where no edit reached

- **Goal:** `sheets diff old.xlsx new.xlsx` explains value changes:
  - the inputs that changed;
  - the formulas that changed;
  - the computed values that moved because of them;
  - values that moved with no edit upstream, flagged.
- **Why:** this is a check that needs no authoring. An unexplained change
  is how a broken formula, a lost result or a bad round trip shows itself.
- **Acceptance:** planted cases in tests:
  - **an input change:** every moved value is explained;
  - **one row's formula broken:** flagged;
  - **an openpyxl round trip that lost results:** flagged;
  - **a sort:** reported as moved rows, not changes.

  Output stays bounded on large sheets.
- **Depends on:** diff's row and column alignment, and the engine's
  dependencies.

### WP-04 Expectations the workbook carries

- **Goal:** `sheets expect book.xlsx '<name>' '<condition>'`, for example
  `'Summary!B5=SUM(Data!F:F)'`.
  - **Storage:** expectations are kept in the workbook as hidden defined
    names, as style records are. References then follow every edit, and
    the expectation travels with the file.
  - **Commands:** `--list` and `--delete`.
  - **Checking:** `sheets check` evaluates them all.
  - **Suggestions:** `expect --suggest` proposes inferred ones for the
    agent to accept, such as total rows equal to their column's sum, or
    totals that tie across sheets.
- **Acceptance:**
  - **Edits:** expectations survive row inserts and deletes, sorts and
    sheet renames.
  - **Excel:** the names are valid and hidden, so Excel accepts them.
  - **A failure:** `check` exits 1 and shows the values involved.
  - **Unknown:** an expectation sheets cannot evaluate is reported as
    unknown, never as passing.
  - **Tests and docs.**
- **Depends on:** WP-01.

### WP-05 `explain` and `trace`

- **Goal:** `sheets explain book.xlsx Summary!B5` shows the formula and
  each input with its value, recursively to a depth.
  - **Large ranges:** summarized (count, sum, smallest and largest), not
    listed.
  - **Names and table references:** resolved.
  - **Cycles:** safe.
  - **`--dependents`:** shows what reads the cell.
- **Acceptance:**
  - **Correctness:** correct across sheets, names, Excel tables, dynamic
    arrays (`D2#`) and unknown values.
  - **Size:** bounded, about 60 lines by default.
  - **Output and checks:** `--json`, tests on fixtures, and docs.

### WP-06 `what-if`

- **Goal:** `sheets whatif book.xlsx 'Inputs!B2=1.1' --show 'Summary!B5:B9'`
  shows values before and after, and saves nothing.
  - **Scenarios:** several, side by side.
  - **Unknown values:** what sheets cannot compute shows as unknown.
  - **Stretch goal:** `--seek` finds the input value that brings a target
    to a given number.
- **Acceptance:**
  - **Nothing saved:** the file's bytes are unchanged and no undo step is
    recorded.
  - **Correctness:** the results equal those of a real edit followed by
    recomputation.
  - **Tests and docs.**

### WP-07a Excel check, round 2

- **Goal:** add to `scripts/excel-check.sh` and `excel-check/CHECKLIST.md`
  what sheets writes since the last round:
  - dynamic arrays with spilled values, and `#SPILL!`;
  - `D2#` references;
  - a sort over shared formulas;
  - a workbook written by `sheets script`;
  - an openpyxl file with unprefixed functions after `sheets set`.
- **Acceptance:**
  - **The script:** runs clean.
  - **The checklist:** each item says exactly what to look for (opens
    without a repair prompt; F2:F4 shows a, b, c; editing A4 grows the
    spill).

### WP-07b Excel check: the new features

- **Goal:** the pack gains a file after `check --fix`, expectations and
  receipts. Max then opens the whole pack in Excel, and his findings
  become packages.
- **Depends on:** WP-01, WP-04, WP-09.

### WP-08 Evals that measure the claims

- **Goal:** the measures in the vision exist and run.
  - **Silent defects:** a column in the reports of `run.py` and `ssb.py`
    for failed runs whose final answer reports success. A simple
    classifier is fine, if its rule is written down and spot-checked.
  - **A long-horizon task:** a realistic multi-sheet workbook (tables,
    names, a chart, conditional formats) and 20 edits in one session.
    It is graded on the end state: every formula's value against an
    independent computation, and every feature intact.
  - **A model-review task:** about 10 planted errors in a model of about
    10 sheets.
  - **Two reference workflows, on synthetic data:** reading a supplier's
    sheet into a standard; and a costing template filled for 50 products.
- **Acceptance:**
  - **Conventions:** the tasks follow `evals/tasks` (`setup.py`,
    `check.py`, `prompt.md`).
  - **The checks:** validated against one hand-made correct solution and
    one hand-made wrong one.
  - **A dry run:** one arm runs once.

### WP-09 Receipts

- **Goal:** `sheets check --receipt out.md` (or HTML) writes a short page
  for a person. It shows:
  - what changed since a version, from the undo history or
    `--against old.xlsx`;
  - the values that moved, and why;
  - the checks and expectations;
  - previews of the changed sheets;
  - what is unknown.
- **Depends on:** WP-01, WP-03, WP-04.

### WP-10 Skill and docs for the new position

- **Goal:** the skill, `agent-docs` and the README teach:
  - when to use sheets and when Python;
  - to finish with `sheets check`;
  - expectations, `explain` and `what-if`.
- **Acceptance:** the trigger eval (`evals/triggers.py`) still chooses the
  skill when it should.
- **Depends on:** WP-01, WP-04, WP-05, WP-06.

### WP-11 An Excel oracle (waiting on Max)

- **Goal:** test sheets against Excel itself, not LibreOffice. Microsoft
  Graph's workbook API opens a file in Excel for the web, recalculates it
  with Excel's engine, and returns values. Three things would then be
  checked automatically:
  - the engine's results;
  - the files sheets writes (Excel opens them without repairing them);
  - SpreadsheetBench grading, without LibreOffice's gaps.
- **Needs:** a Microsoft 365 account Max chooses, with its token in `sec`.
  Nothing starts until he says which account.
- **Acceptance:** a script uploads a workbook, recalculates, and compares
  every formula's value with what sheets stores. It runs on the
  excel-check pack and a corpus sample, and reports where they differ.

## Horizon 2, to be cut into packages after the gate

- **Session mode:** `sheets serve`, a long-lived process holding a workbook,
  talking over stdio.
- **A Python client on top of it:** the command set as functions plus live
  reads, not a second object model.
- **Templates:** fill a workbook template once per record.
- **History and review:** an agent proposes a change, a person sees it
  rendered and accepts it.

## Done

- **2026-10-23 to 25:** reading, writing, the house style, query, lint,
  diff, charts, render, batches, range operations, CSV, streaming edits,
  the agent skill, and a formula engine of its own. Also `sheets script`,
  dynamic arrays and the SpreadsheetBench harness. The details are in git
  history (up to 91b697f) and `evals/results/`.

## Running things

- **Eval data** lives in `~/sheets-evals`, never in the session
  scratchpad: macOS empties `/private/tmp` after three days.
  - **SpreadsheetBench Verified:** `~/sheets-evals/data/ssb/spreadsheetbench_verified_400`
    (15 MB, CC BY-SA 4.0, SHA-256
    `10ef893dd29cb13ab97143ea787e68cdc9574a13873ab9a54e50b31dc03fc949`,
    from https://huggingface.co/datasets/KAKA22/SpreadsheetBench).
  - **Run outputs:** under `~/sheets-evals/runs/<date>-<name>/`.
- **The engine corpus test:**

      SHEETS_CALC_CORPUS=~/sheets-evals/data/ssb/spreadsheetbench_verified_400/spreadsheet go test ./internal/xlsx -run TestEngineCorpus -v -timeout 1h

  `SHEETS_CALC_CORPUS_FN=MATCH` lists one function;
  `SHEETS_CALC_CORPUS_SHOW=n` sets the number of examples.
- **The engine probe:**

      SHEETS_PROBE_FILE=book.xlsx SHEETS_PROBE_CELL=I3 SHEETS_PROBE='A1|SUM(B:B)' go test ./internal/xlsx -run TestEngineProbe -v

- **Evals:**
  - The `claude` CLI is logged in.
  - The skill arm needs `EVAL_BASELINE_SKILL` set to the xlsx skill's
    `SKILL.md`, found under `~/Library/Application
    Support/Claude/local-agent-mode-sessions/skills-plugin/`. That skill
    is proprietary: use it locally, never commit it.
  - The MCP arm uses `uvx excel-mcp-server==2.2.2`.
  - LibreOffice, for grading, is `/opt/homebrew/bin/soffice`.
