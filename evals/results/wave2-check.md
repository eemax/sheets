# Does check beat a hint? SpreadsheetBench Verified, wave 2 (WP-02)

## Pre-registration (written and committed 2026-11-18, before any run was graded)

**Question.** An agent that writes a workbook with Python (openpyxl) leaves
formulas with no stored results and newer functions without their
`_xlfn.` prefix. Does running `sheets check --fix` on what it wrote do
better than telling the agent about the prefix in one line of its prompt?

**Tasks.** The same 90 SpreadsheetBench Verified tasks as 2026-10-25: the
union of `ssb.py --sample 40` and `--sample 80`, seed 2026 (54
cell-level, 36 sheet-level; the 23 UNGRADABLE tasks left out).

**Arms.**

- **python:** openpyxl and pandas (`evals/arms/python.md`), run.
- **python-hint:** the same, plus one line on writing newer functions with
  their `_xlfn.` prefix (`evals/arms/python-hint.md`), run.
- **python+check:** derived, not run. Each python output is copied and
  `sheets check --fix` (this worktree's sheets) is applied before grading,
  so it pairs with python by task, with python's cost and transcript.

Runs: `claude -p` on claude-sonnet-5-5, one per task and run arm (180),
$3 budget each, 20-minute timeout, 8 at a time. Runs that fail to start
are rerun once.

**Gradings.** Each arm is graded both ways:

- **SpreadsheetBench (lo):** as the benchmark does; LibreOffice
  recalculates the result, then openpyxl compares the answer range with
  the golden file.
- **Stored values (stored):** no recalculation; the result's cached values
  are read as openpyxl (`data_only`) and pandas read them, and a formula in
  the answer range with no stored result fails.

**The test.** Under each grading: an exact one-sided McNemar test of
python+check against python-hint, on per-task pass, at p < 0.05. Tasks
whose golden file fails a grading are UNGRADABLE under it for every arm
alike.

**Decision rule (Max, at the plan's gate).** check passes if it beats the
hint under stored values (p < 0.05) **and** is not worse under the
SpreadsheetBench grading: python-hint does not pass more tasks than
python+check there. Both gradings are reported in full. If check does
not pass, check stays a convenience and stops being the pitch.

**Also reported.** Pass rates under both gradings; the paired tables for
python+check vs python-hint, python+check vs python and python-hint vs
python; silent defects (a failed run whose final answer claims success,
by `run.claim`), with the rule spot-checked by hand on 10 transcripts;
failure categories; cost; what drove any difference (stored results,
prefixes or other).

**Expectation, stated in advance.** Stored values will very likely favour
check, because python-hint's openpyxl files store no results. The
SpreadsheetBench grading will likely tie, as the hint fixes the prefixes
too.

---

## Results (graded 2026-11-18, after the pre-registration above)

**Verdict:** check passes Max's rule. Under stored values, python+check
beat python-hint 75 to 15 (60 tasks only it passed, none only the hint
passed; one-sided exact McNemar p ≈ 9 × 10⁻¹⁹). Under the
SpreadsheetBench grading it was not worse, and in fact better: 78 to 70
(11 against 3, p = 0.029). Both gradings point the same way. Most of
the stored-values win is stored results, which the hint cannot give. The
SpreadsheetBench win is prefixes the hint did not cover (`AGGREGATE`).
check stays the pitch.

### Runs

- All 180 runs (90 tasks × python, python-hint) finished. **None failed
  to start**: every run has turns and no error, and no `meta.prev.json`
  exists, so nothing was rerun. python+check was derived from every
  python output (`sheets check --fix` at dd94be1).
- Raw outputs, `results.json`, `report.md` and `run.log` are in
  `~/sheets-evals/wave2/c/`.
- Agents ran on claude-sonnet-5-5 through `claude -p` on Max's plan, at
  most $3 each and 20 minutes. LibreOffice was used only for grading.
- **No task was UNGRADABLE.** LibreOffice converted every file. The
  23 tasks already UNGRADABLE are outside the 90.
- **Validating the stored-values grading** (`ssb.py validate`, log in
  `~/sheets-evals/wave2/c-validate/validate.log`): the 90 golden files
  pass both gradings, 90/90. Each golden file was also loaded and saved
  with openpyxl, which drops the stored results. On every task whose
  answer range holds formulas, those copies fail the stored grading: 30/90
  pass, all of them tasks whose answers are values. The log reports
  "90 tasks, 0 not as expected". So the grading agreed with the golden
  files before any arm was graded. Rerun at the final commit
  (`~/sheets-evals/wave2/c-validate2/`): the same 90 golden files, 0 not
  as expected, and its report.md shows the tables with no PARTIAL line,
  agreeing with validate.txt (golden-resaved: 89/90 lo, 30/90 stored).

### Pass rates

| arm | SpreadsheetBench (lo) | stored values | cell-level lo / stored | sheet-level lo / stored |
|---|---|---|---|---|
| python | 68/90 | 17/90 | 37 / 0 of 54 | 31 / 17 of 36 |
| python-hint | 70/90 | 15/90 | 39 / 0 of 54 | 31 / 15 of 36 |
| **python+check** | **78/90** | **75/90** | 45 / 42 of 54 | 33 / 33 of 36 |

Reference (2026-10-25, the same 90 tasks, lo grading, not rerun): sheets
81, mcp 80, skill 72, python 71. Today's python (68) is within run noise
of 71.

### Paired tables (exact one-sided McNemar, the first arm better)

SpreadsheetBench (LibreOffice recalculates):

| pair | both | first only | second only | neither | p |
|---|---|---|---|---|---|
| **python+check vs python-hint** | 67 | 11 | 3 | 9 | **0.0287** |
| python+check vs python | 68 | 10 | 0 | 12 | 0.0010 |
| python-hint vs python | 65 | 5 | 3 | 17 | 0.363 |

Stored values (no recalculation):

| pair | both | first only | second only | neither | p |
|---|---|---|---|---|---|
| **python+check vs python-hint** | 15 | 60 | 0 | 15 | **8.7 × 10⁻¹⁹** |
| python+check vs python | 17 | 58 | 0 | 15 | 3.5 × 10⁻¹⁸ |
| python-hint vs python | 15 | 0 | 2 | 73 | 1.0 |

check never lost a task that python passed, under either grading. The
per-task table is in `~/sheets-evals/wave2/c/report.md`.

### The two gradings, side by side

They agree on direction and on the decision. For python+check they also
agree task by task, with three exceptions: 32612, 55977 and 56427 pass
under lo but fail under stored. In each of those, sheets left some
formulas unknown rather than guess (`TEXT(A2,"ddd")`, and an
`INDEX`/`VALUE` chain), so no result was stored. That is "unknown over
wrong", and it costs check 3 tasks on stored values. In no task did
sheets store a value that LibreOffice then computed differently in the
answer range. For python and python-hint the gradings disagree
massively, 68 against 17 and 70 against 15. openpyxl stores no results,
so every formula answer fails when read as pandas or a preview reads it.

### What drove the difference

- **Stored results (stored grading).** Of python+check's 60 wins over
  the hint, 59 are tasks where the hint's answer range holds formulas
  with no stored result. The other one is 269-44, where the hint's run
  was simply wrong. 71 of python-hint's 75 stored failures
  are "formula with no stored result". A prompt line cannot fix this,
  because openpyxl cannot compute. `check --fix` stored computed results
  in every one of those files except the three unknowns above.
- **Prefixes (SpreadsheetBench grading).** 8 of check's 11 lo wins over
  the hint are `AGGREGATE` written without `_xlfn.` (10452, 17-35, 1818,
  41410, 50683, 50971, 56378, 58499). LibreOffice gives `#VALUE!`, as
  Excel would give `#NAME?`. The hint lists XLOOKUP, IFS, MAXIFS,
  TEXTJOIN, CONCAT and UNIQUE as examples. The agent applied it to the
  functions it named but not to `AGGREGATE` (Excel 2010), the most
  common newer function in this sample: 8 files carried it unprefixed in
  both run arms, and none after check. Where the hint names the function,
  it works: python lost 516-46 (`MAXIFS`) and 50768 (`TEXTJOIN`) to
  `#NAME?`, and python-hint got both right. A one-line hint only covers
  the long tail if it lists every function, and check knows them all.
- **Other (run variance).** The remaining 3 check-only lo wins (269-44,
  50051, 52532) are tasks python solved and the hint arm solved
  differently and wrongly. The 3 hint-only wins (45896, 59595, 6239) are
  tasks python got wrong and the hint arm's run got right. Neither
  involves check, which only ever turned python failures into passes.
  Remove these 6 and the lo gap is 8 to 0, all prefixes.

### Failure categories

| arm, grading | no stored result | Excel error from a missing prefix | wrong value (the agent's logic) |
|---|---|---|---|
| python, lo | n/a | 10 | 12 |
| python, stored | 69 | (in "no stored result") | 4 |
| python-hint, lo | n/a | 8 | 12 |
| python-hint, stored | 71 | (in "no stored result") | 4 |
| python+check, lo | n/a | 0 | 12 |
| python+check, stored | 3 (sheets: unknown) | 0 | 12 |

check fixes no logic. The 12 wrong-value failures are the same as
python's. They include the tasks that fail in every arm (13284, 146-49,
534-26, 80-42, 54667, 42216, 45300, 45944, 51090), some of them flawed
tasks as on 2026-10-25.

### Silent defects

A silent defect is a failed run whose final answer claims success
(`run.claim`, the rule in `evals/README.md`):

| arm | silent, lo | silent, stored |
|---|---|---|
| python | 7 | 19 |
| python-hint | 6 | 16 |
| python+check | 5 | 6 |

python+check uses python's transcript, so its silent runs are python's
silent runs that check did not rescue.

**Spot check of the rule on 10 transcripts**, read by hand against the
grade. One standard for the hand label: an answer that says the file was
not opened, recalculated or calculated, or that its results are
untested or unchecked, **about the answer itself** is doubt, so it is not
a claim of success (hand label: failure). A remark about something else
(the script's count, assumptions about the task) is not doubt. Agreement
is 4 of 10:

| run | rule | by hand | note |
|---|---|---|---|
| 10452 / python | success | **failure** | "haven't opened the file in Excel or recalculated it, so the results are untested" |
| 59595 / python | success | **failure** | "haven't opened the file in Excel, so the formulas haven't been calculated"; leaves C4:C6 "n/a" |
| 6239 / python | success | **failure** | "haven't opened the workbook to check the results" |
| 42216 / python-hint | success | **failure** | "none of them have been calculated or checked" |
| 146-49 / python-hint | success | success | "didn't reopen the saved file" is about the script's count, not a doubt about the answer |
| 50683 / python | failure | failure | "couldn't recalculate it here … haven't seen it compute" |
| 534-26 / python | failure | failure | "haven't opened the file in Excel to check it" |
| 269-44 / python-hint | failure | failure | "your requirement … isn't met" |
| 1818 / python-hint | success | **failure** | "I haven't opened the file in Excel, so the formulas haven't been calculated" |
| 45300 / python | success | **failure** | "I haven't opened the file in Excel, so the results are unchecked" |

The six disagreements are one miss, all in the same direction: the rule
reads *haven't opened the file in Excel* followed by *haven't been
calculated*, *untested* or *unchecked* as no doubt (tested alone, each
sentence gives `run.claim` → `none`), so the answer's later success
words win. By the other standard (such caveats are not doubt) the labels
of those six would be success and agreement would be 10 of 10; the
agreement figure depends on that choice, and the table uses the
stricter one. Almost every openpyxl answer carries a line like this, so
the silent counts above are upper bounds, not measurements. The fix
belongs to `evals/run.py` (eval-arms' file, outside this piece). It is
named here as a follow-up for integration. It does not touch the verdict, which rests on pass
rates alone.

### Cost

| arm | turns | cost $ (API-equivalent, on Max's plan) | agent minutes |
|---|---|---|---|
| python | 297 | 4.75 | 25.4 |
| python-hint | 294 | 4.75 | 24.7 |
| python+check | 297 (python's) | 4.75 (python's; check adds none) | 25.4 + seconds of `check --fix` |

The two run arms cost $9.51 in total. The hint cost nothing extra and
bought 2 tasks under lo (p = 0.36) and none under stored values.

### What this means

- **The pitch holds.** On what pandas, previews and people read, an
  openpyxl workflow passes 15 to 17 of 90. With one `sheets check --fix`
  it passes 75. A prompt hint cannot close that gap.
- **Even graded as SpreadsheetBench grades**, check beats the hint
  (78 against 70). The reason is that check knows every function needing
  a prefix, and a hint knows only the ones it names.
- **Limits.**
  - This is one run per task and arm.
  - The 6 variance tasks show the noise: a few tasks either way.
  - check's three stored losses are formulas sheets cannot yet compute
    (`TEXT` with a day-name format among them). These are engine
    coverage, not grading.
