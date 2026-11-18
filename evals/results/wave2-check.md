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
