# Handover (2026-11-03)

For the next orchestrator session. Read [AGENTS.md](../AGENTS.md),
[vision.md](vision.md), [roadmap.md](roadmap.md) and
[orchestration.md](orchestration.md) first. This note covers only what
those don't: where things stand, and what was learned running wave 1.

## Where things stand

- **On `main`, pushed (ccc788d):** wave 1 and WP-16.
  - WP-01 `sheets check`;
  - WP-07a, round 2 of the Excel check pack;
  - WP-12, the corrupt new sheet;
  - WP-16, edits inside legacy arrays.
- **Everything is clean:** no worktrees, branches or uncommitted work are
  left, and every package was audited before its push.
- **The roadmap board is current.** The open fixes are WP-13, WP-14,
  WP-15 and WP-18.
- **Usage:** the weekly limit stood at about 61% before WP-16 and resets
  on 2026-11-04 at 22:00 UTC. The budget rule is in orchestration.md.

## Proposed next, waiting on Max's go

- **Wave 2:**
  - **The gate: WP-02 with WP-08's silent-defect measure.** Rerun Python
    and Python + skill on the same 90 SpreadsheetBench tasks (the union
    of `ssb.py --sample 40` and `--sample 80`, seed 2026). Score each as
    is, after `check --fix`, and with the agent told to finish with
    `check`. Then answer the kill criterion.
  - **WP-14:** the engine gives unknown where it computes unlike Excel.
    The list is in the roadmap; the data is in
    `~/sheets-evals/wp01/corpus-op-out/unlike-excel.txt`. Do this before
    any doc tells agents to always finish with `check`.
  - **WP-17,** proposed but not yet agreed by Max: random workbooks from
    openpyxl and XlsxWriter, random sheets edits, checked against
    LibreOffice and fixed rules (rows in order, formula kinds kept,
    nothing outside the edit changed). It turns what the auditor found
    by hand into a machine.
- **Small fixes that fit between waves:** WP-18, WP-13, WP-15.
- **Waiting on Max:** WP-11, the Excel oracle. The route is decided (a
  Business Basic tenant of his own), but he hasn't set it up. Anything
  needing Excel installed is parked.

## Lessons from running wave 1

- **Builders compare old and new on real files when they touch shared
  code.** Run the last pushed binary and the new one over corpus files,
  with the edits the change could affect, and compare formulas, their
  kinds and values. Include the files the change is about: for arrays,
  all 151 corpus files that hold one, not a sample.
  - Two regressions (empty-text results, and formulas becoming arrays on
    a row insert) were caught only by the auditor.
  - Both would have shown up in such a comparison.
- **`check --fix` raises the stakes of every engine gap,** because it
  vouches for what it stores. So trust comes first: WP-14 before WP-10.
- **The auditor's tools are worth keeping.** Its repro files and its
  formula-kind checker are in `~/sheets-evals/wave1-audit/` (`bin/` holds
  the checker). Give a new auditor that path.
- **Briefs should name the corpus checks** (`~/sheets-evals/wp01/*.py`,
  `~/sheets-evals/wp16/compare3.py`) so builders reuse them rather than
  writing new ones.
- **Sub-sessions:** the old session couldn't call `start_session`, so
  builders ran as background subagents in worktrees. A fresh Code
  session with `/orchestrate` can start real sessions that Max can watch
  and steer.

## Data

- **Eval data:** in `~/sheets-evals`. Never use the session scratchpad;
  macOS empties `/private/tmp` after three days.
  - SpreadsheetBench Verified: `~/sheets-evals/data/ssb/`.
  - WP-01's corpus scripts and logs: `~/sheets-evals/wp01/`.
  - WP-16's comparisons: `~/sheets-evals/wp16/`.
  - The auditor's files: `~/sheets-evals/wave1-audit/`.
