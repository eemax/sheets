# Handover (2026-11-20)

For the next session. Read [AGENTS.md](../AGENTS.md),
[vision.md](vision.md) and [roadmap.md](roadmap.md) first. This note
covers only where things stand, where the data lives, and what wave 1
taught.

## Where things stand

- **Wave 2 has landed** (2026-11-20), as Fable 5.2's review of sheets
  planned it (`~/sheets-evals/fable-review/review.md`; Max accepted it
  whole on 2026-11-18). On `main`: A (the engine's unknowns, formula
  kinds and spills), D (`check --against`, the short change report after
  each edit, the Markdown receipt, `sheets diff` as git's diff driver),
  E (the fuzzer: `SHEETS_FUZZ=long`, `scripts/fuzz-mutants.sh`), and the
  evals B and C.
- **The gate waits on Max.** The verdicts, from their results files:
  - **B** ([wave2-claims.md](../evals/results/wave2-claims.md)): fail.
    sheets' defect rate is 5/201 (2.5%); the best Python arm, `skill-lo`,
    is 7/201 (3.5%), and half of that is 1.74%. The cost condition holds
    ($1.27 against $1.85, 0.69 times; the bar is 1.3). Measured without
    D's report.
  - **C** ([wave2-check.md](../evals/results/wave2-check.md)): check
    passes Max's rule: 75 to 15 over the hint under stored values
    (p ≈ 9 × 10⁻¹⁹), 78 to 70 under the SpreadsheetBench grading
    (p = 0.029).
- **Open** (the roadmap's follow-ups, filed in the portal repo as
  `repo:sheets` issues): B rerun with D's report; the fuzzer's 14 known
  bugs (`internal/fuzz/known.txt`; last long run 156.9 findings per
  10,000 edits, 0 new); `check --against`'s report-regress f3.1 to f3.5
  (next step: use the session's recorded ops when sheets made the
  edits); C's silent-defect rule in `evals/run.py`. WP-13 is still open.
- **Waiting on Max:** the gate, and WP-11, the Excel oracle (a Business
  Basic tenant). WP-07b stays parked.

## Data

- **Eval data** is in `~/sheets-evals`. Never use the session scratchpad;
  macOS empties `/private/tmp`.
  - SpreadsheetBench Verified: `~/sheets-evals/data/ssb/`.
  - Wave 2's raw run outputs: `~/sheets-evals/wave2/`; results files go
    in `evals/results/`.
  - The fuzzer's runs: `~/sheets-evals/wave2/fuzz/`; the mutant runs:
    `~/sheets-evals/wave2/fuzz-mutants-final/`.
  - The review: `~/sheets-evals/fable-review/`.
  - Wave 1: `~/sheets-evals/wp01/` (corpus scripts), `wp16/` (comparisons)
    and `wave1-audit/` (repro files; `bin/` holds the formula-kind
    checker).
- **The skill arm** is the committed Apache 2.0 xlsx skill in
  `evals/arms/xlsx-skill/`.

## Lessons from wave 1

- **Compare old and new on real files** when touching shared code: the
  last pushed binary against the new one, over the corpus files the change
  could affect. Two regressions were caught only by the auditor.
- **`check --fix` raises the stakes of every engine gap,** because it
  vouches for what it stores. Trust comes first, hence piece A.
- **Reuse the corpus checks** (`~/sheets-evals/wp01/*.py`,
  `wp16/compare3.py`) and the auditor's tools rather than writing new ones.
  The fuzzer (piece E) now does much of that: run its long mode on
  shared code before an audit.
