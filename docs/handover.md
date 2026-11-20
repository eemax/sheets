# Handover (2026-11-18)

For the next session. Read [AGENTS.md](../AGENTS.md),
[vision.md](vision.md) and [roadmap.md](roadmap.md) first. This note
covers only where things stand, where the data lives, and what wave 1
taught.

## Where things stand

- **Wave 2 is under way,** revised on 2026-11-18 after Fable 5.2's review
  of sheets (`~/sheets-evals/fable-review/review.md`). Max accepted it
  whole. The core is changing a workbook people already use, with `check`
  as the on-ramp. The roadmap board lists pieces A to E.
- **On `main`:** wave 1, WP-16 and the portal's checks
  (`.portal-checks.yaml`). WP-13 is still open and not in wave 2.
- **Order:** A (WP-14, then WP-18 and WP-15) and E (the fuzzer) first;
  then B (evals that can fail) and C (`check` against a prefix hint); D
  (the change report) builds meanwhile. The gate follows B and C. Their
  verdicts are reported to Max, not used to cut pieces.
- **Waiting on Max:** WP-11, the Excel oracle (a Business Basic tenant).
  WP-07b stays parked.

## Data

- **Eval data** is in `~/sheets-evals`. Never use the session scratchpad;
  macOS empties `/private/tmp`.
  - SpreadsheetBench Verified: `~/sheets-evals/data/ssb/`.
  - Wave 2's raw run outputs: `~/sheets-evals/wave2/`; results files go
    in `evals/results/`.
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
  The fuzzer (piece E) is meant to replace much of that by hand.
