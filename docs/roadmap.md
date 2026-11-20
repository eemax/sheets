# Roadmap

Where sheets is going, and why, is in [vision.md](vision.md). How the work
is run (workers, audits, merging) is in [orchestration.md](orchestration.md).
This file is the plan: a status board, then Horizon 1 as work packages.
Where the last session left off is in [handover.md](handover.md).

## Board

Wave 2 is revised (2026-11-18, from the review of sheets in
`~/sheets-evals/fable-review/review.md`). Its pieces are A to E.

| Package | What | Wave | Status |
|---|---|---|---|
| WP-01 | `sheets check` | 1 | done |
| WP-07a | Excel check, round 2: what sheets writes today | 1 | done |
| WP-12 | Fix: a dynamic array set on a sheet the same command adds is saved corrupt | 1 | done |
| WP-16 | Fix: every edit Excel refuses inside an array is refused; hand over to the full editor only when an edit reaches an array | 1 | done |
| WP-14 | A: engine: what it computes unlike Excel becomes unknown, so check --fix never stores it | 2 (A) | done |
| WP-18 | A: fix: moving or copying a plain range formula makes it spill; sort refuses single-cell dynamic arrays | 2 (A) | done |
| WP-15 | A: fix: spills lost by a script's sheet copy, and after deleting a row inside a spill | 2 (A) | done |
| WP-08 | B: evals that can fail (narrowed): a 20-edit session and a feature-loss task | 2 (B) | done |
| WP-02 | C: does `check --fix` beat a one-line prefix hint? (cheap) | 2 (C) | done |
| WP-03 + WP-09 | D: the change report (merged): `check --against`, the short form after each edit, a Markdown receipt | 2 (D) | done |
| WP-17 | E: the fuzzer | 2 (E) | done |
| Gate | After B and C: Max reviews their verdicts (below) | 2 | waiting on Max |
| WP-05 | `explain` and `trace` | after the gate | planned |
| WP-04 | Expectations the workbook carries | after the gate | planned |
| WP-06 | `what-if`: folded into `set --dry-run --show`, later | later | folded |
| WP-10 | Skill and docs for the new position | beyond this run | planned |
| WP-13 | Fix: render computes what read calls unknown; dates written into narrow columns show ######## | — | planned |
| WP-07b | Excel check: the new features; then Max opens the pack in Excel | — | parked |
| WP-11 | An Excel oracle: real Excel computing, through Microsoft Graph | — | waiting on Max (Business Basic tenant) |

**Order of wave 2:** A and E first, in parallel. Then B and C. D builds
meanwhile. The gate comes after B and C.

**Wave 2 landed (2026-11-20).** A, B, C, D and E are done, on `main`.

**The verdicts at the gate**, as their results files state them:

- **B** ([wave2-claims.md](../evals/results/wave2-claims.md)): "**Verdict:**
  fail. sheets' defect rate is 5/201 (2.5%); the best Python arm is
  `skill-lo` at 7/201 (3.5%), and half of that is 1.74%, so the first
  condition fails. The second holds: sheets cost $1.27 against
  `skill-lo`'s $1.85 (0.69 times; the bar is 1.3)." Measured on sheets
  after piece A, without D's change report.
- **C** ([wave2-check.md](../evals/results/wave2-check.md)): "**Verdict:**
  check passes Max's rule. Under stored values, python+check beat
  python-hint 75 to 15 (60 tasks only it passed, none only the hint
  passed; one-sided exact McNemar p ≈ 9 × 10⁻¹⁹). Under the
  SpreadsheetBench grading it was not worse, and in fact better: 78 to 70
  (11 against 3, p = 0.029)."

**Follow-ups for the next wave** (filed as issues in the portal repo,
labelled `repo:sheets`):

- **B again, with D's per-edit report** in the sheets arm.
- **The fuzzer's known bugs.** Its last long run (seed 2721491, 20
  minutes, [internal/fuzz/README.md](../internal/fuzz/README.md)) made
  874 findings in 55,712 edits, **156.9 per 10,000, 0 new**, all from
  bugs `internal/fuzz/known.txt` names. Its 23 bug lines name 14 bugs,
  each a fix to make: `escape`, `sort-editors`, `map-order`,
  `script-legacy-dynamic`, `table-header-control`, `spill-overwrites`,
  `insert-turns-dynamic`, `table-header-renamed`, `replace-loses-kind`,
  `full-keeps-array-metadata`, `intersection-kept`, `stream-note-prefix`,
  `sheet-copy-note` and `spill-into-merge`.
- **`check --against`, report-regress f3.1 to f3.5:** the cases where
  pairing two versions by their contents alone guesses wrong (rows
  typed over, cleared or removed at a table's end). Next step: when
  sheets made the edits, use the session's recorded ops instead of
  inferring them.
- **Silent-defect counts:** `evals/run.py`'s rule now reads "haven't
  opened the file in Excel" as doubt (d4710c1), and both results files
  are regraded with it. What is open is the 10-transcript spot-check
  under the landed rule, which was not repeated.
- **The git diff driver** is built (`sheets diff` as git's diff driver,
  change-receipt); what is left of History in Horizon 2 is receipts on
  pull requests and `check` as a CI step.

**Parked (Max, 2026-11-03):** whatever needs Excel installed. That's
WP-07b, and Max opening the Excel check pack.

**Gate after wave 2.** B's and C's verdicts, the two numbered tests in the
vision's kill criterion, are reported to Max. They are reported, not used
to cut pieces. B measures sheets as of piece A, without D's report; a
rerun with D is for the next wave. WP-05 and WP-04 come after the gate.

Not in wave 2: `explain`, `what-if`, expectations, session mode, the skill
rewrite (WP-10), WebAssembly and MCP mode.

Statuses: `planned`, `working`, `auditing`, `fixing`, `done`, `parked`,
`folded`.

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

### WP-02 C: does `check` beat a one-line prefix hint?

- **Goal:** test the second numbered test of the vision's kill criterion,
  cheaply.
- **Sample:** SpreadsheetBench tasks, as the harness draws them, with the
  sample and seed written in the results.
- **Arms:**
  - **python:** plain, as before.
  - **python + a one-line prefix hint:** told to write newer functions
    with their `_xlfn.` prefix.
  - **python, then `check --fix`:** `sheets check --fix` applied to the
    output before grading.
- **Grading, twice, both reported in full:**
  - **SpreadsheetBench's way,** recalculating in LibreOffice;
  - **stored values,** read with no recalculation, as pandas, previews and
    people read them. The hint arm's openpyxl files store no results.
- **Pass bar (Max, 2026-11-18):** `check` passes only if it beats the hint
  arm on stored values on a paired test, and is no worse under the
  SpreadsheetBench grading. The results say whether the win is more than
  stored results. If it fails, `check` stays a convenience and stops being
  the pitch.
- **Acceptance:**
  - **Results:** a document in `evals/results` with paired statistics,
    failure categories and the verdict.
  - **Harness:** the changes, committed.
  - **Run outputs:** kept under `~/sheets-evals/wave2/`.
- **Depends on:** WP-01, WP-14 (piece A) and the eval arms' interface.

### WP-03 with WP-09: the change report (piece D)

- **Goal:** one report on what a change did, whoever made it.
  - **`sheets check book.xlsx --against before.xlsx`:** inputs and formulas
    that changed; computed values that moved, each traced to its cause;
    values that moved with no cause, flagged; features lost (charts,
    validations, conditional formats, shapes, names); unknowns.
  - **After every sheets edit:** the same report, short.
  - **A Markdown receipt** with rendered before and after, for a pull
    request or a message.
- **Why:** an unexplained change is how a broken formula, a lost result or
  a bad round trip shows itself, and it needs no authoring. This catches
  what openpyxl silently drops, which no skill line fixes.
- **Acceptance:** planted cases in tests:
  - **an input change:** every moved value is explained;
  - **one row's formula broken:** flagged;
  - **an openpyxl round trip that lost results:** flagged;
  - **an openpyxl round trip that drops a chart and a validation:**
    flagged;
  - **a sort:** reported as moved rows, not changes.

  Output stays bounded on 150,000 rows.
- **Depends on:** diff's row and column alignment, and the engine's
  dependencies. Part of wave 2; it builds meanwhile B and C run.

### WP-04 Expectations the workbook carries

Comes after the gate.

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

Comes after the gate.

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

### WP-06 `what-if` (folded)

Folded into `set --dry-run --show`, later (the review, 2026-11-18). That
shows values before and after and saves nothing, which is what `what-if`
was for. Not in wave 2.

- **Still to hold:** the file's bytes unchanged, no undo step recorded, and
  results equal to a real edit followed by recomputation. Unknown values
  show as unknown.

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

### WP-08 B: evals that can fail

- **Goal:** test the first numbered test of the vision's kill criterion.
  Two tasks where the claim lives:
  - **A 20-edit session** on a feature-rich workbook (tables, names, a
    chart, conditional formats), graded on the end state: every formula's
    value against an independent computation, and every feature intact.
  - **A feature-loss task:** edit a workbook with a chart, validations,
    conditional formats and a shape.
- **Silent defects:** a failed run whose final answer reports success,
  counted with a written, spot-checked rule.
- **Arms:** sheets; Python; Python with the xlsx skill and LibreOffice
  available. The skill arm is the committed Apache 2.0 skill. Three runs
  each.
- **Pass bar:** sheets has at most half the best Python arm's defect rate,
  at no more than 1.3 times its cost.
- **Acceptance:**
  - **Conventions:** the tasks follow `evals/tasks` (`setup.py`,
    `check.py`, `prompt.md`).
  - **The checks:** validated against one hand-made right solution and one
    hand-made wrong one.
  - **Results:** in `evals/results`; raw outputs under
    `~/sheets-evals/wave2/`.
- **Measured as of piece A,** without D's report.
- **Not in this package:** the model-review task and the two reference
  workflows from the earlier WP-08.

### WP-09 Receipts (merged into WP-03)

Merged with WP-03 as the change report, piece D of wave 2. See WP-03.

### WP-10 Skill and docs for the new position

Beyond this run. Wave 2 changes the skill only as far as a new command
needs it to stay true.

- **Goal:** the skill, `agent-docs` and the README teach:
  - when to use sheets and when Python;
  - to finish with `sheets check`;
  - expectations, `explain` and `what-if`.
- **Acceptance:** the trigger eval (`evals/triggers.py`) still chooses the
  skill when it should.
- **Depends on:** WP-01, WP-04, WP-05, WP-06.

### WP-12 Fix: a corrupt sheet from a dynamic array on a new sheet

- **Found by:** WP-07a's builder.
- **Reproduction:** `printf '%s\n' Name d b a c | sheets write n.xlsx Data`,
  then `sheets set n.xlsx S!A1 x S!B3 '=SEQUENCE(4)' S!B8 u`.
- **What goes wrong:** the new sheet's XML holds a second `B3` cell inside
  row 8, and its rows come out of order. Without B8, row 3 appears twice.
  The same happens through `sheets script` (`addWorksheet`, then
  `setFormula`), where a later value also lands in the wrong cell.
- **Acceptance:**
  - the reproduction gives valid sheet XML: rows in order and unique,
    cells in order and unique, the array's anchor `ref` spanning the
    spill;
  - a test covers the case, in the full and the streaming editor and
    through `script`;
  - `excel-check/` file 20c is then built as a normal entry.

### WP-13 Fix: render and narrow columns

- **Render and unknown values:** `render` draws values that `read` calls
  unknown (an unprefixed function, which Excel shows as `#NAME?`). Render
  should show what Excel would show, or the unknown marker.
- **Dates in default-width columns:** a date that `set` or `script` writes
  into one shows `########`. Writing a date should widen the column as
  Excel does when you type one.
- **Smaller:** a long decimal in General format draws as `########`
  (Excel shows 0.348837), and TRUE and FALSE should be centred.

### WP-14 Engine: unknown where it computes unlike Excel

- **Found by:** WP-01's builder. It round-tripped all 800 corpus files
  through openpyxl, then ran `check --fix`, and compared each stored
  result with what Excel had stored.
- **What it found:** of 65,460 results, 2,218 differ.
  - **About 2,000 differ only in spaces.** Strings in the dataset files
    were trimmed, so this is not the engine.
  - **The rest are wrong values that `--fix` would store:**
    - `TEXT(date,"mmmm")` in English where Excel used its own language;
    - a Swedish `ÅÅÅÅ-MM-DD` format;
    - `SUMPRODUCT` over whole columns;
    - `ERROR.TYPE` inside `IF`;
    - `SUMIFS` with `DATE` criteria.
  - **The full list** is in `~/sheets-evals/wp01/corpus-op-out/unlike-excel.txt`.
- **Also:** in 1904-date workbooks, date results are unknown for now (WP-01).
  Teach the engine the 1904 system instead. A date written as text in a
  criterion (`SUMIF(A1:A2,"<1/2/2020")`) is mishandled in both date
  systems.
- **Goal:** each class either matches Excel or gives unknown, never a
  value Excel would not show. Month and day names, and other formats
  that depend on Excel's language, are unknown (the language rule).
- **Acceptance:**
  - the comparison from `~/sheets-evals/wp01/roundtrip.py` shows no
    difference outside spaces;
  - the engine corpus test does not get worse;
  - tests cover each class.

### WP-15 Fix: two spill bugs

- **Found by:** the wave-1 audit. Both predate it.
- **A script's sheet copy:** in `sheets script`, `ws.copy()` after
  `setFormula("=SEQUENCE(3)")` gives a copy whose array formula has no
  `cm` mark and no value. Excel would show a one-cell legacy array. The
  same copy through `do` (`sheet copy`) is fine.
- **A row deleted inside a spill:** `rows delete Data!3` inside a spill
  D1:D4 leaves the old spilled values as plain values, so D1 becomes
  `#SPILL!`. Excel would spill a, c, d again.
- **Acceptance:** both cases match Excel in both editors; tests; XML order
  checked with `internal/xlsxtest`.

### WP-16 Fix: edits inside legacy arrays

- **Found by:** the wave-1 audit's third round. These are left over, not
  regressions.
- **Writes into part of a legacy array:** `set`, `write`, `replace` and
  script `setValue` refuse them, but `clear`, `copy`, `move` and script
  `Range.clear()` still go through. They leave values Excel would not
  show, and `check` vouches for them. Excel refuses all of these.
- **Cutting an array:** a row or column insert or delete through its
  range leaves the old results behind as plain values. Deleting the
  anchor's row gives no warning. Excel refuses these edits too.
- **Speed:** one small legacy array sends every edit of a large workbook
  to the full editor (4.1 s instead of 1 s on 150,000 rows). Hand it
  over only when an edit reaches an array, or moves rows or columns.
- **Acceptance:**
  - every edit Excel refuses inside an array is refused, with the same
    message;
  - `check` never vouches for values inside an array that sheets did not
    compute;
  - plain edits on large workbooks are back to their earlier speed;
  - tests in both editors.
- **Repro files:** the auditor's, copied to `~/sheets-evals/wave1-audit/`.

### WP-18 Fix: formula kinds in move, copy and sort

- **Found by:** WP-16's last audit. Both predate it.
- **Move and copy:** moving or copying the cell of a plain formula that
  reads a range (`=B1:B2*100`, implicit intersection) turns it into a
  dynamic array that spills. Excel keeps it plain: at F5, F1's formula
  shows `#VALUE!`. The cause is `put`, through `setFormulaAs`, then
  `dynamicFormula`. Repro: `~/sheets-evals/wave1-audit/k1.xlsx`.
- **Sort:** it refuses rows holding a single-cell dynamic array (`cm="1"`,
  a ref of the cell itself), which Excel 365 saves for many ordinary
  formulas that use arrays. Excel sorts them. 110 of the 800 corpus
  workbooks hold such cells.
- **Acceptance:**
  - formulas keep their kind (plain, legacy array, dynamic array) through
    move, copy and sort, as through row and column edits now;
  - the kind checker in `~/sheets-evals/wave1-audit/bin/` (kcheck.py)
    finds nothing;
  - tests in both editors.

### WP-17 The fuzzer (piece E)

- **Goal:** turn what the auditor found by hand into a machine. Wave 1
  needed three audit rounds per package and turned up about 15 old bugs.
- **Shape:** random workbooks from openpyxl and XlsxWriter, random edits,
  and fixed invariants:
  - valid XML;
  - formula kinds kept;
  - nothing outside the edit changed;
  - both editors agree;
  - undo restores the bytes.
- **LibreOffice** only as a second opinion on plain functions, never the
  judge. The Graph oracle (WP-11) replaces it when it exists.
- **Acceptance:**
  - it finds the wave-1 audit's known bugs when they are put back;
  - it runs in `portal check full` and keeps it under 135 s on the
    MacBook (101 s today);
  - findings go to the portal repo as issues labelled `repo:sheets`.

### WP-11 An Excel oracle (waiting on Max)

- **Goal:** test sheets against Excel itself, not LibreOffice. Microsoft
  Graph's workbook API opens a file in Excel for the web, recalculates it
  with Excel's engine, and returns values. Three things would then be
  checked automatically:
  - the engine's results;
  - the files sheets writes (Excel opens them without repairing them);
  - SpreadsheetBench grading, without LibreOffice's gaps.
- **Route (Max, 2026-11-03): a Microsoft 365 Business Basic tenant of
  Max's own.**
  - **Why business:** Graph's workbook API works only on OneDrive for
    Business and SharePoint, so a personal plan won't do.
  - **Not:** his New Wave Group work account, a client's tenant.
  - **Setup is Max's step:** buying the plan and registering the app.
    The token then goes into `sec`.
  - **Until then:** nothing starts.
- **Acceptance:** a script uploads a workbook, recalculates, and compares
  every formula's value with what sheets stores. It runs on the
  excel-check pack and a corpus sample, and reports where they differ.

## Horizon 2, to be cut into packages after the gate

- **Session mode:** deferred (see [deferred.md](deferred.md)).
- **A Python client:** a thin wrapper over the CLI, not a client of a
  session.
- **History:** through git: `sheets diff` as a diff driver for `.xlsx`,
  receipts on pull requests, `check` as a CI step.
- **`what-if` and `explain`:** `set --dry-run --show`, then `explain`.
- **Templates:** fill a workbook template once per record.
- **Change review:** an agent proposes a change, a person sees it
  rendered and accepts it.
- **Later:** an MCP mode of four tools (look, change, check, render).
  WebAssembly is cut.

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
  - The skill arm uses the committed Apache 2.0 skill in
    `evals/arms/xlsx-skill/`, under the eval-arms piece's contract.
  - The MCP arm uses `uvx excel-mcp-server==2.2.2`.
  - LibreOffice, for grading, is `/opt/homebrew/bin/soffice`.
