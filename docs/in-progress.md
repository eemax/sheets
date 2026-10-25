# Work in progress (2026-10-25)

The plan agreed after the usefulness assessment
(evals/results/2026-10-24-assessment.md, committed as 316df0b):

1. Trust fixes.
2. Close the gaps the assessment found.
3. Harder evals.
4. Scripting with the Office Scripts API.

Prebuilt binaries are deferred (decided 2026-10-24).

Also agreed:

- **Downloads:** SpreadsheetBench Verified may be downloaded, and LibreOffice may be installed (it now is, `/opt/homebrew/bin/soffice`).
- **Scripting:** JavaScript inside sheets, using goja with esbuild to strip TypeScript.

Everything below is **uncommitted**; `go vet ./...` and `go test ./...` pass.

## Done

### 1. A formula engine of our own: `internal/engine`

It replaces excelize's calculator, which disagreed with Excel on 7.2% of real formulas.

- **What it is:** a parser, typed values (number, text, bool, error, empty, reference, array), Excel's operators and coercions, and about 250 functions. It follows Excel's array rules:
  - A range in a one-value slot gives the cell in the formula's row or column.
  - Array formulas (`grid.Sheet.Arrays`, new) evaluate as arrays.
  - `SUMPRODUCT`, `INDEX`'s array, `LOOKUP`, `AGGREGATE` and similar force array evaluation.
  - Array constants lift everywhere.
- **Unknown results:** whatever it cannot compute (functions it lacks, links to other workbooks, `OFFSET` with several offsets) is marked unknown, and anything reading it is unknown too. `IFERROR` never hides it.
- **Language rule:** text is read as a number or date only where every Excel language reads it alike: digits, percentages, exponents, ISO dates and times. Decimal or thousands marks, currency signs, and dates written with slashes or month names make a result unknown, so it is not stored. In the corpus this costs 38 of 68,281 formulas.
- **Against Excel:** over 800 Excel-saved workbooks from SpreadsheetBench, it agrees with Excel's stored results on 95.3% of formulas, differs on 0.35%, and 1.5% are unknown (894 of those are links to other workbooks). What remains:
  - 186 cells: `TEXT` month and day names, which Excel writes in its own language.
  - 1,898 cells: trailing spaces lost in one benchmark file (an artifact of the file, not the engine).
  - A handful of oddities: a `SUMIFS` whose criteria range misses the formula's row, a three-column whole-column array in a one-cell array formula, and a circular `MAX`.
- **Wired in:**
  - The judge asks the engine. Results compare typed.
  - Frozen and native-aggregate formulas come in through `engine.Override`.
  - The read path (`calc.go`) and the full editor use `gridBook` (`calcbook.go`). The full editor reads its current state through `stateNow()`.
  - The streaming editor uses `cellsBook`, over its cell map.
- **`Editor.Evaluate`** returns what saving would store, cached per edit count (`e.edits`). `Editor.EvalAt` computes any formula over the current state.
- **Tests:**
  - `internal/engine/engine_test.go` holds Excel's answers, including every excelize bug found.
  - The corpus test is `internal/xlsx/corpus_test.go`.
  - The engine probe is `internal/xlsx/probe_test.go`.
  - Older tests that assumed excelize's limits now expect Excel's answers.

### 2. Table syntax in formulas (`internal/cli/tablerefs.go`)

- **Translation:** `set` and `write` turn `Sales[Revenue]` into `Sales!$B$2:$B$61` when Sales is a sheet's table rather than an Excel table.
- **Refusal:** a formula naming a table that doesn't exist is refused.
- **Growth on append:** `write --append` grows every range that covers a table column whole (`formula.GrowRanges`; the streaming `x` op; `Editor.ExtendTable`), as an Excel table would. Ranges over part of a column stay.
- **Test:** `cmd/sheets/testdata/tablerefs.txtar`.

### 3. CSV

- **Sniffing:** a delimiter missing from the header counts half. This fixes the corruption seen when a column of decimal-comma numbers was added.
- **Row width:** rows copied as they were are padded when the file gains columns and trimmed when it loses them; ragged files are untouched.
- **In-place changes:** `set` on a CSV computes formulas from the values as they were and stores values. `set D2:D41 '=ROUND(D{r}*1.05,2)'` changes a column in place and keeps its style (`1.311,45`).
- **Batches:** `copy --values` now works within a batch.
- **Tests:** `csvtransform.txtar` and a sniff test.

## Next, in order

### A. Finish trust and gaps: done (2026-10-25)

- `copy --values` keeps a destination's number format; the CSV writer
  reuses old text only when it shows the value as the cell now does.
- Appending fills formula columns (Excel tables' calculated columns, and
  columns filled down in 3 of the last 5 rows); running ranges
  (`$B$2:B9`) no longer grow; a header row maps to the table appended to.
- `analyze`: section headings and blocks continuing a table (reports);
  Excel tables are blocks of their own.
- `lint`: missing-formula, double-count, label-mismatch, unknown-table,
  implicit-intersection, number-as-text outside tables; no false
  positives from heading rows or indented labels. Checked on the corpus.
- `query` reads a sectioned report as one table with `_section`; `info`
  lists sections.
- `read` trims float noise (`ShowNumber`); formulas needing array
  evaluation are stored as dynamic arrays (`engine.ArrayUse`).
- Notes for links to other workbooks no longer claim Excel computes them
  on open. Docs, skill and README updated.

### B. `sheets script`: done (2026-10-25)

- `internal/cli/script.go` (command, session, TypeScript through esbuild,
  goja runtime, console, timeout) and `scriptapi.go` (the ExcelScript
  objects). The script runs inside a batch: writes go to an editor; a read
  after writes saves them into the batch's memory and re-reads, so reads
  see computed results; one save and one undo step at the end.
- Range operations (clear, copy, move, sort) save pending writes first:
  they need a fresh editor. Tables: Excel tables, and sheet tables named
  as `query` names them. Unsupported parts (addTable, charts, conditional
  formats) fail with a pointer to the command that does it.
- Pinned goja to a 2026-09-30 commit (the latest was a day old);
  esbuild v0.28.2. The binary grew from 28.8 MB to 42.8 MB (30.6 MB
  stripped).
- Test: `cmd/sheets/testdata/script.txtar`; 200k rows read in 80 ms.

### C. Harder evals: done (2026-10-25)

- **SpreadsheetBench Verified: done.** 90 tasks × 4 arms on Sonnet
  (`evals/results/2026-10-25-spreadsheetbench.md`): sheets 81, mcp 80,
  skill 72, python 71. Five tasks are flawed; on the sound 85, sheets
  failed four.
  - **Bugs found and fixed:**
    - Unsharing filled-down formulas (clear, copy, move, sort, script)
      stored their results as text; now they keep their types
      (`TestUnshareKeepsResults`).
    - `INDEX(…,,col)` comparisons are now stored as arrays.
  - **New check:** openpyxl's unprefixed functions (`AGGREGATE`, Excel
    shows `#NAME?`) read as unknown, and lint reports them
    (`unprefixed-function`, `unprefixed.txtar`).
  - **Grader:** pooled LibreOffice profiles fix a race; agents are
    stopped after 20 minutes.
- **Our own tasks: done.** The 18 on Sonnet, Haiku and Opus
  (`evals/results/2026-10-25-own-tasks.md`). sheets 53/54, python 54,
  skill 54, mcp 52; stored results 45/45 against 0–4.
  - **Bug fixed:** the one sheets failure was bug 5 again, which the
    earlier fix had only partly fixed. When sheets edits a CSV, a plain
    number now keeps its decimals (`84,50`) and gets no thousands
    separators it did not have (`delim.Editing`, `csvtransform.txtar`).

### D. Finish

- **Checks: done (2026-10-25).** All of these pass:
  - `go vet` and `go test ./...`
  - `SHEETS_STREAM=always`
  - the race detector
  - `STREAM_ROUNDS=300`
  - the engine corpus: 95.30% agree, 0.19% differ (128), unchanged
- **Left:** commit and push when asked.

## How to resume

- **SpreadsheetBench data** lives in the scratchpad and may be gone after a reboot. To re-fetch it (15 MB, CC BY-SA 4.0, SHA-256 `10ef893dd29cb13ab97143ea787e68cdc9574a13873ab9a54e50b31dc03fc949`):

      curl -L -o spreadsheetbench_verified_400.tar.gz https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/main/spreadsheetbench_verified_400.tar.gz

  Unpack it into its own folder.
- **Corpus test:**

      SHEETS_CALC_CORPUS=<dir>/spreadsheetbench_verified_400/spreadsheet go test ./internal/xlsx -run TestEngineCorpus -v -timeout 1h

  `SHEETS_CALC_CORPUS_FN=MATCH` lists one function; `SHEETS_CALC_CORPUS_SHOW=n` sets the number of examples.
- **Engine probe:**

      SHEETS_PROBE_FILE=book.xlsx SHEETS_PROBE_CELL=I3 SHEETS_PROBE='A1|SUM(B:B)' go test ./internal/xlsx -run TestEngineProbe -v

- **Evals:**
  - The `claude` CLI is logged in.
  - The skill arm needs `EVAL_BASELINE_SKILL=<path to the xlsx skill's SKILL.md>`. The local copy is under `~/Library/Application Support/Claude/local-agent-mode-sessions/skills-plugin/.../skills/xlsx/`.
  - The MCP arm uses `uvx excel-mcp-server==2.2.2`.
