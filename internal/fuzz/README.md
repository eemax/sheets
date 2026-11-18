# The fuzzer

A machine doing what wave 1's audits did by hand. It runs the sheets CLI in
process (`cli.Main`) over random workbooks from openpyxl and XlsxWriter. It
makes random edits and checks fixed invariants on the saved files after
each one.

## The default mode

`TestFuzz` runs in every `go test ./...`, so `portal check full` runs it
too.

- **Seed and edits:** seed 1, and 2,000 edits.
- **Time cap:** without `-race`, all 2,000 are made well inside the
  90-second cap. That cap only guards a slow machine, and the run says when
  it cut.
- **Under `-race`:** each edit takes about fifteen times as long. The run
  detects `-race` (a build tag) and stops at a 40-second cap once at least
  200 edits are made. Those edits are a prefix of the plain run's, so they
  make the same findings as far as they go.
- **The log** (`go test -v ./internal/fuzz/`) prints the edit count, the
  edits applied, and each finding's signature. With one seed, two runs
  print the same.

A run is a list of sequences, each with its own seed (seed 1 gives
sequences 100000, 100001, …).

- A sequence's seed picks a workbook from the pool and 1 to 12 edits. Each
  edit is made from the state the edits before it left.
- Each edit runs twice, on two copies of the workbook: once with the
  streaming editor (`SHEETS_STREAM=always`) and once with the full one
  (`never`).
- A sequence ends at its first finding.
- **Edit classes:** `set`, `write`, `clear`, `copy`, `move`, `sort`,
  `replace`, `rows-insert`, `rows-delete`, `cols-insert`, `cols-delete`,
  `sheet-rename`, `sheet-copy`, `do` (a batch of 2 to 4 of the others) and
  `script` (Office Scripts range operations: values, formulas, clear,
  `copyFrom`, `moveTo`, insert and delete of rows and columns, sort).
- **Targets:** a cell is often a formula's cell. The formula is picked by
  its group first (plain, plain reading a range, legacy array, dynamic
  array), so the few arrays get as many edits as the many filled-down
  formulas.

## The invariants

They judge files, never a command's output text.

1. **xml:** every part (`.xml`, `.rels`, `.vml`) is well-formed XML. Every
   worksheet lists its rows in order, each once, and the cells of each row
   in order, each once (`internal/xlsxtest`).
2. **kinds:** every formula keeps its kind: plain, legacy array (entered
   with Ctrl+Shift+Enter) or dynamic array. An array's range must start at
   its cell, and a dynamic array's cell metadata must exist. The model is
   below.
3. *reach* is fuzzer-long's.
4. **editors:** the streaming and the full editor give the same sheets,
   cells, formulas, kinds (arrays and spills) and values, pending or not.
   Both must also take the edit, or both refuse it.
5. **undo:** `sheets undo` restores the bytes from before the edit. An edit
   refused leaves the bytes as they were.
   - Even edits check undo on the streaming copy and odd ones on the full
     copy.
   - The other copy's edit runs with the journal off (`SHEETS_UNDO=off`),
     so each journal holds just the change undo takes back.

A command that panics is reported as `panic`, since nothing after it can be
judged.

### The model of kinds

Every formula in the pool, and every one an edit adds, names its kind in
its text: `N("p")`, `N("l")` or `N("d")`. N of a text is 0, so a marker
changes no result. The model of what each edit does to kinds:

- **Adding formulas:** `set`, `write` and a script's `setFormula` add
  formulas of the kind their text names. That is the kind Excel 365 stores
  for the formula typed in: an array result, or a range where one value
  goes, makes a dynamic array (`=B2:B4*2`, `=SUM(IF(B2:B6>3,B2:B6))`).
- **Keeping kinds:** copy pastes a formula as the kind it was, and
  `--values` pastes none. Move, sort, the rows and columns edits, sheet
  rename and copy, and replace all keep kinds wherever the formulas go.
  Replace in formula text only swaps SUM with PRODUCT and MIN with MAX,
  which changes no kind.
- **Removing formulas:** clear removes a cell's formula, and so does an
  edit that overwrites the cell.

So after any edit, a formula's stored kind is the kind its marker names. A
formula that names none, or more than one, is outside the model and not
judged. The marker survives every edit, because no edit rewrites a string
inside a formula. That is how the model follows a formula through sorts
and moves without tracking positions.

## The pool

There are 40 small workbooks, about 330 KB in all, in `testdata/pool/`:
20 written by openpyxl and 20 by XlsxWriter. They come from
`scripts/fuzzgen.py`, each from its own seed. Each holds:

- values of every type;
- a filled-down formula column, and total rows;
- plain formulas reading a range where one value goes
  (`='Data'!C2:C3*10`);
- legacy arrays, over one cell and over several;
- dynamic arrays (XlsxWriter only, since openpyxl cannot write them);
- second sheets reading the first, some with spaces in their names;
- defined names, Excel tables, charts, validations and conditional
  formats.

To regenerate the pool:

    uv run --with openpyxl --with xlsxwriter python scripts/fuzzgen.py internal/fuzz/testdata

The script reads each file back and refuses one whose stored kinds don't
match their markers. XlsxWriter, for one, stores any formula calling
SEQUENCE as a dynamic array. Regenerating changes the edits each seed
makes, so `known.txt` must then be made again.

`testdata/refuse/` holds the must-refuse cases: pool files with one defect
each, named by the invariant that must catch them. They cover a malformed
part, rows out of order, a cell twice, a dynamic array or legacy array
stored plain, and a plain formula stored as a dynamic array. Two
`editors-*--<pool file>` cases differ from their pool file by a value or a
formula. `TestRefuse` checks each is caught by its invariant. `TestUndoRefuses`
does the same for undo: an undo with nothing recorded, and a journal that
restores other bytes. `TestPool` checks every pool file holds every
invariant, and that every formula in it names its kind.

## Findings and known.txt

A finding's signature is exact: the edit class, the invariant, and the
sequence's seed (`move kinds 100037`).

- **Passing:** the default mode passes only when every finding it makes is
  listed in `known.txt`.
- **A new finding** fails the test. It is first shrunk: whole edits are
  dropped, then lines of batches and scripts, each removal kept when the
  last edit still breaks the same invariant. The failure prints the seed,
  the workbook and the shrunk edits, as shell commands on `book.xlsx`.
  Shrinking replays the edits many times, so only the first five new
  findings of a run are shrunk; the rest print as the run made them.
- **To replay one sequence**, printing every edit and its outcome, then
  the shrunk repro:

      SHEETS_FUZZ_REPRO=100037 go test ./internal/fuzz -run 'TestFuzz$' -v

- **A fix removes its lines.** A run that is not capped lists the entries
  it no longer makes.

Findings on the code this was built on (99ba275):

- **31 × WP-18:** moving or copying a plain formula that reads a range where
  one value goes stores it as a dynamic array that spills. This comes
  through `move`, `copy`, and `do` and `script` batches that move or copy.
- **2 × new:** a sort that finds its rows already in order. The full editor
  then saves a formula on that sheet without its result, while the
  streaming editor stores it (`Summary!B2`: 64 against pending).

`SHEETS_FUZZ_SEED` and `SHEETS_FUZZ_EDITS` change the seed and the count.
That is not the default mode, and `known.txt` lists only the default seed's
findings. The long mode is fuzzer-long's.

## Timings

On Max's MacBook (M4), 2026-11-18, `time scripts/check.sh full`. The bar is
135 s.

| | check full | plain fuzz | race fuzz |
|---|---|---|---|
| main (99ba275), without the fuzzer | 76.6 s | — | — |
| with the fuzzer | 94.8 s | 2,000 edits, 22.9 s | 203 edits, 47.5 s |

- **The race count** is from the race step run with `-v`
  (`go test -race -count=1 -v ./...`), under the same load as in check
  full. Alone, `go test -race -v ./internal/fuzz/` makes 259 edits.
- **Plain** `go test -count=1 -v ./internal/fuzz/` makes 2,000 edits
  (1,654 applied) in 322 sequences, about 21 s, with 33 findings, all
  known. Two runs print the same.

- **Plain:** an edit takes about 11 ms, most of it sheets' own file writes.
- **Under `-race`:** an edit takes about 160 ms.
- **In check full**, the plain fuzz is the plain step's longest package.
  The race fuzz runs beside `cmd/sheets`, whose race run takes about a
  minute.
- **If check full ever nears 135 s:** lower the race cap first, never
  below 200 edits. Never weaken the invariants.
