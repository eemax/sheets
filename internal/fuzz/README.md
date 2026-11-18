# The fuzzer

A machine doing what wave 1's audits did by hand. It runs the sheets CLI in
process (`cli.Main`) over random workbooks from openpyxl and XlsxWriter. It
makes random edits and checks fixed invariants on the saved files after
each one.

## The default mode

`TestFuzz` runs in every `go test ./...`, so `portal check full` runs it
too.

- **Seed and edits:** seed 1, and 2,000 edits.
- **No time cap without `-race`:** a plain run always makes its 2,000
  edits, so two runs make the same edits. A slow machine (heavy load,
  throttling) takes longer and never makes fewer; `go test`'s own timeout is
  the backstop.
- **Under `-race`:** each edit takes about fifteen times as long. The run
  detects `-race` (a build tag) and stops at a 40-second cap once at least
  200 edits are made. Those edits are a prefix of the plain run's, so they
  make the same findings as far as they go. Only a race run's edit count
  depends on the machine, and the test fails below 200.
- **The log** (`go test -v ./internal/fuzz/`) prints the edit count, the
  edits applied, each finding's signature, and how many saved files held a
  formula result of each stored form (see the invariants). With one seed,
  two runs print the same.

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
  `sheet-rename`, `sheet-copy`, `sheet-add`, `do` (a batch of 2 to 4 of the
  others) and `script` (Office Scripts range operations: values, formulas,
  clear, `copyFrom`, `moveTo`, insert and delete of rows and columns, sort).
- **Targets:** a cell is often a formula's cell. The formula is picked by
  its group first (plain, plain reading a range, legacy array, dynamic
  array), so the few arrays get as many edits as the many filled-down
  formulas. The source of a copy or a move is, in three of ten, cells some
  formula reads (the whole range it names, or part): moving those cells
  rewrites the formula, which is where a formula can change its kind, and a
  random cell is rarely one. Formulas that read a range where one value goes
  (`=B3:B4*2`) are picked first.
- **New sheets:** `set`, `write`, `copy` and `move` make the sheet they name
  when it is new, and the fuzzer names a new one now and then. A `do` batch is
  made from the workbook as it is before the batch *and from the sheets its
  own lines make*: half of the batches start by making one (`sheet add`,
  `sheet copy`, or a `set` or `copy` to a new name) and go on to edit it in
  the same session. (A sheet made in a session and edited after the workbook
  was written out saved with rows out of order in wave 1.)

## The invariants

They judge files, never a command's output text.

1. **xml:** every part (`.xml`, `.rels`, `.vml`) is well-formed XML, read
   strictly. Go's decoder accepts an attribute twice on one element
   (`<c r="B2" r="B2">`) and a prefix no declaration binds (`<zz:x/>`), and
   Excel calls the file corrupt, so `strictXML` reads raw tokens and checks
   those itself, with tags, one root and no text outside it. (A streaming
   editor that rewrites a cell's attributes with a regular expression, and
   misses one the cell has, writes the first.) Every worksheet lists its
   rows in order, each once, and the cells of each row in order, each once
   (`internal/xlsxtest`).
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

Besides these, **repeat** is the fuzzer's own check: an edit that removes two or
more dynamic arrays (a paste over several) is made again, eight times with
each editor, on the bytes it first found, and a workbook that holds other
cells, formulas or values than the first is a finding. Sheets lays out the
results of dynamic arrays in map order, so such a paste gives a different
workbook from one run to the next (see Findings), and a finding that comes
and goes would make two runs of one seed differ. Byte-level differences are
not compared: the order of the shared strings a spill adds varies and means
nothing. Only these edits are made again: all of them would cost a third
of the run.

**What the run reaches.** The log counts the files the editors saved that
hold a formula result of each stored form: number, text, empty text
(`t="str"` with an empty `<v>`), boolean, error, none (pending), and a
shared formula. The test fails when a run saved no file holding a number,
text, empty text, boolean or shared formula: the pool and the edits must put
the invariants to each. (A run that stored every result as a number would
pass the invariants and prove nothing.)

### The model of kinds

Every formula in the pool, and every one an edit adds, names its kind in
its text as the condition of an IF that is always true: `IF(N("p")=0,B2*C2)`,
with `N("l")` or `N("d")` for the others. N of a text is 0, so the marker
changes no result, and, unlike an addition (`B2*C2+N("p")`, which the
fuzzer's first version used), no result's type: a text, an empty text, a
boolean and an error stay what they are. The model of what each edit does to
kinds:

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

There are 40 small workbooks, about 340 KB in all, in `testdata/pool/`:
20 written by openpyxl and 20 by XlsxWriter. They come from
`scripts/fuzzgen.py`, each from its own seed. Each holds:

- values of every type;
- a filled-down formula column, and total rows. In about two files in three
  the column is rewritten as shared formulas (one group or two: the first
  cell holds the text and the range, the others point to it), as Excel and
  LibreOffice store a formula filled down; neither producer writes them;
- formulas whose results are not numbers: texts, empty texts
  (`IF(B2>1000,1,"")`) and booleans, filled down, on second sheets, and in a
  legacy and a dynamic array;
- plain formulas reading a range where one value goes
  (`='Data'!C2:C3*10`);
- legacy arrays, over one cell and over several;
- dynamic arrays (XlsxWriter only, since openpyxl cannot write them);
- second sheets reading the first, some with spaces in their names;
- defined names, Excel tables, charts, validations and conditional
  formats.

XlsxWriter stores each formula's result as it does for one given a result: a
text as `t="str"`, a boolean as `t="b"`, an empty text as an empty `<v>`
with no type (which reads as no result) in half the files and, in the other
half, as Excel stores it, `t="str"` with an empty `<v>`. openpyxl stores no
result, which sheets reads as pending. So the pool holds stored and pending
results of every type but an error, and the edits add the rest.

To regenerate the pool:

    uv run --with openpyxl --with xlsxwriter python scripts/fuzzgen.py internal/fuzz/testdata

The script reads each file back and refuses one whose stored kinds don't
match their markers. XlsxWriter, for one, stores any formula calling
SEQUENCE as a dynamic array. Regenerating changes the edits each seed
makes, so `known.txt` must then be made again. (The openpyxl files come out
with another `modified` date in `docProps/core.xml` each time; nothing else
differs.)

`testdata/refuse/` holds the must-refuse cases: pool files with one defect
each, named by the invariant that must catch them. They cover a malformed
part, rows out of order, a cell twice, an attribute twice and an
undeclared prefix (what Go's decoder lets through), a dynamic array or
legacy array stored plain, and a plain formula stored as a dynamic array.
Five `editors-*--<pool file>` cases differ from their pool file by a value,
a formula, a stored text, a stored boolean, or an empty-text result that is
lost. `TestRefuse` checks each is caught by its invariant. `TestStrictXML`
holds 22 more cases of well-formed and not, each with expat's verdict. `TestUndoRefuses`
does the same for undo: an undo with nothing recorded, and a journal that
restores other bytes. `TestPool` checks every pool file holds every
invariant, and that every formula in it names its kind; `TestPoolResults`,
that the pool holds a stored result of each form (text, empty text,
boolean, number, none) and shared formulas.

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

Findings on the code this was built on (99ba275), 37 in all:

- **32 × WP-18:** moving or copying a plain formula that reads a range where
  one value goes stores it as a dynamic array that spills. This comes
  through `move`, `copy`, and `do` and `script` batches that move or copy.
- **4 × new:** a sort that finds its rows already in order. The full editor
  then leaves the formulas on that sheet as they were, while the streaming
  editor computes and stores them all: the full editor's are missing
  (pending) or stale (the 0 XlsxWriter stores). A text result is among them
  (`Data!L2`: "LIME" against pending).
- **1 × new, `copy repeat 100139`:** a paste over several dynamic arrays
  gives a different workbook from one run to the next. Copying `Data!P4:R6`
  onto `M2`, over the arrays in `M2:M4`, `N2` and `O2:O4`, leaves some
  pasted formulas out at random (`Data!M3`, or `M3` and `M4`), whichever
  editor makes it: 20 runs of that one command gave three different files
  (one right, two with formulas missing). `spillResults` (`xlsx/spill.go`) lays the arrays
  out in map order. It comes with the new formulas in the pool: before them
  the source cells of such a paste were empty. Run it twice to see:
  `SHEETS_FUZZ_REPRO=100139 go test ./internal/fuzz -run 'TestFuzz$' -v`.

`SHEETS_FUZZ_SEED` and `SHEETS_FUZZ_EDITS` change the seed and the count.
That is not the default mode, and `known.txt` lists only the default seed's
findings. The long mode is fuzzer-long's.

## Timings

On Max's MacBook (M4), 2026-11-18, `scripts/check.sh full`, wall time. The
bar is 135 s. Other sessions were running tests on the Mac too (load
average 2 to 3.5), so the rounds alternate, the branch then main's base
commit (99ba275, without the fuzzer).

| round | main (99ba275), without the fuzzer | with the fuzzer |
|---|---|---|
| 1 | 87 s (980 tests) | 96 s (1,026 tests) |
| 2 | 61 s | 94 s |

(The fuzzer's first version measured 87 to 107 s against 75 to 77 s on
main, with a simpler pool and no repeat check.)

- **Edit counts:** the plain run always makes 2,000 edits (1,648 applied)
  in 322 sequences, with 37 findings, all known, in 24 s alone. Two runs
  print the same: six at once, under load, did.
- **The race count** is from the race step run with `-v`
  (`go test -race -count=1 -v ./...`), under the same load as in check
  full: 204 edits in 40 s. Alone, `go test -race -v ./internal/fuzz/` makes
  251. It never makes fewer than 200: past the cap, the run goes on until
  200 are made.
- **What the fuzzer adds** is its plain run, about 24 s, and the race
  step's longer run beside `cmd/sheets` and `internal/xlsx`, whose race
  runs take about 48 s each, so the race step takes no longer than on
  main.
- **Plain:** an edit takes about 11 ms, most of it sheets' own file writes.
  Only 8 of the 2,000 edits are made again by the repeat check (the
  edits that remove several dynamic arrays). Counting arrays that only
  moved as removed too would repeat 124 of them and add 14 s.
- **Under `-race`:** an edit takes about 160 ms.
- **Under heavy load the bar is missed.** With the Mac busy with other
  sessions (a load average of about 7), `portal check full` took 153 s
  (1,028 tests): the plain fuzz 60 s (24 s alone), and the race fuzz 79 s,
  held up by its 200-edit minimum at about 390 ms an edit. Main took 64 s
  at a load of about 5. The fuzz runs are one thread each, so a busy Mac
  slows them more than the packages that run in parallel; an unloaded Mac
  gives the 94 to 96 s above.
- **If check full ever nears 135 s:** lower the race cap first, never
  below 200 edits. If 2,000 plain edits still don't fit, report the
  numbers. Never weaken the invariants.

## How the fuzzer was checked

Besides the must-refuse cases, one mutation of sheets was run and
reverted, to see the fuzzer fail: with `keepEmptyText`
(`xlsx/recalc.go`, the fix of wave 1's empty-text regression, e7a9d5b)
returning at once, the plain run reports dozens of `editors` findings, such
as `Data!L2: stream holds text "", full holds nothing (pending)`. With the
formulas' results all numbers (the first version's `+N("p")` marker), it
could not.
