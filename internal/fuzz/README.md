# The fuzzer

A machine doing what wave 1's audits did by hand. It runs the sheets CLI in
process (`cli.Main`) over random workbooks from openpyxl and XlsxWriter. It
makes random edits and checks fixed invariants on the saved files after
each one. The default mode runs in every `go test ./...`; long mode
(`SHEETS_FUZZ=long`) runs for a time over pools it makes fresh; and seven
mutants put wave 1's bugs back into sheets, to show the fuzzer finds each
(How the fuzzer was checked).

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
- **A hang is a finding.** Each edit, with both editors and the checks made
  after it, must finish in 30 seconds (it takes milliseconds: 180 ms under
  `-race`). One that does not is a finding of the invariant `hang`: the test
  prints the seed, the workbook and the edits that led to it, like any other
  finding, and the run stops there (the command's goroutine cannot be
  stopped, and goes on running in the process). A hang is not shrunk, since
  every replay would wait out the 30 seconds, and `known.txt` cannot list
  one. Before this, a command that hung ran until `go test`'s own timeout (10
  minutes, 30 in `check.sh`'s race step), which printed a dump of goroutines
  and no workbook, no edits, and the seed only in hex in a stack frame.
  `TestHangIsAFinding` holds the case. (The watchdog covers the edit: a hang
  in the fuzzer's own reading of a file, between edits, has none.)
- **A workbook gone is a finding.** An edit that deletes or moves the
  workbook (data loss, the worst slip), or leaves nothing at its path that
  can be read, is a finding of the invariant `missing`, with the seed, the
  workbook and the edits, shrunk like any other, and the other sequences go
  on. The same file gone after `sheets undo` is an `undo` finding, and in the
  repeat check a `repeat` one. (Before, each was an error that stopped the
  test with one line, `open …/book.xlsx: no such file or directory`: no seed,
  no workbook, no edits, and the sequences after it never ran.) A `missing`
  finding can be listed in `known.txt`.
  `TestMissingWorkbookIsAFinding` holds the cases.
- **A panic outside the commands is a finding.** The commands are guarded by
  `run`. The rest of an edit's step is guarded too: the fuzzer's own judging,
  and sheets' reader on the files the editors saved (`compareEditors`, the
  model, the generator). A panic there is a finding of the invariant `panic`,
  with the stack in its detail, the seed, the workbook and the edits, and the
  run goes on. A panic of the generator on the workbook the last edit saved
  is that edit's finding. (Before, the test binary crashed with a dump of
  goroutines, no seed and no workbook, and every later test of the package
  was skipped: the failure that matters most is a workbook sheets' own reader
  cannot read.) `TestPanicInJudgingIsAFinding` and
  `TestGeneratorPanicIsAFinding` hold the cases, with a reader and a
  generator that panic.
- **The log** (`go test -v ./internal/fuzz/`) prints the edit count, the
  edits applied, each finding's signature, how many saved files held a
  formula result of each stored form (see the invariants), how many
  edits the streaming editor could take itself (see The streaming editor),
  how many texts it could write itself, by what they hold, and how many sorts
  were made and taken by both editors.
  With one seed, two runs print the same.

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
- **Texts that need escaping.** What `set`, `write` and scripts type holds,
  in about half of the texts, what a writer must escape or quote: `&`, `<`,
  `>`, double quotes, an apostrophe, spaces at the ends (`xml:space="preserve"`),
  `_x0041_` and the like (what Excel reads as an escape), and non-ASCII
  (`Åäö 日本語 😀`). A `set` and a script also type a line break, a tab and
  control characters, which XML cannot hold (Excel stores them as `_x0001_`);
  a batch and a `write` do not, since one is a line to a command and the other
  a row of a tab-separated table. `replace` finds and puts in such texts too,
  and new sheets are named `R&D 2`, `Bob's`, `A<B` and `Q1 "draft"`.
  Without them the streaming editor's own text writer
  (`xlsx/stream_sheet.go`, `newCell`: `escapeOOXML` and `xml:space`) was never
  reached: with its escaping taken out, the default run still passed
  (wave 2's audit, f2.1). The log counts what the streaming editor could write
  itself, by what the text holds, and the test fails when a plain run puts it
  fewer than 5 of each.
- **Sorts that sheets does.** Of the 137 sorts the first version made, at
  random, 97 were refused by both editors (the sheet holds several tables, the
  range takes in a dynamic array or a legacy array, a formula in it reads
  another sheet by its row), so the sort class was barely judged: 40 reached
  the invariants. Now 85 in 100 are picked from what a model of sheets'
  refusals (`sortOK` in `edits.go`; since WP-18 it lets through a dynamic
  array in one cell, which sorts with its row as in Excel) passes: the data
  rows of an Excel table (by its name, with its column names, or as a
  range), or the rows under the
  header of the block at the sheet's top left, over some of its leftmost
  columns, by a column with no formula first. The other 15 in 100 are random,
  as before, so the refusals are still exercised. 320 of 320 sorts the model
  picked were done by sheets (`TestSortsAreDone`). The log counts the sorts made
  and taken by both editors, and the test fails when a plain run has fewer
  than 60 taken, or fewer taken than refused.
- **Formulas typed.** A formula with a row in it (`B{r}*2`, `A{r}&"-x"`) is
  typed with the row of the cell it goes in. For a range, that is the range's
  first cell: `set 'Data!U9:U12' '=ROUND(C9*1.1,2)'` fills the rest down and
  across with the references shifted, as Excel does. (A range set once typed
  the template's own `{r}`, `=IF(N("p")=0,ROUND(C{r}*1.1,2))`, which sheets
  accepts, reports as a circular reference, and Excel cannot parse, in 54
  lines of the default run: the class this is for, a relative formula set over
  a range, was not exercised, and the later edits of those 37 sequences ran on
  workbooks Excel would repair.) `TestNoPlaceholderTyped` generates 6,000
  edits and fails on a literal `{r}` or `{s}` anywhere, and on a range set
  whose formula is of another row than the range's first. Of those, 100 are
  range sets of a relative formula, 88 over several rows. The default run's
  findings are the same 34 with the formulas typed as they should be.
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
   misses one the cell has, writes the first.) A text (`<t>`) with white space
   at its ends must say `xml:space="preserve"`, or Excel drops the spaces:
   sheets' own reader keeps them, so a writer that forgot the attribute was
   never noticed. Every worksheet lists its
   rows in order, each once, and the cells of each row in order, each once
   (`internal/xlsxtest`).
2. **kinds:** every formula keeps its kind: plain, legacy array (entered
   with Ctrl+Shift+Enter) or dynamic array. An array's range must start at
   its cell, and a dynamic array's cell metadata must exist. And every
   formula is where the edit put it: the model below follows each formula
   cell through the edit, so a formula an edit loses, or one it makes out of
   nothing, is a finding. (Before, only the formulas still in the file were
   judged.)
3. **reach:** nothing outside the edit's reach changed. A model of each edit
   (`reach.go`, driven by the model of kinds) says which cells it may
   change: the cells a `set`, `write` or `clear` names, a `copy`'s target, a
   `move`'s source and target, the range a `sort` takes (the whole sheet when
   it names a sheet or a table), the values a `replace` finds (as sheets'
   replace reads them, `grid.Display`: a date as 2026-07-01, 12%; a number
   with no format too; ignoring case: the whole find must be in one of these
   texts), the rows or columns an insert adds, and a sheet copied. Every
   other cell is followed through the edit (rows and columns inserted or
   deleted move it, a sheet renamed takes it along) and must hold what it
   held, where it stands now; an empty one must stay empty, and no sheet may
   appear that the edit does not name. It must be shown as it was, too: a
   cell outside the reach keeps its number format (the format, not its
   index), whether its value may change or not (a date shown as a plain
   46022 passed every invariant before: fuzzer-long's audit, f0.5). Two
   changes are expected:
   - **shifted references:** an edit that moves cells (move, rows and
     columns, a script's `moveTo`, `insert` and `delete`) or renames a sheet
     rewrites the formulas that read them, in their references only, and each
     reference as the edit's shift puts it. A model of each edit's rewrites
     (`xform` in `reach.go`, in order through a batch or a script) does what
     Excel does: a reference wholly inside a moved block follows it (to its
     sheet), one wholly inside the cells it lands on becomes `#REF!`, the
     rest stay; lines inserted move each end at or past them (a range across
     them grows); lines deleted make a cell in them, or a range wholly in
     them, `#REF!`, shrink a range partly in them and move the ends past
     them; a sheet renamed renames the references to it. Each reference of
     the formula saved must name the same cells as the model's (its sheet
     resolved, its dollars kept), and a reference left as it was where the
     model moves it is a finding too. Any other edit leaves every formula
     outside its reach as it was, letter for letter. The references are read
     by a scanner that skips strings and names (`LOG10`, `_xlfn.SEQUENCE`).
     **Defined names** are judged the same way (`checkNames`), since no edit
     names one: each name of the workbook before is there after, under its
     scope (a sheet renamed takes it along), its references as the edit's
     shifts put them (rows inserted in `Units` grow it, a move of `Data!C2`
     takes `Rate` along), and letter for letter after an edit that moves no
     cells; no name appears but one scoped to a sheet the edit copied.

     The first version took any rewrite of a formula's references, by any
     edit that moves cells. A slip in `formula.MoveRefs` that reads
     unqualified references as on the moved block's sheet made 109 wrong
     rewrites in the default run (the audit's count), and none was a reach
     finding (fuzzer-long's audit, f0.2). It is found now, as `move reach`,
     `do reach` and `script reach`: after
     `sheets move book.xlsx 'Q 2026!D3' 'Q 2026!A3'`, `Data!G2` holds
     `=IF(N("p")=0,'Q 2026'!B2*'Q 2026'!C2)`, and the edit's shift of
     references gives `=IF(N("p")=0,B2*C2)`.
   - **recomputed dependents:** a formula's value may change when it reads a
     cell whose value may (in the reach, or a formula whose value may, to a
     fixpoint), when its references were rewritten, when it uses a defined
     name the edit's shifts rewrote (rows deleted from `Units`), or when its
     result was unknown before (pending: sheets computed a result it did not
     know). A result known before and unknown after is a finding: the first
     version excused it too, which hid a save that drops results across the
     workbook (fix round 3, f2.2). An array's other
     cells change with its formula, and with an edit that reaches into it or
     cuts through it.

   So a slip that writes a cell the edit does not name, moves one it should
   not, loses or rewrites a formula elsewhere (or shifts a reference other
   than the edit does), recomputes one with nothing it reads changed, or
   shows a cell another way, is a finding. `TestReach` holds 78 cases, right
   and wrong (18 of them the shifts, 10 the names, 3 a replace's numbers, 2
   two dynamic arrays whose results meet),
   `TestReachFormats` the formats (and a replace reading a date),
   `TestRefToken` what the scanner reads, and `TestModelReadsGeneratedEdits`
   that the model reads every edit the generator makes on the pool. The
   default run makes no reach finding on today's code, with the shifts and
   the formats as with the first version. A scratch mutation in which
   `clear` takes one more column is found as `clear reach`: after
   `sheets clear book.xlsx 'Data!K1'`,
   `Data!L1: outside the edit's reach, it held text "Label", and holds nothing now`.
   WP-12's mutant (How the fuzzer was checked) is found by it first in the
   default run.

   Two slips of the model were found by running it, and fixed. Its first
   version took a defined name's range as it stands after the edit, and
   called `SUM(Units)` changed with nothing it reads changed when rows were
   deleted from `Units` (21 false findings in the default run): a formula
   that uses a name the edit defined again may change now. And it read the
   double quotes in a sheet's name (`'Q1 "draft"'!B9`) as a string, which hid
   the reference (8 false findings in long mode's 20 minutes, each of which
   replays clean now). The check costs no time a run can see (36 s for the
   default run, as before).

   Two holes the audit found, and closed in fix round 2. The model took any
   defined name the edit rewrote as expected, whatever the edit, and every
   formula using it as free to change: a slip that rewrote `Units` on a
   `set 'Data!Z1' 5` (`Data!$B$2:$B$2` to `Data!$B$2:$B$3`), and the
   `SUM(Units)` it took from 1 to 6, were no finding (f1.3). It is now:
   `the defined name Units, outside the edit's reach, referred to =Data!$B$2:$B$2, and refers to =Data!$B$2:$B$3 now, after an edit that moves no cells`.
   And a replace put every number in its reach when its find held a digit,
   a space, an `e` or a dot, as seven of the generator's 15 finds do
   (`Apple`, `N/A`, `Acme Inc.`): a `replace Apple Pear` that also wiped
   1200 was no finding (f1.4); it is now, as
   `Data!B1: outside the edit's reach, it held number 1200, and holds nothing now`.
   Neither change makes
   a new finding in the default run, and a scratch copy that calls every
   rewrite of a name a finding shows the default run rewrites names often
   (renames, inserts, moves, sheet copies), so the check is put to work.
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
judged, and so is a panic in judging the files it saved. One that does not
finish is a `hang`, and one that leaves no workbook at its path is `missing`
(see the default mode).

Besides these, two checks are the fuzzer's own: **text** and **repeat**.

**text:** a text a `set` or `write` typed is stored as typed: a text that is
no number, date, boolean, error or formula must read back from the saved file
exactly, with `_x0041_` read as Excel reads it. The editors invariant compares
one editor with the other, and both write text through the code that escapes
it (`ooxmlText`), so a slip in it (`_x0041_` stored unescaped, which Excel
reads as `A`) gave the same file twice; the xml invariant sees only a file
that is not well-formed. Only the top-level `set` and `write` are judged (in a
batch or a script a later command may move what an earlier one typed), and a
table's header is not (the table names its columns after it). It found a bug
on today's code (see Findings).

**repeat:** an edit that removes a dynamic array (net of those it adds on
the sheet) is made again, 20 times
with each editor, on the bytes it first found, and a workbook that holds
other cells, formulas or values than the first is a finding. Sheets lays out
the results of dynamic arrays in map order, so a paste over arrays gives a
different workbook from one run to the next (see Findings). A finding that
comes and goes would make two runs of one seed differ, so the check must see
such an edit nearly every time.
- **How often it misses:** the shrunk paste of `copy repeat 100139` gives
  one workbook 73% of the time and two others about 14% each. The check
  misses it when all its 42 runs (the two editors' first runs and 20 repeats
  each) give the common one: 0.73 to the 42nd power, under 0.0002%. With 8
  repeats it was 0.35%: one replay in 300 printed `no finding`.
- **600 replays** of sequence 100139
  (`SHEETS_FUZZ_REPRO=100139 go test -count=600 -run 'TestFuzz$' -v ./internal/fuzz/`)
  found it 600 times (`copy repeat 100139`), none printed `no finding`.
- **Under `-race`** the repeats are 3, not 20: each edit takes ten times as
  long, and a miss only drops a finding that the run lists as stale.
- **Byte-level differences are not compared:** the order of the shared
  strings a spill adds varies and means nothing.
- **Only these edits are made again:** all of them would cost a third of the
  run. A paste over one array was missed while only edits that removed two
  or more were repeated (`copy repeat 100164` in an earlier pool: one array
  overwritten, found in 4 of 12 replays by the editors' comparison alone).

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

### Where the formulas go

The marker says what kind a formula is, not whether it is still there.
`model.go` follows every formula cell through the edit (it starts from the
cells that hold a formula before it, applies what the edit does, and compares
with the saved file, cell for cell and kind for kind):

- **set, write, script setValue and setFormula:** a cell takes a formula
  when the value typed is one, else it holds none. (A formula typed in an
  Excel table's header is not kept: sheets names the column after it, as
  Excel does.)
- **clear:** the cells hold none.
- **copy:** each cell of the target holds what its source cell did, a
  formula of the same kind or none (`--values`: none; `--format`:
  unchanged). A paste of an empty cell empties the target's formula.
- **move:** the same, and the cells it empties hold none.
- **rows and cols, insert and delete:** formulas go where their cells go;
  those in the rows or columns deleted are gone. A single row or column
  (`Data!4`, `Data!C`, what an insert names) is read as the command reads it;
  the model once gave up on it, as grid.ParseRef reads no range there, and
  so judged no insert, nor any batch holding one (fuzzer-core's audit, f4.1).
- **sheet rename, copy and add:** the formulas stay, and a copy holds the
  same as its source. What the model knows of a sheet goes with its new
  name: that it was sorted, and that the model lost it (a sheet sorted,
  edited and renamed was judged by column under its new name: f4.2).
- **replace:** no cell gains or loses a formula.
- **sort:** no column gains or loses a formula of a kind (the rows move,
  and the model does not say where). A sheet sorted in a batch or a script
  and edited afterwards is not judged.
- **do and script:** the model of each command or statement, in order. An
  edit with a command the model has no rule for is not judged.

A paste that lost the second formula of three (`copy Data!G2:G4 Data!W2`
leaving W3 empty) is a finding of the kinds invariant; `TestModel` holds
the cases, right and wrong, and the run's mutation (see How the fuzzer was
checked) puts that bug into sheets.

## The streaming editor

The editors invariant compares the streaming editor with the full one, but
the streaming editor does very little itself. By design it does `set` and the
insert and delete of rows, and hands every other edit over to the full editor
(`xlsx/auto.go`), so for those the two copies run the same code and the
comparison proves nothing. It hands over `set` and rows edits too:

- **`set`:** once the workbook holds a dynamic array, when the value typed is
  one, when the cell is in a legacy array over several cells, and when the
  sheet is new;
- **rows:** once the workbook holds a dynamic array or a legacy array over
  several cells (`anySpills`, `anyLegacy` in `xlsx/stream.go`).

With arrays in nearly every pool file, the first version's streaming editor
took 76 of about 1,700 edits itself (set 38 of 238, write 19 of 133, rows
insert 9 of 121, rows delete 4 of 149, script 6 of 72, none of the other
classes). Two streaming-only bugs in the deletion of rows, references
rewritten as if one row were deleted and the rows after a multi-row delete
shifted by one, passed the default run (sheets' own `TestStreamMatchesFull`
failed on each).

So the pool and the edits aim at it:

- **The pool:** 17 of the 40 files hold no dynamic array and no legacy array
  over several cells (a coin in each file's plan, half the files, and the
  arrays of the others are not certain either), so the streaming editor
  takes `set` and rows edits on them. They keep single-cell legacy arrays,
  tables, formulas of every kind but those, charts and the rest.
- **The edits:** while the workbook is one the streaming editor takes (no
  dynamic array, no multi-cell legacy array: `streamable`), `set`, `rows`
  insert and delete are weighted three times, `write` and `script` twice, and
  a formula that spills is rarer in what `set`, `write` and scripts add (3%
  of what they add, against 20%). The workbook stays one it takes for
  longer. The other classes still run on it: they are full-editor code in
  both copies.
- **The count:** the log prints how many edits the streaming editor could take
  itself, by class (`streamed`: a `set` or `write` on a sheet that exists
  that adds no dynamic array, or rows, on a workbook that is one it takes).
  It is the fuzzer's estimate from the edit and the workbook, and it leaves
  out `set` and rows lines inside `do` batches and scripts. The test fails
  below 300 in a plain run.
- **Checked against the editor:** with a counter put into `Auto.Save` of a
  scratch copy (not committed), the streaming editor saved 461 of the run's
  1,884 applied edits itself: set 175, write 56, rows insert 77, rows delete
  89, do 29, script 35. The estimate is 366 of them (the 77 and 89 of rows
  are exact; it counts no `do` or script, and falls short on `set` (153 of 175)
  and `write` (47 of 56)).

The two bugs above, and one in the other classes, are among the mutations in
How the fuzzer was checked: each is caught.

## The pool

There are 40 small workbooks, about 338 KB in all, in `testdata/pool/`:
20 written by openpyxl and 20 by XlsxWriter. They come from
`scripts/fuzzgen.py`, each from its own seed. Each holds:

- values of every type, among them texts that need escaping (`R&D`,
  `<b>x</b>`, `Smith & Sons`, `say "hi"`, `it's`, ` padded `, `Åäö 日本語`),
  in column A of `Data` and in the keys of the second sheets;
- a filled-down formula column, and total rows. In about two files in three
  the column is rewritten as shared formulas (one group or two: the first
  cell holds the text and the range, the others point to it), as Excel and
  LibreOffice store a formula filled down; neither producer writes them;
- formulas whose results are not numbers: texts, empty texts
  (`IF(B2>1000,1,"")`) and booleans, filled down, on second sheets, and in a
  legacy and a dynamic array;
- plain formulas reading a range where one value goes
  (`='Data'!C2:C3*10`);
- legacy arrays over one cell, and, in some of the files, over several;
- dynamic arrays, in some of the XlsxWriter files (openpyxl cannot write
  them);
- second sheets reading the first, some with spaces in their names and some
  with `&` or an apostrophe (`R&D`, `Bob's Plan`), which the workbook's parts
  and the formulas that quote them must escape;
- defined names, charts, validations and conditional formats;
- an Excel table in 19 of the files: over the data (A1:G, 9 files; sheets
  does not move cells into, out of or within a table, so the cells the
  formulas read cannot be moved there), or in a block of its own (W1:Y5,
  10 files). The block's tables are named `Block<n>`: the first version's
  `Aux<n>` read as a cell (column AUX), which Excel does not allow in a
  table's name, and a command given it took it for the cell.

**Arrays are in 23 of the 40 files** (13 of them with a dynamic array): the
other 17 hold none over several cells, so the streaming editor does its own
work on them (see The streaming editor).

XlsxWriter stores each formula's result as it does for one given a result: a
text as `t="str"`, a boolean as `t="b"`, an empty text as an empty `<v>`
with no type (which reads as no result) in half the files and, in the other
half, as Excel stores it, `t="str"` with an empty `<v>`. openpyxl stores no
result, which sheets reads as pending. So the pool holds stored and pending
results of every type but an error, and the edits add the rest.

To regenerate the pool:

    uv run --with openpyxl==3.1.5 --with xlsxwriter==3.3.1 python scripts/fuzzgen.py internal/fuzz/testdata

- **The versions are pinned:** a newer release may write other bytes, which
  would change every edit each seed makes and need `known.txt` made again.
  The script refuses a release but these two.
- **The output is the same each time:** byte for byte. The dates in the
  files are fixed: the document properties (openpyxl writes its own
  `modified` stamp over the one it is given) and the time of each part in
  the package. Regenerating with the pinned versions changes nothing.
- **The script reads each file back** and refuses one whose stored kinds
  don't match their markers. XlsxWriter, for one, stores any formula calling
  SEQUENCE as a dynamic array.
- **A different pool** (or `--seed`) makes the edits of each seed other ones,
  so `known.txt` must then be made again.

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

- **A fix removes its lines.** A run lists the entries for its seed whose
  sequences it made in full without making them (fixed), and never fails on
  them.

Findings on the code this was built on (99ba275), 34 in all:

- **14 × WP-18:** moving or copying a plain formula that reads a range where
  one value goes, or the cells it reads, stores it as a dynamic array that
  spills. This comes through `move` (8), `copy` (5), and `do` and `script`
  batches that move or copy. Moving the cells such a formula reads is now a
  third of the sources of copy and move, so a bug there (the one fixed in
  65873b1, put back, is caught) shows in the first minutes of a run.
  Fixed since (WP-18, kinds-spills): a moved or copied formula keeps its
  kind, and known.txt no longer lists them.
- **3 × new, `copy repeat`:** a paste over dynamic arrays gives a different
  workbook from one run to the next. The shrunk `copy repeat 100139` sets
  a dynamic array in `Q 2026!H6:H8` (`UPPER(A2:A4)`), then copies
  `Q 2026!A2:A4` onto `H6`: the pasted values are missing at random
  (`H7`, `H8`), whichever editor makes it. `spillResults` (`xlsx/spill.go`)
  lays the arrays out in map order. Replay it to see:
  `SHEETS_FUZZ_REPRO=100139 go test ./internal/fuzz -run 'TestFuzz$' -v`.
  Its sequence is the one picked for the audit's replays (see `seedMix` in
  `fuzz.go`; regenerating the pool for f2.1, with texts that need escaping,
  left it a paste over dynamic arrays).
- **7 × new, `sort editors`:** after a sort, the streaming editor stores the
  results of the formulas on the sheet it sorted (it computes them), and the
  full editor leaves them as they were: pending, or the 0 XlsxWriter stores.
  The two save different cells, whether the sort moved rows or found them in
  order. The shrunk `sort editors 100047` is
  `sheets sort book.xlsx 'Q 2026!A2:A4' A A` on `openpyxl-12.xlsx`:
  `Q 2026!B2` (`=SUM(Data!D2:D10)`) is 2088 in the streaming editor's file
  and has no result in the full editor's. The default run did not reach it
  before: of its 137 sorts, 97 were refused (see Sorts that sheets does).
  (The first version described it as a sort that finds its rows in order;
  six of these seven moved rows.)
- **10 × new, `set text` (3) and `write text` (7):** a text typed with two
  escape-like sequences that share an underscore, `_x005F_x0041_`, is stored
  as another text. `ooxmlText` (`xlsx/stream_sheet.go`) escapes each
  `_xHHHH_` with a regular expression, and the match for the first
  (`_x005F_`) uses the underscore the second starts with, so the second is
  left as it is: the file holds `_x005F_x005F_x0041_`, which Excel reads as
  `_x005FA`. Both editors store it so, since the code is shared; only the
  text invariant sees it. (Found by the text invariant, which was added
  because the editors invariant could not see a slip in that shared code.)
  Replay: `SHEETS_FUZZ_REPRO=100046 go test ./internal/fuzz -run 'TestFuzz$' -v`.

Other seeds find more (`SHEETS_FUZZ_SEED=<n>`, not the default mode), so they
are for filing as follow-ups too (both replay on the pool as it is now):

- **`script kinds`: rows, then a column, turn a legacy array dynamic.** On
  `xlsxwriter-11.xlsx`, a script that does
  `ws.getRange("2:2").insert("Down"); ws.getRange("O:O").insert("Right");`
  leaves `U3`, the shifted legacy array `UPPER(A3:A5)`, with the cell
  metadata `cm="1"` of the dynamic array that was in `U`: it is stored as a
  dynamic array.
- **`set editors`: a circular formula over an overwritten spill.** On
  `openpyxl-18.xlsx`: `set 'Backup!F7' =SEQUENCE(1,2)`, `Backup!F7` a date
  in the same command, then `set 'Backup!F7' '=NOT(F7)'`. The full editor
  stores F7 with `cm="1"` (a dynamic array's metadata on a plain formula);
  the streaming editor does not.

`SHEETS_FUZZ_SEED` and `SHEETS_FUZZ_EDITS` change the seed and the count.
That is not the default mode, and `known.txt` lists only the default seed's
findings of it. Long mode (below) runs other seeds on fresh pools.

**known.txt's lines:** `<edit class> <invariant> <seed>`, or several seeds of
one class and invariant on a line; a line that starts with `long` holds long
mode's findings, on the pools it made (the same class, invariant and seed on
the committed pool is another finding). **Fixed:** a run lists the entries
for its seed whose sequences it made in full without making them (fixed, or
no longer reached; the last sequence counts only when the edit count did not
cut it short). It never fails on them. The default mode does not list them
under `-race`, whose 3 repeats may miss what 20 see.

**Bug lines** name a known bug by its findings' face, for long mode, whose
fresh seeds make findings no list of seeds can hold:

    bug <name>: <class>[,<class>...] <invariant>; repro [long] <seed>; detail <pattern>[; edit <pattern>]

A finding shows the bug when its edit class is one of those, its invariant
is that one, the first line of its detail matches the detail pattern (a Go
regular expression), and with an edit pattern its last edit as a shell
command matches that too (`(?m)^sort |getSort\(\)`: a batch or a script
holding a sort). A bug may have several faces, one line each (the escape: a
`text` finding, and an `editors` one where a formula reads the text). Where
a face needs what the edit does not show, the invariant says it: a typed
text, or a formula lost, in a dynamic array's results (before the edit or
after), and a cell outside the reach that changed in one's results before
the edit, ends its detail with
`, in the results of the dynamic array at U2:U4`,
which the bug `spill-overwrites` reads (a face on what the cell held could
not take a text the array put there without hiding every slip of the text
writer). The repro is a sequence that shows it, `long` for
one on long mode's pools; a long run replays each at its end, up to 10 times
(the map-order bug's editors face shows one replay in two), and
`TestBugRepros` replays those on the committed pool in every `go test`
(all of them with `SHEETS_FUZZ=long`). A replay shows the bug when its
finding shows any of the bug's faces, not only its own line's: map-order's
script paste shows its editors face or its kinds face from one repro, as the
editors' layouts fall; held to its own face, the kinds line read as fixed in
one of two runs of `TestBugRepros` (`TestBugReproAnyFace`). A repro that no longer shows its bug
is reported as fixed, and one whose pool cannot be made as not replayed,
with that pool's own failure; neither fails a run or the suite, so a fix of
a listed bug passes before its lines are removed (the first version failed
`TestBugRepros` on it, `fixed? escape`, in every `go test`: fuzzer-long's
audit, f1.2; `TestBugsFixedReported` holds both). Faces are kept narrow, so a new bug of
the same class and invariant still reads as new, and so do the mutants':
`TestKnownFaces` holds real findings of each bug, and findings that no face
may take (WP-12's mutant's formula lost by a set on a new sheet, the
rows-insert mutant's formula reading a range turned dynamic, a dynamic array
lost by a move, a text typed and stored upper case, a sheet's part not XML,
a script's paste with no array set before it). The default mode does not
read them: its findings are listed by seed, as before.

## Long mode

`SHEETS_FUZZ=long go test ./internal/fuzz -count=1 -v -timeout 30m` runs the
fuzzer for a time, not a count, over pools it makes fresh.

- **Fresh pools:** `scripts/fuzzgen.py --pool-only` makes a pool of 40
  workbooks (20 from openpyxl, 20 from XlsxWriter) for each round of 400
  sequences (about 2,600 edits, a minute), each from its own seed: the run's
  seed × 1000 plus the round (`PoolSeed`). A 20-minute run starts its
  sequences from about 800 workbooks, against the committed pool's 40. A pool
  is the same byte for byte from its seed (the releases are pinned), so a
  finding replays from its sequence's seed alone:
  `SHEETS_FUZZ=long SHEETS_FUZZ_REPRO=700012 go test ./internal/fuzz -run 'TestFuzz$' -v`
  makes pool 7000 again and replays sequence 12 of seed 7. A pool takes
  about half a second to make. (`--pool-only` leaves out the must-refuse
  cases, which are made from the committed pool's files: another seed may
  hold none that fits a case.)
- **The variables:** `SHEETS_FUZZ_TIME` (default 10m; the run finishes the
  sequence it is in), `SHEETS_FUZZ_SEED` (default a fresh one, printed),
  `SHEETS_FUZZ_EDITS` (a cap), `SHEETS_FUZZ_UNTIL=new` (stop at the first
  finding known.txt does not list: the mutants), `SHEETS_FUZZ_BASE` (with
  it, for the mutants: a copy of the tree without the mutant whose test
  binary is built; see The mutants), and `SHEETS_FUZZ_SOFFICE` (LibreOffice,
  when `soffice` is not on PATH).
- **The same edits and invariants** as the default mode. A finding is known
  when known.txt lists its signature or it shows one of the bugs known.txt
  names (bug lines, above); one that is neither fails the run, the first
  five shrunk. A fresh seed's findings of known bugs pass, then, and a new
  bug's stand out. (The first version knew long findings by their seeds
  alone, which only the seed they were listed from could ever meet: every run
  of a fresh seed failed on about 115 known findings per 10,000 edits, among
  which a new bug looked like any other; fuzzer-long's audit, f0.1.)
- **The log** gives the rate of findings per 10,000 edits, by edit class and
  invariant, and by the bug they show; the bugs whose repro no longer shows
  them (each bug line's repro is replayed at the run's end) and the entries
  of known.txt for the run's seed no longer made (fixed: reported, never a
  failure); and LibreOffice's second opinion.
- **No uv, or no network:** when uv cannot make a pool (or fetch openpyxl and
  XlsxWriter), that round runs on the committed pool instead, the log says
  why and for how many sequences, and their findings are not long ones. The
  bugs known.txt names are known there too, and a long entry whose sequence
  ran on the committed pool is not reported fixed: its own pool never ran.
  The repros on long mode's pools cannot be replayed, and the log says which,
  each with its own pool's failure. (The first version failed such a run on
  known bugs and called 12 long entries fixed, two of them bugs that still
  reproduce: f0.4. The second passed `TestFuzz` but failed the package:
  `TestBugRepros` erred once for each long repro, every line naming pool
  7006's command, the first to fail, and `TestFreshPools` failed: f1.5.)
  `TestPoolsFallBack`, `TestFixed` and `TestBugsFixedReported` hold the
  cases. `TestFreshPools` (with `SHEETS_FUZZ=long` only: it runs uv, and the
  default tests use no network) checks that a seed's pool is the same each
  time and another seed's another, with files from both producers that hold
  the invariants and name their kinds. When uv cannot run openpyxl and
  XlsxWriter at all, it is skipped and says why; a pool that cannot be made
  with uv working fails it. So `SHEETS_FUZZ=long go test ./internal/fuzz`
  with no uv passes, and says what it could not do.
- **LibreOffice's second opinion:** up to 120 workbooks the sequences leave,
  spread over the run, are recalculated by LibreOffice
  (`soffice --convert-to xlsx`, recalculating on load), and the values sheets stored
  for plain formulas (not arrays, not pending) are compared with LibreOffice's.
  Only workbooks that started from an openpyxl file are sent: openpyxl stores
  no results, so each one in them is sheets' own (XlsxWriter stores 0 for a
  number, which sheets keeps where no edit reaches). It is reported, never a
  finding: LibreOffice is not Excel, and it is no judge here. LibreOffice
  reads a `#REF!` in a formula's text as a name it does not know and gives
  `#NAME?`, so sheets' `#REF!` and its `#NAME?` count as agreeing.

**Seed 7, 20 minutes**
(`SHEETS_FUZZ=long SHEETS_FUZZ_SEED=7 SHEETS_FUZZ_TIME=20m go test ./internal/fuzz -count=1 -v -timeout 30m`,
2026-11-19, on Max's MacBook; the logs are in `~/sheets-evals/wave2/fuzz/`):

- **The first run** (`long-seed7-20m-baseline.log`): 70,934 edits (65,929
  applied) in 11,477 sequences, on 29 fresh pools; 864 findings, 121.8 per
  10,000 edits. Eight were the reach model's own slip (a sheet named with
  double quotes, fixed; each replays clean). known.txt listed the other 856
  by their seeds (it names their bugs now, in bug lines), by the bug each
  shows:
  - **540, known:** the `_x005F_x0041_` escape (`set text`, `write text`, and
    a formula reading such a text, which each editor computes from another
    text: `write editors`, `script editors`);
  - **278, known:** sort editors (2 of them a batch holding one sort);
  - **28, known:** the repeat (map order), by a paste, a move or a batch
    landing on dynamic arrays;
  - **2, known:** a script's rows and column turning a legacy array dynamic;
  - **6, new:** a control character typed into an Excel table's header is
    written raw into the table's part by the streaming editor
    (`xl/tables/table1.xml`: illegal character code U+0001). Shrunk,
    `long set xml 700021` is `sheets set book.xlsx 'Data!I4' … 'Data!B1'
    $'ctl\x01x' …` on a file whose table is over the data;
  - **1, new:** a text written into a dynamic array's results is lost to
    them (`long write text 711157`: `_x0041_` written into O3, which the
    dynamic array in O2 fills, is saved as the array's 18; Excel keeps what is
    typed, and shows `#SPILL!`; the bug `spill-overwrites`, with the formula
    seed 40016 found lost the same way);
  - **1, the repeat's after all:** a script that sets a dynamic array and
    then pastes over it (`ws.getRange("C2").copyFrom("A1:C4")`) saves other
    cells in each editor (`long script editors 701746`). It comes and goes,
    one replay in two, the pasted values missing in one editor's file or the
    other's: map order, seen by the editors invariant, as the repeat check
    does not run on an edit that removes no array net (the bug `map-order`,
    with the repeat's findings).
- **The run again, on the code as it is now** (`long-seed7-20m.log`, the Mac
  running the mutants beside it): 61,912 edits (57,623 applied) in 9,996
  sequences on 25 pools; 755 findings, all in known.txt (0 new), 121.9 per
  10,000 edits; no entry fixed; it passes.
- **The rate**, about 122 findings per 10,000 edits, is three known bugs
  nearly all: the escape (63%) and sort editors (32%), then the repeat.
  Without those three, about 1.4 per 10,000 are left (the 10 others of the
  first run).
- **LibreOffice's second opinion:** 117 workbooks each run, 4,465 and 4,449
  plain formulas with a result sheets stored; 4,450 and 4,433 agree, 15 and
  16 differ. Where they differ: a plain formula reading another sheet's range
  where one value goes (`=Data!C2:C3*10` in `R&D!C3`: sheets 684.9,
  LibreOffice 50, or `#VALUE!` elsewhere); a text with a line break
  (`=UPPER(A3)` of `line1\nline2`: LibreOffice drops the break); a boolean
  compared with a number (`=IF(B13>3,B13,0)` with B13 TRUE: sheets TRUE, as
  Excel, which ranks a boolean above any number; LibreOffice 0); `ISNUMBER`
  of a cell the two read otherwise; and `=SUM(Units)`, 121 against 122, once.
  Reported, not judged: they are for a look with the Excel oracle (WP-11).

**Fresh seeds** (fuzzer-long's fix round 1, 2026-11-19, on Max's MacBook,
several runs side by side; the logs are in `~/sheets-evals/wave2/fuzz/fix1/`
and `~/sheets-evals/wave2/fuzz/`). Before the bug lines, the audit's 3-minute
run of seed 40016 made 114 findings in 9,733 edits, every one of them new to
known.txt; the same run now makes 124 in 10,708, of which 123 show a known
bug. Each run's findings that no face took were looked at, and known.txt
named their bugs (or a new face of one), each with a repro:

| runs | edits | findings | per 10,000 | not known then | what they were |
|---|---|---|---|---|---|
| seed 40016, 3 minutes | 10,708 | 124 | 115.8 | 1 | spill-overwrites (a formula lost to a spill) |
| seeds 61001 to 61004, 6 minutes each | 52,176 | 628 | 120.4 | 3 | spill-overwrites twice (a text lost, a formula), insert-turns-dynamic |
| seeds 62001 and 62002 | 34,018 | 370 | 108.8 | 2 | map-order (two arrays meeting, by the editors invariant) |
| seeds 63001 and 63002 | 27,589 | 291 | 105.5 | 1 | replace-loses-kind |
| seeds 64001 and 64002 | 24,837 | 308 | 124.0 | 0 | |
| seed 341566, 20 minutes (fresh) | 49,048 | 555 | 113.2 | 4 | replace-loses-kind (named during the run), table-header-renamed, insert-turns-dynamic, spill-overwrites (a text the array put there) |
| seed 757171, 20 minutes (fresh) | 47,432 | 529 | 111.5 | 1 | sort-editors (the full editor's result empty, not pending: the face widened) |
| seed 341566 again, 20 minutes | 46,878 | 518 | 110.5 | 0 | |
| seed 664355, 3 minutes (the audit's, fix round 2) | 11,470 | 140 | 122.1 | 1 | insert-turns-dynamic, by a batch's cols insert (the face took only the rows and cols classes: f1.1) |
| seed 664355 again, 3 minutes (fix round 2) | 10,590 | 131 | 123.7 | 0 | |
| seeds 812207, 455903 and 290611, 3 minutes each, side by side (fix round 2) | 25,813 | 267 | 103.4 | 0 | |
| seed 548657, 20 minutes (fresh, fix round 2) | 74,582 | 837 | 112.2 | 1 | insert-turns-dynamic, a plain `=MIN(C2:E2)` turned dynamic by a rows insert (the face took no range at all: widened to an aggregate's) |
| seeds 173309, 926041, 384127, 519884 and 237760, 5 minutes each, two at a time beside a 20-minute run (fix round 2) | 56,285 | 690 | 122.6 | 0 | |
| seed 702553, 5 minutes (fix round 2) | 10,765 | 143 | 132.8 | 1 | full-keeps-array-metadata (new) |
| seed 915488, 20 minutes (fresh, fix round 2) | 50,025 | 582 | 116.3 | 2 | the reach model's own slip (two arrays whose results meet), intersection-kept (new) |
| seeds 611573, 847201, 158843 and 433019, 10 minutes each, two at a time beside a 20-minute run (fix round 2) | 89,020 | 1,025 | 115.1 | 0 | |
| seed 402252, 20 minutes (fresh, fix round 2) | 45,397 | 503 | 110.8 | 1 | spill-overwrites, a typed text with a quote in it (`say "hi"`: the face took none, widened) |
| seeds 771942 and 290155, 10 minutes each, side by side beside a 20-minute run (fix round 2) | 44,221 | 554 | 125.3 | 1 | intersection-kept, moved to a row its range does not cross (`#VALUE!`: the face took a number only, widened) |
| seeds 366284 and 905117, 10 minutes each, side by side beside a 20-minute run (fix round 2) | 43,849 | 512 | 116.8 | 1 | script-legacy-dynamic, by a script's column delete (the face took an insert only, widened) |
| seed 718113, 20 minutes (fresh, fix round 2) | 45,052 | 522 | 115.9 | 1 | spill-overwrites, a number typed into an array's results, seen by the reach invariant as its format lost (a new face) |
| seed 720212, 20 minutes (fresh, before map-order's paste faces) | 58,497 | 717 | 122.6 | 1 | map-order, a script's paste over a dynamic array an earlier edit set, a formula in one editor's file and none in the other's (a new face) |
| seed 720212 again, 20 minutes (with the editors face) | 57,912 | 711 | 122.8 | 1 | map-order, the same sequence, its formula lost in both editors' files, seen by the kinds invariant (a new face) |
| seed 720212 a third time, 20 minutes (the last commit's known.txt) | 55,240 | 677 | 122.6 | 0 | |

- **The rate** is about 115 findings per 10,000 edits, nearly all the
  escape (60 to 67% of each run's) and sort editors (32 to 36%), then map
  order (up to 3%). The other bugs made 30 findings in the 198,376 edits of
  the table's first six rows, about 1.5 per 10,000: table-header-control 10,
  spill-overwrites 7, script-legacy-dynamic 4, replace-loses-kind 3,
  insert-turns-dynamic 3, map-order's editors face 2 (counted here), and
  table-header-renamed 1, counting each once, as named now.
- **Two new bugs**, named in known.txt with a repro each, for follow-ups:
  - **full-keeps-array-metadata** (seed 702553): where a dynamic array's
    cell was written over and moved, the full editor keeps the array's cell
    metadata on a plain formula typed there and stores it as a dynamic
    array; the streaming editor stores it plain. A batch sets
    `'Bob''s Plan'!F2` to a dynamic array and then 1,200 in one command, a
    script inserts rows 2:3, and a write of a plain formula at F4 leaves F4 a
    dynamic array in the full editor's file alone (`long write editors 70255300604`).
  - **intersection-kept** (seed 915488): a plain formula whose range gives
    it one value by its row (`R&D!C2`, `=IF(N("p")=0,'Data'!C2:C3*2)`,
    reads `Data!C2`) keeps the result it had when an insert moves it to
    another row, where it reads another (in C3, `Data!C3*2`, 0.58, not 6). A
    set that computes it and a rows insert after it leave 6 in both editors'
    files, which no invariant can see (the reach model does not know a
    result depends on its row); a script shows it, since the full editor's
    computes the sheet again at its end and the streaming editor's does not
    (`long script editors 91548805029`). Moved to a row its range does not
    cross, it is `#VALUE!` in the full editor's file (seed 771942: a script
    deletes row 1 of Notes, and C1, from C2, `=IF(N("p")=0,'Data'!C2:C3*10)`,
    is `#VALUE!` there and 637.4 in the streaming one's).

  For their faces, the editors' differences say more: a value names the
  formula both editors hold (`R&D!C3 (=IF(N("p")=0,'Data'!C2:C3*2)): stream holds number 6, full holds number 0.58`),
  and the arrays end with those one editor has alone
  (`dynamic arrays: stream has F5, full has F4 F5 (the full editor's alone: F4)`),
  so a face can ask for a range read by its row, or an array the full editor
  adds rather than one the streaming editor loses (`TestCompareRanges`).
- **A slip of the reach model, fixed:** it took a cell in the results of
  two dynamic arrays that meet as one array's, so a rows delete that took
  the other array (`D6`'s `B2:B4*2` over D6:D8, `C8`'s `SEQUENCE(1,2)` over
  C8:D8) read as D8 lost (`long rows-delete reach 91548806198`). A cell
  changes with any array it is in now (Excel shows `#SPILL!` at one; sheets
  lays them out in map order), and the sequence replays as map-order.
- **A batch or a script shows a bug its single edits show** when it holds
  one of them: sort-editors and insert-turns-dynamic each have a face for
  `do` and `script` whose edit pattern asks for the edit in it (a sort; an
  insert or a delete of rows or columns), so a batch that holds none still
  reads as new. The audit's fresh seed, 664355, found insert-turns-dynamic
  through a batch (a script that sets `'Bob''s Plan'!E2` to a plain formula
  and deletes column A, then a batch's `cols insert 'Data 2!A' -n 2`: D2 is
  stored as a dynamic array), which the face for rows and cols alone did not
  take (f1.1). script-legacy-dynamic's face asks for an insert or a delete
  (seed 366284: a script sets `Data!I3` to `=B2:B4*2` and deletes column C,
  and the legacy array that moves into I3 is stored as a dynamic array); a
  script that moves the array (`moveTo`) still reads as new.
- **A bug seen by another invariant** gets a face for it. A number typed
  into a dynamic array's results is lost to them as a text is
  (spill-overwrites), but no invariant reads a number back; the cell keeps the
  number's format, though, and once the array no longer covers it, sheets
  drops the cell, format and all (seed 718113: set `Data!M3` 1,200 into the
  array at M2, then set M2 to `=SEQUENCE(1,2)`: M3, shown as `#,##0` before,
  is General now). The reach invariant says when a cell was in a dynamic
  array's results, as the text and kinds invariants do
  (`, in the results of the dynamic array at M2:M3`), and the face takes only
  a format lost there, to General.
- **map-order by a script's paste over an earlier edit's array:** a
  script's paste over a dynamic array that an earlier edit set (seed 720212: a set
  lays `=B2:B4*2` at `Data!H5`, over H5:H7, and a script's
  `ws.getRange("F4").copyFrom("H5:J8")` pastes J8's formula onto H7) leaves
  the formula in one editor's file and none in the other's, which side at
  random (two replays of four, either way round: `long script editors 72021203186`).
  When both editors lose it, the kinds invariant sees it (the same sequence,
  the run again: `Data!H6: the edit should leave a plain formula there, and it holds no formula (it was lost), in the results of the dynamic array at H5:H7`).
  The editors face for arrays set in the same edit asks for the array's
  `N("d")` in the script and for a value. map-order's two new faces each
  take a script holding a `copyFrom`: one whose cell holds a formula in one
  editor's file and none in the other's, and one whose formula is lost in a
  dynamic array's results in both. A paste in which both editors hold a
  formula, rewritten otherwise, still reads as new, and so do a formula a
  paste loses outside any array's results and a lost formula in a script
  with no paste.
- **Faces that must not hide a mutant:** the rows-insert mutant turns a
  formula that reads a range where one value goes dynamic (`=B3:B4*2`), so
  insert-turns-dynamic takes only one that reads no range but through an
  aggregate (`SUM`, `MIN`, `MAX`, `COUNTIF`, `AVERAGE`: seed 548657's
  `=MIN(C2:E2)`, which the first face, taking no colon at all, did not);
  WP-12's mutant loses a formula on a new sheet, in no
  array's results, so spill-overwrites takes only one in an array's results.
  The seven mutants are found with every face in place (The mutants).
- **The 20-minute runs** are `SHEETS_FUZZ=long SHEETS_FUZZ_TIME=20m go test ./internal/fuzz -count=1 -v -timeout 30m`
  (`long-fresh-20m-fix1.log`, seed 341566, and `long-fresh-20m-fix1-b.log`,
  seed 757171, each a fresh seed, 20 fresh pools of openpyxl and XlsxWriter
  files), and the first again with `SHEETS_FUZZ_SEED=341566` once known.txt
  named its four (`long-seed341566-20m-fix1-again.log`): it passes, 518
  findings in 46,878 edits, every one a known bug's, and each of the 13 bug
  lines' repros still shows its bug. The second's one finding no face took
  replays as sort-editors now. A fresh seed may still find a bug known.txt
  does not name, about one in 10,000 to 50,000 edits by these runs: that is
  what the run is for, and it fails until known.txt names it.
- **The last run** is seed 720212's third 20 minutes, on the last commit's
  known.txt (`long-seed720212-20m-third.log`): 55,240 edits (51,245 applied)
  in 9,059 sequences on 23 fresh pools; 677 findings, **122.6 per 10,000
  edits**, 0 new: escape 439, sort-editors 217, map-order 10,
  spill-overwrites 4, script-legacy-dynamic 3, table-header-control 3,
  insert-turns-dynamic 1. It passes, and each of the 19 bug lines' repros
  still shows its bug. The seed's first run (`long-fresh-20m-final.log`, a
  fresh seed) and its second (`long-seed720212-20m-again.log`) each made one
  finding no face took, the same sequence's two faces of map-order, which
  known.txt names now.
- **LibreOffice's second opinion**, in the three 20-minute runs of fix round
  1: 116 workbooks each; 3,984, 4,068 and 4,231 plain formulas with a result
  sheets stored, of which 15, 12 and 11 differ (the same kinds of difference
  as seed 7's). In the last run: 117 workbooks sent and recalculated, 4,381
  plain formulas with a result, 4,371 agree and 10 differ. Seven are a
  range read where one value goes (`=Data!C2:C3*2` in `Plan!C5`: sheets
  189.82, LibreOffice `#VALUE!`; or another row's value). Three are a `SUM`:
  1910 against 1911 once, and twice a `SUM` of another sheet's column, 0
  against 2 and 5. They are for the Excel oracle (WP-11) to settle.

## Timings

On Max's MacBook (M4), 2026-11-18, `scripts/check.sh full`, wall time, from a
clean copy of each commit. The bar is 135 s. Other sessions were running
tests on the Mac too (load average about 3), so the rounds alternate, main's
base commit (99ba275, without the fuzzer) then the branch.

| round | main (99ba275), without the fuzzer | with the fuzzer |
|---|---|---|
| 1 | 88 s (980 tests) | 104 s (1,139 tests) |
| 2 | 71 s | 104 s |
| 3 (after the audit's fourth round of fixes, load about 4) | 83 s | 118 s (1,159 tests) |
| 4 | 74 s | 122 s |

The Mac was busier for rounds 3 and 4 than for the first two. To tell the
load from the fixes, the commit before them (6bbf967) was timed beside the
branch, alternating: 121 s twice, against 126 s twice for the branch. So the
fixes cost about 5 s (the plain fuzz run 44 s to 47 s), and the rest of the
difference from rounds 1 and 2 is the load. Both stay under 135 s, with a
margin of 9 to 17 s at a load of 3 to 5.

- **Edit counts:** the plain run always makes 2,000 edits (1,884 applied)
  in 324 sequences, with 34 findings, all known, in 36 s alone. Two runs
  print the same (compared again for the audit's fixes: the log, minus the
  seconds, is identical).
- **The race count** is from the race step run with `-v`
  (`go test -race -count=1 -v ./...`, 67 s), at a load of about 3: 203 edits
  in 44 s. Alone, `go test -race -v ./internal/fuzz/` makes 203 too. It never
  makes fewer than 200: past the 40 s cap, the run goes on until 200 are made.
  Under -race an edit takes about 200 ms, so the cap and the minimum are near:
  a slower machine makes 200 and takes longer, not fewer.
- **What the fuzzer adds** is its plain run, about 34 s, and the tests that
  go with it (the sorts test takes 3 s, and skips under `-race`, where it
  would take 39 s), and the race step's longer run beside `cmd/sheets` and
  `internal/xlsx`, whose race runs take about 48 s each, so the race step takes
  about as long as on main. The package takes 47 s in the plain step and 59
  to 72 s in the race step (the cases that run the whole fuzzer, for the
  fixes of the audit's fourth round, skip under `-race`: they would add 13 s).
- **Plain:** an edit takes about 12 ms, most of it sheets' own file writes.
  34 of the 2,000 edits remove a dynamic array and are made again by the
  repeat check, 20 times by each editor: that is 9 s of the 34. (With 8
  repeats, and only for edits that removed two arrays, it was 1 s; the
  check missed the known paste in one replay of 300.) The watchdog costs
  nothing a run can see: a goroutine and a channel for each edit.
- **Under `-race`:** an edit takes about 200 ms, and the repeat check makes 3
  repeats, not 20.
- **Under heavy load the bar can be missed.** At a load average of about 7
  an earlier version (153 s, with half the repeats) took longer than 135 s:
  the plain fuzz 60 s, and the race fuzz 79 s, held up by its 200-edit
  minimum at about 390 ms an edit. The fuzz runs are one thread each, so a
  busy Mac slows them more than the packages that run in parallel.
- **If check full ever nears 135 s:** lower the race cap first, never
  below 200 edits; then the plain repeats (20 by each editor), which are 9
  s. If 2,000 plain edits still don't fit, report the numbers. Never weaken
  the invariants.

**With the reach invariant** (fuzzer-long, 2026-11-19),
`time scripts/check.sh full` on the same MacBook, alternating with this run's base commit
(5803d77: fuzzer-core and kinds-spills, without the reach invariant, long
mode and the mutants), other sessions running tests beside it:

| run | load average | check full | fuzz, plain | fuzz, race |
|---|---|---|---|---|
| branch (two tests not yet cut down) | 3.8 to 4.4 | 133.7 s (1,325 tests) | 52.7 s | 72.9 s |
| branch | 2.2 to 3.7 | 114.2 s (1,326 tests) | 46.1 s | 62.2 s |
| base (5803d77) | 3.4 to 4.9 | 117.6 s (1,186 tests) | 46.1 s | 59.7 s |
| branch | 4.6 to 3.6 | 124.2 s | 47.0 s | 65.5 s |

- **The first was 1.3 s from the bar.** `TestStopAtNew` and
  `TestPoolsFallBack` then ran 2,000 and 60 edits for what 3 and 12 show
  (Stop after a finding, a pool that cannot be made): about 4 s of the plain
  step. Cut down, check full takes about 5 s more than its base, with 11 to
  21 s to spare under 135 s.
- **The reach invariant costs no time a plain run can see:** the default
  run's 2,000 edits take 36 s, as before.
- **Under -race the run is at its floor:** 206 edits in 43.7 s (about 212 ms
  an edit, 200 before), its length set by the 200-edit minimum, not the 40 s
  cap. Lowering the cap would gain nothing, and the minimum is fuzzer-core's,
  so it stays.

**After fix round 1** (the reach invariant's shifts and formats, known.txt's
bug lines and `TestBugRepros`, `TestKnownFaces`, `TestUntilNew`),
`time scripts/check.sh full` on the same MacBook, 2026-11-19, at a load
average of 2.1 to 3.6: **115 s** (1,372 tests), the fuzz package 46.2 s in
the plain step and 62.3 s in the race step, 20 s under the bar. The shifts
and formats cost no time a run can see (the default run's 2,000 edits in
36 s, as before), and `TestBugRepros` replays three sequences on the
committed pool in under a second.

**After fix round 2** (defined names and a replace's numbers in the reach
invariant, `TestBugRepros` reporting, `TestBugsFixedReported`),
`time scripts/check.sh full` on the same MacBook, 2026-11-19, alone, at a
load average of 2.3 to 4.2: **114.6 s** (1,400 tests), the fuzz package
46.8 s in the plain step and 61.7 s in the race step, 20 s under the bar.
The first try took 133.1 s: `TestBugsFixedReported` replayed a fixed bug's
repro 10 times under `-race` too, 16 s there (1 s plain), and now checks that
half in the plain run only. Names and the replace's display cost no time a
run can see (the default run's 2,000 edits in 35.5 s).

**On the last code** (06fd17d: fix round 2's faces, the reach invariant
naming a dynamic array's results, script-legacy-dynamic's delete, map-order's
editors paste; the kinds face and `TestBugReproAnyFace` after it add 0.2 s
to the plain step and 0.4 s to the race step),
`time scripts/check.sh full` on the same MacBook, 2026-11-19, alone, at a
load average of 2.8 to 5.4: **118.5 s** (1,406 tests), the fuzz package
47.8 s in the plain step and 63.7 s in the race step, 16.5 s under the bar
(the log: `~/sheets-evals/wave2/fuzz/fix2/check-full-time-final.log`).

## How the fuzzer was checked

### The mutants

Wave 1's audits found seven bugs, each fixed since. Seven patches in
`mutants/` put them back into sheets' code, written against today's (after
kinds-spills: the rewritten move, copy and sort), each saying first what it
puts back:

| mutant | wave 1's fix | what it puts back |
|---|---|---|
| `wp12-sheet-rows` | 8e2e9cd (WP-12) | `writeWorkbook` no longer drops the sheets excelize holds after a write: a sheet made in a session saves with a cell twice, or rows out of order |
| `rows-insert-arrays` | 634b2ef | an edit of rows or columns writes the formulas it moved with `setFormula`, which reads the kind off the text: a plain formula reading a range where one value goes (`=B3:B4*2`) turns into a dynamic array |
| `rewrite-loses-kind` | 65873b1 | a formula a move rewrites where it stands is written again by its text (`rewriteFormula` ends with `setFormula`, not `putBack`) |
| `one-cell-array-moved-plain` | 2270292 | `take` no longer marks an array formula in one cell, so a move, copy or sort writes it plain |
| `move-leaves-formula` | 4410f38 | the cells a move empties keep the array formulas it took: a moved array formula stands in two places (see below) |
| `empty-text-lost` | e7a9d5b | `keepEmptyText` returns at once: the full editor saves an empty-text result as no result |
| `part-of-legacy-array` | c065389 | `checkArrays` no longer refuses a clear, copy or move that takes in part of a legacy array over several cells |

**4410f38 cannot be put back as it was.** The bug was excelize's: it kept an
array formula when `put` set a cell empty, and the fix took the formula away
first (`SetCellFormula(sheet, cell, "")`). excelize v2.11 takes it away itself
(`SetCellDefault` calls `removeFormula`), so taking the fix out changes
nothing today: the default run made no new finding with it. The mutant is the
nearest patch that puts the same bug back: the cells a move empties keep the
array formulas it took (in one cell, legacy or dynamic), as the old excelize
kept them, so the formula stands where it went and where it was.

`scripts/fuzz-mutants.sh OUT [--only NAME]` checks the fuzzer finds each. It
copies the working tree as it is (`git ls-files -co --exclude-standard`:
uncommitted work included; never HEAD, a branch or a worktree) into a
temporary folder: once as it is, the base, whose fuzz test binary it builds,
and for each mutant once more, with its patch applied. On the mutated copy it
runs long mode with seed 7 until the mutant's finding, for at most 4 minutes
(`SHEETS_FUZZ_UNTIL=new`, `SHEETS_FUZZ_TIME=4m`, `SHEETS_FUZZ_BASE`): the first
finding known.txt does not list, by its signature or as a bug it names,
whose sequence the base makes with no finding. A finding the code makes
without the mutant is logged as the code's (a bug for known.txt), and the
run goes on. It writes `OUT/NAME.log` for each (the mutant, the seed, the
shrunk repro, the time) and `OUT/summary.txt`, and exits 0 only if every
mutant is found. When uv cannot make long mode's pools, each mutant runs the
default mode on the committed pool instead (seed 1, at most 2,000 edits,
with the same stop and the same base), and the logs say so.

(The first version took the first finding known.txt did not list for the
mutant's, and never looked at the code without it. That held only while
known.txt listed exactly what seed 7 made on today's code: with seed 40016,
`rewrite-loses-kind` was "found" as `long sort editors 4001600015`, the known
sort bug, which the unpatched tree makes too; and a change to sheets that
moved seed 7's findings would have let every mutant pass with nothing found:
fuzzer-long's audit, f0.3. A known bug's findings are known now whatever
their seed, and the base must make the sequence clean.)

`TestMutantsApply` checks in every `go test ./...` that each patch
still applies to the tree (`git apply --check`), so a change to the code a
mutant patches fails the suite until the mutant is written again, putting
back the same bug; `TestMutantsApplyRefuses` holds a stale one.

**All seven are found**, each well within its 4 minutes: on Max's MacBook,
2026-11-19, `scripts/fuzz-mutants.sh ~/sheets-evals/wave2/fuzz-mutants` took
55 s in all, building each copy included, and exited 0 (the first round's
logs are in `~/sheets-evals/wave2/fuzz-mutants/first-round/`). Run again with
the check against the tree without the mutant (fuzzer-long's fix round 1),
it took 66 s, building the base's test binary included, and found the same
seven findings, each one whose sequence the base makes with no finding; no
finding was the code's. Run once more after fix round 2 (the reach
invariant's names and replace, insert-turns-dynamic's batch face), it took
66 s and found the same seven findings again, each the base makes clean.
And on the last commit, with every face known.txt holds, it took 57 s and
found the same seven again. The logs are in `~/sheets-evals/wave2/fuzz-mutants/`
(fix round 1's in `fix1-logs/` there, fix round 2's first run in
`fix2-logs/`). With the seed the audit used, 40016, `rewrite-loses-kind` is found as
`long move kinds 4001600026` after 3 s, which replays with no finding on the
tree as it is (`seed40016-rewrite-loses-kind/` there); the first version
took the known `long sort editors 4001600015` for it. Each finding of seed
7, as shrunk (replay one on the mutated copy with
`SHEETS_FUZZ=long SHEETS_FUZZ_REPRO=<seed>`):

| mutant | found after | finding | shrunk repro |
|---|---|---|---|
| `wp12-sheet-rows` | 15 s | `long set kinds 700123` | `sheets set book.xlsx 'Q1 "draft"!F7' '=IF(N("p")=0,IF(B7>3,"big",""))' 'Q1 "draft"!B9' … 'Q1 "draft"!A1' '=IF(N("d")=0,SEQUENCE(2))'`: a set on a new sheet that spills loses F7's formula |
| `rows-insert-arrays` | 1 s | `long cols-insert kinds 700000` | a batch that sets cells on `Backup` and renames it, then `sheets cols insert book.xlsx 'Q 2027!U' -n 2`: `Data!I2`, `=B3:B4*2`, is stored as a dynamic array |
| `rewrite-loses-kind` | 2 s | `long move kinds 700012` | `sheets move book.xlsx 'Data!C2:C3' 'Plan!H6'`: the formula reading them, rewritten `=Plan!H6:H7*2`, is stored as a dynamic array |
| `one-cell-array-moved-plain` | 1 s | `long copy kinds 700005` | `sheets copy book.xlsx 'Summary!D2:D3' 'Summary!D2'`: the one-cell legacy array in D2 is stored plain |
| `move-leaves-formula` | 1 s | `long move kinds 700006` | `sheets move book.xlsx 'Data!I2:J3' 'Old!A5'`: J3 still holds its legacy array |
| `empty-text-lost` | 2 s | `long write editors 700017` | `sheets cols delete book.xlsx 'Data!Q:Q'`, then a write: `Data!L2` holds an empty text in the streaming editor's file, no result in the full editor's |
| `part-of-legacy-array` | 14 s | `long move kinds 700119` | `sheets move book.xlsx 'Data!T2' 'Data!G7'`, the first cell of the legacy array T2:T3: G7 is stored plain |

The default run (seed 1, 2,000 edits) finds each of the seven as well
(counted in the first round), with 7, 496, 20, 19, 13, 72 and 3 new findings
in the order above
(`move-leaves-formula`'s first is `script kinds 100010`:
`ws.getRange("J4").moveTo("M3")` leaves J4's legacy array behind). Of WP-12's
seven, two are the reach invariant's, and the first by sequence
(`set reach 100026`, `do reach 100091`; then `sheet-copy xml 100029` and the
others).

### Earlier mutations

Besides the must-refuse cases, mutations of sheets were put into a clean
copy of the committed tree (never into the repo), each to see the default run
fail on findings that are not in `known.txt`:

| mutation | findings it adds to the 34 |
|---|---|
| `copy` drops the formula of the second pasted row (`xlsx/ranges.go`: `h.formula != "" && !o.Values && k[0] != 1`) | 17: 12 `copy kinds`, 3 `script kinds`, such as `Data!O11: the edit should leave a plain formula there, and it holds no formula (it was lost)` |
| the streaming editor deletes rows as if only one were deleted (`xlsx/stream.go`: `DeleteLines(…, op.from, op.from)`) | 21: 15 `rows-delete editors`, 1 `rows-delete kinds`, such as `Data!I3: stream has =IF(N("p")=0,B4:B5*10), full has =IF(N("p")=0,#REF!*10)` |
| the streaming editor shifts the rows after a multi-row delete by one (`rowNow`: `r -= op.to - op.from`) | 66: 49 `rows-delete editors`, 5 `rows-delete kinds` |
| a sheet's rows saved out of order after an in-session write (`xlsx/package.go`: the loop in `writeWorkbook` that drops excelize's cached sheets removed; wave 1's WP-12) | 6: 3 `sheet-copy xml` (`cell T3 after U3 in row 3`), a `sheet-copy kinds`, a `do xml` |
| a rewritten formula loses its kind (`xlsx/edit.go`: `rewriteFormula` calls `setFormula`, not `putBack`; the bug fixed in 65873b1) | 20: 17 `move kinds`, 3 `do kinds` |
| `keepEmptyText` (`xlsx/recalc.go`, the fix of wave 1's empty-text regression, e7a9d5b) returns at once | 73, mostly `set` (40) and `write` (13) `editors`, such as `Data!L2: stream holds nothing (pending), full holds text ""` |
| the streaming editor writes a text unescaped (`xlsx/stream_sheet.go`, `newCell`: `b.WriteString(v.Text)` for `b.Write(escapeOOXML(v.Text))`; wave 2's audit, f2.1) | 70: 21 `set xml` and 16 `write xml` (`invalid character entity & (no semicolon)`), 12 `write editors`, 7 `set editors`, 7 `script editors` |
| the streaming editor leaves out `xml:space="preserve"` (`newCell`: the test of spaces at the ends) | 24: 9 `write xml`, 8 `set xml`, 3 `script xml`: `the text " lead" has spaces at its ends and no xml:space="preserve": Excel drops them` |
| the `_x005F` of an escape-like text left out (`ooxmlText`, the code both editors share) | 17: 12 `write text`, 4 `set text`: `typed the text "A_x0020_B", and the cell holds text "A B"` |
| control characters written raw (`ooxmlText`, shared) | 26: 24 `set text`: `typed the text "ctl\x01x", and the cell holds text "ctl�x"` |
| `relines` (`xlsx/structure.go`) never returns for the columns of a sheet named `Summary` (f2.3's mutation: `select {}`) | 1 `cols-insert hang` after 30 s, with the seed, the workbook (`openpyxl-30.xlsx`) and the edits that led to it (`sheets cols insert book.xlsx 'Summary!A' -n 1`); the run stops there, in 3 sequences |
| `clear … --all` deletes the workbook (`internal/cli/root.go`: `defer os.Remove(args[1])`; wave 2's audit, f3.1) | 18 new: 17 `clear missing` (`sheets clear book.xlsx 'R&D!B7' --all`, shrunk to the one edit: `the edit (exit 0) left no workbook at book.xlsx`) and a `move kinds` of a later seed; the run goes on to its 2,000 edits |
| sheets' reader (`xlsx.ReadBytes`) panics on a workbook with a sheet named `Backup` (f3.2) | 33 new: 30 `panic` (`sheet-copy` 8, `do` 8, `sheet-rename` 6, `sheet-add` 6, `copy`, `write`), shrunk, such as `sheets sheet rename book.xlsx Data Backup`: `judging the files the editors saved: reader bug`, with the stack (`compareEditors`), and 3 findings of later seeds the shortened sequences let the run reach (`set text`, `write text`, `move kinds`); the run goes on to its 2,000 edits |

(The counts are from the scratch run of the default seed, which ends a
sequence at its first finding; they depend on the pool. A mutation that ends
sequences early lets the run reach later seeds, so a count holds a finding or
two of a seed past the default run's last: `move kinds`, `set text`.) The first
version could not see the first four: its model looked at the formulas a file
held, not at those it lost; its edits never put a `set` or a delete of rows to
the streaming editor on a workbook it takes; no batch edited a sheet it made;
and moving the cells a formula reads was rare. With the formulas' results all
numbers (the first version's `+N("p")` marker), it could not see the sixth
either. And the next four it could not see because its edits typed no text that
needs escaping: the streaming editor's text writer was never reached (the
seventh), and once it was, the editors invariant, which compares one editor
with the other, could not see a slip both make (the ninth and tenth), nor
sheets' own reader a missing `xml:space` (the eighth): that took a check of
the file (xml) and a check of the text typed (text).

The last two it could not see because the fuzzer's own failures ended the run
instead of being findings (wave 2's audit, f3.1 and f3.2): an edit that left
no workbook stopped the test with an error, and a panic in the files' judging
crashed the binary, each with no seed, no workbook and no edits. Both are
findings now, with the mutations above to show it. A third slip was in the
generator, not the checks: a range set typed a literal `{r}` (see Formulas
typed).
