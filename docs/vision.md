# Vision

## Where sheets stands

The core is changing a workbook people already use, the way Excel would,
and saying what the change did and what it cannot vouch for.

On a single, short task, an agent with plain Python (openpyxl, pandas) does
about as well as one with sheets, and a little more cheaply: on our own 18
tasks it passed 54 of 54. On SpreadsheetBench sheets led 81 to 71, but most
of the gap was file-format knowledge (openpyxl leaving off the `_xlfn.`
prefix), which one line in a skill could teach. So sheets does not win by
finishing more short tasks. It is meant to win where a script cannot follow:

- **A calculation engine.** It gives live values for formulas, without Excel.
- **Edits that follow Excel's rules.** References, tables, charts, names and
  conditional formats follow every structural edit.
- **Fidelity.** It keeps what a workbook already holds.
- **Checks.** It says what is right, what is unknown and what is wrong.

Today `sheets check` and `check --fix` exist. They compute and store
results, add missing prefixes and report what sheets cannot know. The
engine is being made to say unknown wherever it differs from Excel, and
formulas keep their kind through every edit. The change report is not
built yet. Whether the claims hold is not yet measured: wave 2 builds the
tests that can fail (see How we know it works).

Agents don't write their own compiler or their own git for each task. They
use tools whose correctness was engineered once. That's what sheets is for
spreadsheets.

## North star

**Make spreadsheets as safe for agents to change as code is.**

Agents are good at code partly because code comes with infrastructure.
Spreadsheets, the most widely used and least checked programs in business,
have none of it outside Excel's window. sheets is that infrastructure:

| Code has | sheets has |
|---|---|
| a compiler and runtime | the engine |
| a linter | `lint` |
| tests | expectations the workbook carries, checked after every change |
| version control | `diff` and `undo`; history across sessions through git (session mode is deferred) |
| code review | a receipt: what changed, rendered, with the checks that passed |
| a debugger | the change report now; `what-if` as `set --dry-run --show` later; `explain` later |

## Identity: changing the workbook people already use

The core is changing a workbook people already use, the way Excel would,
and saying what the change did and what it cannot vouch for. It means
references that follow, features that survive, values that are computed or
called unknown, and a report a person can read.

sheets is for where no Excel runs: servers, pipelines, background agents,
CI. When a person has Excel open, the assistants inside Excel own that
moment. sheets owns the rest: a workbook changed headlessly, by an agent,
and handed to a person who has to trust it.

## Position: living workbooks first, `check` as the on-ramp

Agents will always script, so sheets sits beside Python rather than
competing with it.

- **Building something new from data:** use whatever is natural, Python
  included.
- **Reading or changing a workbook someone already uses:** use sheets. It
  keeps the workbook whole, and reports what the change did.
- **On-ramp: Python writes, sheets proves.** Whatever wrote the file,
  `sheets check` computes and stores every result, fixes what Excel would
  reject, runs the checks, and says what's right, what's unknown and
  what's wrong. It is how a Python user first meets sheets.

The workbook goes to a person with a **change report**: what changed
(rendered), which numbers moved and why, what was lost, and what sheets
cannot vouch for. People decide which tools their agents use. A report is
how they come to require sheets.

## Principles

1. **Excel is the specification.** Its semantics, its file format, its
   quirks. Where sheets and Excel disagree, sheets is wrong.
2. **Unknown over wrong.** What sheets cannot be sure of, it says it does
   not know. It never guesses a value.
3. **Checks need no authoring.** Agents rarely write tests unprompted, so
   sheets finds what it can by itself:
   - formulas that break their column's pattern;
   - totals that leave out rows;
   - values that changed where no edit reached;
   - expectations it can infer from the structure.

   Written expectations come on top of these.
4. **Every change is visible and reversible:** diff, undo, receipts.
5. **Output costs little context.** What a model reads is sized for a model.
6. **One engine, one command set.** The CLI is the API. Scripts (Office
   Scripts in JavaScript, a Python client) map onto the same operations
   rather than growing object models of their own.

## Where it must clearly win

These are the reference workflows. They are measured on synthetic data and
drive the roadmap more than any benchmark does.

1. **Living workbooks.** Agents update a workbook over months while people
   also edit it. It stays whole, and every change can be seen and undone.
2. **Numbers that must be right.** Models and reports, with expectations,
   explanations and what-ifs.
3. **Recurring work.** A template filled once per product, a monthly
   refresh, an export to an ERP system and back.
4. **Messy intake.** Supplier and factory sheets with odd headers, read into
   a standard and checked against its rules.

## Not goals

- A spreadsheet user interface, or a hosted service.
- Running VBA, or writing legacy `.xls` files.
- Google Sheets or Numbers, until xlsx and CSV are done.
- Chasing benchmark percentages for their own sake. Coverage and benchmarks
  are polish, run alongside the roadmap.

## How we know it works

| Measure | What it shows |
|---|---|
| B: silent-defect rate and cost against the best Python arm, on a 20-edit session and a feature-loss task | Whether sheets is the better way to change a living workbook |
| C: `check --fix` against a one-line prefix hint, on a paired test | Whether `check` is worth pitching |
| Fuzzer findings per 10,000 edits | Trust: no file corrupted |
| Integrity after 20 edits in a row on a real workbook | Living workbooks |
| Pass rate on SpreadsheetBench and our tasks | No regressions |

**Kill criterion for the position.** Two numbered tests, each reported whatever its result:

1. **B's test.** sheets has at most half the best Python arm's defect rate,
   at no more than 1.3 times its cost. If not, the claim that sheets is the
   better way to change a workbook is not shown.
2. **C's test.** `check --fix` beats a one-line prefix hint on a paired
   test, or `check` stays a convenience and stops being the pitch.

## Horizons

- **Horizon 1, weeks:**
  - the trust of the engine and the kinds of formulas (wave 2, piece A);
  - the fuzzer: random workbooks, random edits, fixed invariants;
  - the change report: `check --against`, the short form after each edit,
    a Markdown receipt;
  - evals that can fail: B's and C's tests;
  - the Excel check.
- **Horizon 2, months:**
  - session mode is deferred, with no measured need;
  - the Python client is a thin wrapper over the CLI;
  - history goes through git: `sheets diff` as a diff driver, receipts on
    pull requests, `check` as a CI step;
  - `what-if` as `set --dry-run --show`, and `explain`;
  - templates, and change review: an agent proposes, a person sees it
    rendered and accepts.
- **Later:**
  - an MCP mode of four tools: look, change, check, render;
  - perhaps other spreadsheet backends behind the same engine.

WebAssembly is cut.

The detailed plan, as work packages, is in [roadmap.md](roadmap.md).
