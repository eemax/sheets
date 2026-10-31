# Vision

## Where sheets stands

On a single, short task, an agent with plain Python (openpyxl, pandas) does
about as well as one with sheets, and a little more cheaply: on our own 18
tasks it passed 54 of 54. Even on SpreadsheetBench, where sheets led 81 to 71,
most of the gap came from file-format knowledge (openpyxl leaving off the
`_xlfn.` prefix), which one line in a skill could teach.

So sheets does not win by helping an agent finish more tasks. That gap
closes as models and skills improve. It wins with what an agent cannot
rebuild inside one session:

- **A calculation engine.** It gives live values for formulas, without Excel.
- **Edits that follow Excel's rules.** References, tables, charts, names and
  conditional formats follow every structural edit.
- **Fidelity.** It keeps what a workbook already holds.
- **Checks.** It can tell whether the result is right.

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
| version control | `diff` and `undo`; history across sessions |
| code review | a receipt: what changed, rendered, with the checks that passed |
| a debugger | `explain`, `trace` and `what-if` |

## Identity: Excel for agents without Excel

sheets is for where no Excel runs: servers, pipelines, background agents,
CI. When a person has Excel open, the assistants inside Excel own that
moment. sheets owns the rest: a workbook changed headlessly, by an agent,
and handed to a person who has to trust it.

## Position: Python writes, sheets proves

Agents will always script, so sheets completes Python rather than competing
with it.

- **Building something new from data:** use whatever is natural, Python
  included.
- **Reading or changing a workbook someone already uses:** use sheets. It
  keeps the workbook whole.
- **Always finish with `sheets check`.** Whatever wrote the file, it computes
  and stores every result, fixes what Excel would reject, runs the checks,
  and says what's right, what's unknown and what's wrong.

The workbook goes to a person with a **receipt**: what changed (rendered),
which numbers moved and why, and the checks that passed. People are the
ones who decide which tools their agents use. A receipt is how they come to
require sheets.

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
| Silent-defect rate: the agent reports success, but the file is wrong | Trust: the number that matters |
| Python plus `sheets check` against Python alone | Whether sheets improves the workflows agents already use |
| Integrity after 20 edits in a row on a real workbook | Living workbooks |
| The reference workflows, run end to end | Real use |
| Pass rate on SpreadsheetBench and our tasks | No regressions |

**Kill criterion for the position.** If `sheets check` does not clearly cut
Python's silent defects, "Python writes, sheets proves" is the wrong wedge.
sheets then focuses on editing living workbooks.

## Horizons

- **Horizon 1, weeks:**
  - `sheets check` with receipts and expectations;
  - values that changed where no edit reached;
  - `explain`, `trace` and `what-if`;
  - the Excel check;
  - evals for the measures above.
- **Horizon 2, months:**
  - a session mode that keeps a workbook in memory;
  - a Python client on top of it, once Horizon 1 shows the position holds;
  - templates;
  - history, and change review: an agent proposes, a person sees it
    rendered and accepts.
- **Later:**
  - the engine in WebAssembly, for a browser that computes;
  - an MCP mode for chat clients without a shell;
  - perhaps other spreadsheet backends behind the same engine.

The detailed plan, as work packages, is in [roadmap.md](roadmap.md).
