# Is sheets useful to agents? Assessment, 2026-10-24

## Verdict

Agents pick sheets up reliably, and it is the fastest tool here and about
as cheap as Python. It is also the only one that leaves files with
formula results stored. But it does not yet make agents more successful:
with plain Python (openpyxl, pandas) Claude did every task, Sonnet and
Haiku alike, while sheets caused six failures of its own, three of them
wrong numbers reported with confidence. The first job is trust: sheets
must never store or report a wrong value. After that, its edge
(computed results, fidelity, checking) needs harder tasks to show, and
its install needs to stop requiring Go.

## How it was tested

Each task was done by fresh `claude -p` sessions in four arms, isolated
from user settings, skills, plugins and MCP servers:

- **sheets:** the CLI, with `sheets agent-docs` in the prompt (what the skill gives).
- **python:** openpyxl and pandas through `uv run`.
- **skill:** Python with Anthropic's xlsx skill. LibreOffice was not installed, so its `recalc.py` cannot run, as on most machines.
- **mcp:** haris-musa/excel-mcp-server 2.2.2 (openpyxl underneath, 46 tools, its own calculator for reads).

The tasks were evals/tasks 01–18. Six new ones aim where the tools should
differ:

- **13:** a downside case through a chained plan, with NPV.
- **14:** an edit to a dashboard with sparklines, a text box, Excel 2010 data bars and a chart.
- **15:** questions over a messy export (two header rows, subtotals, a hidden row, a figure stored as text).
- **16:** five planted errors in a budget.
- **17:** removing a repeated import from 150,000 rows.
- **18:** a price change in a German ERP CSV (semicolons, decimal commas, Windows-1252, CRLF).

Runs: Sonnet on all 18 tasks once, plus tasks 12–18 twice more; Haiku on
all 18 once. That is 50 runs per arm and 200 in all, for about $12.

## Results

Pass rates, with grading artifacts counted as passes (see the end):

| arm | Sonnet (32 runs) | Haiku (18) | all | stored results |
|---|---|---|---|---|
| sheets | 27/32 | 17/18 | **44/50** | **37/42** |
| python | 32/32 | 18/18 | **50/50** | 5/42 |
| skill | 32/32 | 18/18 | **50/50** | 4/42 |
| mcp | 32/32 | 16/18 | **48/50** | 0/41 |

Sonnet effort, summed over its 32 runs per arm:

| arm | turns | cost $ | seconds |
|---|---|---|---|
| sheets | 149 | 2.11 | 609 |
| python | 114 | 1.87 | 776 |
| skill | 151 | 2.51 | 930 |
| mcp | 374 | 3.41 | 1,109 |

**The six sheets failures were all caused by the tool:**

- **Task 13, four runs of four (Sonnet ×3, Haiku):** the agent reported
  an NPV of $310,202; the right figure is $1,190,870. sheets computed it
  wrongly and stored it (bug 1). The Python agents recomputed the plan in
  Python and got it right.
- **Task 12, one Sonnet run of three:** the agent wrote
  `=SUMIFS(Sales[Revenue],Sales[Region],A2)` into the Summary. That is
  sheets' own way to name a sheet's table, and Excel cannot read it as a
  formula (bug 14). sheets accepted it and said "Excel computes it on
  open"; `lint` passed it.
- **Task 18, one Sonnet run of three:** prices were written `84,5` for
  `84,50` (bug 5), while the agent's summary said two decimals were kept.
  Another run passed only by setting sheets aside and using `awk`.

**The other arms' failures:**

- **mcp, task 17 with Haiku:** the server's cell limits (100,000 per call,
  1,000,000 per copy) stopped it on 150,000 rows. It cannot compute
  results either, so it could not give the totals.
- **mcp, task 12 with Haiku:** named the file `Build report.xlsx`, a misreading of the prompt.

## What the runs showed

- **Plain Python is strong.** Agents wrote correct openpyxl code
  throughout. On the dashboard (task 14), Sonnet knew openpyxl drops
  sparklines, data bars and the chart, so it edited the package XML
  directly and even extended the chart. sheets' fidelity is real (openpyxl
  alone loses the text box, sparklines and 2010 data bars), but a strong
  model can work around openpyxl. It is riskier that way, not impossible.
- **Stored results are sheets' clearest advantage:** 37/42 against at
  most 5/42. Files from the other arms read as blank formulas in pandas
  and previews until Excel opens them, and the Python agents had to
  recompute plans by hand to answer what-if questions. That advantage is
  worth something only while the stored values are right.
- **Context use does not separate the arms.** Tool output over the first
  18 Sonnet tasks was about 14k tokens with sheets, 12k with Python, 13.5k
  with the skill and 15k with MCP: agents write targeted scripts and do
  not dump sheets. sheets adds about 1.7k tokens of guide per task and
  uses the fewest output tokens (19k against 24k for Python).
- **sheets is the fastest and about as cheap as Python.** The MCP server
  takes 2.5× sheets' turns and 1.6× its cost. Its tool schemas load on
  demand, its reads come in pages, and its cell caps bite on large files.
- **The xlsx skill adds nothing measurable over plain Python** without
  LibreOffice. Its agents tried `recalc.py`, which failed.
- **The tasks are too easy to separate the arms.** Python passed all 50.
  SpreadsheetBench 2 (June 2026) puts the best model at 35%, and at 12% on
  debugging, mostly from too little inspection and wrong target cells.
  Inspection, checking and computed results are where sheets should win,
  and these tasks barely test them.

## Does the skill trigger?

`evals/triggers.py` ran the 20 requests of `triggers.json` twice each, with
the skill installed in a fresh project (Sonnet):

- **Alone:** 20/20 requests on the right side. Every spreadsheet request
  loaded sheets; no near miss did.
- **Beside Anthropic's xlsx skill:** 20/20, with sheets chosen over xlsx
  every time. With xlsx renamed so that it is listed first, still 20/20.

## Bugs found

1. **Wrong NPV and AVEDEV stored.** excelize reads a range given to NPV or
   AVEDEV as its first cell: `NPV(0.09, A2:A6)` gives -91.74 where Excel
   gives -0.28, and `AVEDEV(A2:A6)` gives 104.8 where Excel gives 42.24.
   sheets counts both as computable, so it stores the result and reports
   it. A wrong stored result is worse than none, because pandas, previews
   and the agent all trust it.
2. **`AND(range)` reported as `#VALUE!`.** Excel gives TRUE or FALSE, so
   the agent goes looking for a mistake that is not there.
3. **Misleading "not evaluated" reason.** GCD, LCM, IMSUM and MULTINOMIAL
   over a range are left to Excel "because Excel answers with an array".
   In fact Excel takes a range there and returns one value.
4. **CSV corrupted after adding a column.** `sniff` picks the delimiter
   that splits the first ten lines most consistently, and keeps the comma
   on a tie. After `set` adds a column of decimal-comma numbers to a
   semicolon file (the header stays a field shorter), comma and semicolon
   tie. The next command, or the next line of a `do` batch, then reads the
   file as comma-separated and writes rows back mangled
   (`Mühle," elektrisch;…;3",45`). In that batch, `copy --values` and
   `cols delete` reported success and changed nothing.
5. **A CSV column's number style is not kept on new values.** A
   `1.249,00`-style column gets `1311,45` or `84,5`: the dialect survives,
   the style does not.
6. **Appending to an Excel table leaves its formula columns empty.**
   `write --append` adds the row without the table's calculated column
   (`SUM(Sales[@[Q1]:[Q4]])`); Excel fills it in. `lint` does not notice.
7. **`lint` false positives.** In a report with region heading rows, every
   subtotal is flagged "row 6 left out", where row 6 is the blank heading.
8. **`lint` misses common errors.** Of the budget's five planted errors it
   found three. It missed a total that also sums the total row (double
   counting), and a Summary line that points at the wrong row (its label
   does not match). It also missed a number stored as text in a report's
   second block.
9. **`query` sees only a sheet's first block.** On a report with three
   region blocks, a query "over all regions" silently covers one.
10. **`info` overstates an Excel table's extent:** an adjacent header
    joins the block.
11. **`read` prints float noise** (`24836587.2099999`), which agents may
    copy into answers.
12. **`next:` suggests `fmt` after appending to someone's formatted
    dashboard**, which the guide itself warns against.
13. **Installing needs a Go toolchain.** There are no releases and no CI.
    A static Linux build works (21 MB, no cgo), so prebuilt binaries are
    straightforward.
14. **sheets' address syntax leaks into formulas.** `Sales[Revenue]`
    names a sheet's table in sheets commands, but in a formula it means an
    Excel table called Sales.

## The alternatives

- **haris-musa/excel-mcp-server** (MIT, 4.3k stars, 2.2.2 on Oct 21):
  - *Strengths:* robust, and has more features than sheets (pivots,
    slicers, sparklines, images, goal seek, formula tracing).
  - *Weaknesses:* uses the most turns, cannot store results, and its cell
    caps stop it on large files.
  - *Where it fits:* the natural choice for a client with no shell (chat
    apps).
- **Anthropic's xlsx skill:** good guidance on top of openpyxl. Its
  results depend on LibreOffice, which is installed in claude.ai's sandbox
  but rarely on users' machines.
- **xlsx-for-ai:** a hosted service, so files leave the machine. Paid
  beyond a small free tier.

## Where to go next

1. **Trust (now).**
   - Fix bugs 1–4.
   - Test every function sheets computes in two forms, over a range and
     over the same cells listed. Mark any function that disagrees as one
     sheets does not compute.
   - Validate the calculator against Excel itself. Take a corpus of
     Excel-saved workbooks, recompute their formulas and compare with the
     results Excel stored. This is the test the main advantage rests on.
   - Make sheets turn its table syntax into ranges inside formulas, or
     refuse it. Have `lint` flag structured references to tables that do
     not exist.
   - Say "Excel computes it on open" only where that is true.
2. **Close the gaps the runs exposed:**
   - Fill calculated columns on append.
   - Fix lint's false positives, and add rules for double counting and
     mismatched labels.
   - Make `query` cover multi-block sheets.
   - Keep a CSV column's number style.
   - Allow an in-place transform of a column: on a CSV,
     `set D2:D41 '=ROUND(D{r}*1.05,2)'` should apply to the old values.
3. **Harder evals, so the advantage can show:**
   - Debugging and financial-modelling tasks on real Excel-made workbooks,
     in the style of SpreadsheetBench 2.
   - Multi-step sessions.
   - Several runs per task, and Opus as well as Sonnet and Haiku.
4. **Distribution:**
   - Release binaries (macOS, Linux, Windows; amd64 and arm64).
   - A Homebrew tap, and `npx` or `uvx` shims.
   - Point the skill's compatibility line at those.
   - Consider a `sheets mcp` mode for clients with no shell. The CLI
     stays the main way in: in these runs it took 2.5× fewer turns than
     the MCP server.

## Grading notes

Three failures were artifacts of these evals, counted as passes above:

- **mcp, task 05 (Sonnet):** its `SUMPRODUCT((range=A2)*range)` formulas
  are right, but the grader's calculator leaves array arithmetic to Excel.
- **Task 16, three runs (Sonnet python and mcp, Haiku mcp):** the figure
  planted as text was ten times its line's other months. The agents read
  it as a slipped decimal, entered 485, and said so in their answers. The
  task now plants a figure that fits its line.
- **skill, task 12:** showed negative gaps red with a `[Red]` number
  format. The check now accepts that as well as a conditional format.

Two Python runs hung for 12 minutes on a command waiting for stdin. Agents
now get an empty stdin, and those two runs are left out of the timings.
