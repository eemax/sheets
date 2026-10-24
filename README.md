# sheets

A command-line tool for spreadsheets, built for AI agents. It reads Excel
files without flooding an agent's context window, writes data with
sensible types, and applies an opinionated house style in one command.

```
$ sheets info q3.xlsx
q3.xlsx · xlsx · 1.1 MB · 1 sheet
Orders  A1:I20001 · 20000 formulas
  A1:I20001  table · 20000 rows · Order ID, Date, Region, Customer, Product, Units, Price, Discount, … (+1)
next: sheets peek q3.xlsx Orders

$ sheets peek q3.xlsx Orders --sample 0
Orders!A1:I20001 · table · header row 1 · 20000 rows
col  header    type        filled  values
A    Order ID  integer     20000   100000 … 119999
B    Date      date        20000   2026-01-01 … 2026-09-27
C    Region    text        20000   West, South, East, North
D    Customer  text        20000   "Customer 583", "Customer 484", "Customer 30", … 800 distinct
E    Product   text        20000   Doohickey, Gadget, Widget, Gizmo
F    Units     integer     20000   1 … 500 · sum 5014027
G    Price     currency $  20000   12.5, 7.25, 120, 30
H    Discount  percent     20000   5%, 0%, 10%, 15%
I    Revenue   decimal     20000   =F{r}*G{r}*(1-H{r}) · 6.1625 … 60000 · sum 195082686.525
next: sheets read q3.xlsx 'Orders!A1:I20'
```

## Install

```bash
go install github.com/eemax/sheets/cmd/sheets@latest
sheets skill --install    # teach Claude Code to use it (~/.claude/skills/sheets)
```

### The agent skill

[skills/sheets](skills/sheets) is an [Agent Skill](https://agentskills.io):
a `SKILL.md` that tells agents when to reach for `sheets` and how to use it
(the commands used most, a check before finishing, and the gotchas), and
`references/`, one file per topic of the guide, which the agent reads when
a task needs it. Only the skill's name and description stay in an agent's
context; the instructions (about 1,600 tokens) load when a task involves a
spreadsheet, and a reference (60 to 800 tokens) only when its topic comes
up.

| | |
|---|---|
| `sheets skill --install` | `~/.claude/skills/sheets`, for Claude Code |
| `sheets skill --install --agents` | `~/.agents/skills/sheets`, for VS Code, Codex and other agents |
| `sheets skill --install --project` | the current project's `.claude/skills` (with `--agents`, `.agents/skills`) |
| `sheets skill --install --dir DIR` | `DIR/sheets` |

The skill comes from the binary, so installing it again after updating
`sheets` keeps the two in step. Its frontmatter uses only the fields of the
specification, so it also loads where other fields are refused (claude.ai
uploads). A test checks it against the specification and keeps
`skills/sheets` the same as what `sheets skill` writes; after changing the
guide, run `go run ./cmd/sheets skill --install --dir skills`. To check it
with the specification's own validator:

```bash
uvx --from "git+https://github.com/agentskills/agentskills#subdirectory=skills-ref" skills-ref validate skills/sheets
```

Agents without the skill can run `sheets agent-docs` for the same
instructions, and `sheets agent-docs TOPIC` for a topic.

## Commands

| Command | What it does |
|---|---|
| `info <file>` | Lists sheets, used ranges, detected tables and titles, named ranges |
| `peek <file> [sheet\|range]` | Profiles each column: type, value range or distinct values, formula patterns, sample rows |
| `read <file> [range]` | Prints cells with row numbers and column letters, paged; `--formulas`, `--format tsv\|csv\|json` |
| `find <file> <text>` | Finds cells by value; `--regex`, `--exact`, `--formulas` |
| `query <file> [sql]` | SQL (SQLite) over the detected tables; `--schema`, `--into Sheet` writes the result back |
| `lint <file>` | Finds error values, broken references, inconsistent formulas, totals that miss rows, numbers stored as text, ...; exits 1 on errors |
| `diff <old> <new>`, `diff <file>` | Compares workbooks with rows and columns aligned first; with one file, shows what the last change did |
| `chart add\|list\|delete` | Adds charts in the house style, refusing ones that would mislead |
| `render <file> [sheet\|range]` | Draws a sheet to PNG (or HTML) with its formatting and charts, for agents that can look |
| `write <file> [sheet!cell]` | Writes TSV/CSV/JSON from stdin with type inference; `--append` (into the last table, above its total row; matches headers), `--clear` |
| `set <file> <cell> <value>...` | Sets cells or formulas; a range or a table column (`Sales[Margin]`) fills down with shifted references |
| `rows insert\|delete`, `cols insert\|delete` | Insert or delete rows and columns; references follow, and broken ones become `#REF!` and are reported |
| `sheet add\|rename\|delete\|copy\|move\|hide\|show` | Manage sheets; renaming updates every reference, deleting refuses while it is referenced |
| `fmt <file>` | Strips formatting and applies the house style, keeping what style, cond and view set; `--preset finance`, `--as COL=type`, `--check`, `--reset` |
| `style <file> <target> <prop>...` | Fills, fonts, number formats, borders, alignment, widths, merges on cells; `--show` prints how cells look, `--list` the record fmt keeps |
| `cond <file> <target> <condition>` | Conditional formats: comparisons, top/bottom, averages, duplicates, text, formulas, colour scales, data bars, icons; `--list`, `--clear` |
| `validate <file> <target> <rule>` | Data validation: dropdown lists, number, date and length limits, formulas; `--prompt`, `--error`, `--warn`, `--list`, `--clear` |
| `view <file> [sheet]` | Freeze panes, tab colour, gridlines, zoom, hidden rows and columns, filter buttons, active sheet |
| `sort <file> <target> <col>[:desc]...` | Sorts a table's rows as Excel does; refuses a sort that would change what a formula means |
| `copy <file> <from> <to>`, `move` | Copy (formulas shift; `--values`, `--format`) or cut and paste (references follow) |
| `clear <file> <target>` | Clears contents; `--format`, `--notes`, `--links`, `--all` |
| `replace <file> <find> <with> [target]` | Find and replace in values, or `--formulas`; `--exact`, `--regex` |
| `note`, `link` | Notes and hyperlinks on cells; `--list` |
| `convert <from> <to>` | Between .xlsx and .csv/.tsv |
| `do <file>` | Runs a script of commands against one file: one save, one undo step, nothing saved on an error |
| `undo <file>` | Restores the file to before the last change sheets made; `--list` |
| `agent-docs` | Prints the guide for agents: the skill's instructions, a topic (`agent-docs style`), or `all` ([agent-docs.md](internal/cli/agent-docs.md)) |
| `skill` | Prints the agent skill's `SKILL.md`, or installs the skill folder (`--install`; see [The agent skill](#the-agent-skill)) |

Every command takes `--json`, and every command that changes a file takes
`--dry-run`.

`style`, `cond` and `validate` take a range, a sheet, or a table reference
in Excel's syntax, named by a sheet holding one table or by an Excel table:
`Sales[Revenue]`, `Sales[[Units]:[Revenue]]`, `Sales[#Headers]`,
`Sales[#Totals]`, `Sales[[#Totals],[Units]]`. `set` takes a table
reference too, to fill exactly a column's data rows:
`sheets set book.xlsx 'Sales[Margin]' '=E{r}/D{r}'`.

```bash
sheets do report.xlsx <<'EOF'
write Sales!A1 -i sales.csv
fmt
style 'Sales[#Headers]' fill=lightblue
style 'Sales[Revenue]' format=currency decimals=0
cond 'Sales[Margin]' '< 0' color=red bold
validate 'Sales[Region]' list North,South,East,West
view Sales --freeze B2
EOF
```

### Formulas

- **Modern functions:** newer Excel functions (XLOOKUP, IFS, TEXTJOIN,
  LET, LAMBDA, ...) are stored with the prefixes Excel's file format needs;
  without them Excel shows `#NAME?`.
- **Dynamic arrays:** functions that return arrays (FILTER, UNIQUE, SORT,
  SEQUENCE, ...) are stored as dynamic arrays and spill.
- **Feedback:** after a write, `sheets` reports what each formula evaluates
  to:
  - a few sample values
  - every error, with the cell
  - formulas its calculator cannot evaluate, which Excel computes on open
- **Row placeholder:** `{r}` stands for "this row" in written formulas, the
  same notation `peek` uses for formula patterns. So
  `write --append` with `=C{r}*D{r}` lands right whatever row it ends up on.

### Query, lint, diff

- **`query`** turns every detected table into a SQLite table named after
  its sheet (or Excel table name), with header labels as columns. Numbers
  stay numbers, dates become ISO text, total rows are left out, and `_row`
  points back at the sheet row. `--into` writes the result as a new table
  keeping the source columns' number formats.
- **`lint`** checks for error values, `#REF!`, unknown functions, cells
  that break their column's formula pattern, total rows whose ranges miss
  rows, numbers stored as text, duplicate headers, merged cells in tables,
  hard-coded constants repeated in formulas, external links and volatile
  functions. `--fail-on` sets the severity that exits 1.
- **`diff`** aligns rows (patience diff over row contents) and columns
  before comparing cells, and detects renamed sheets, so an inserted row
  is one change. `sheets diff book.xlsx` compares with the undo history.

### Charts and render

- **Charts** come in column, bar, line, area, pie and scatter, plain or
  stacked, with defaults chosen from the data. The house style follows
  data-visualisation practice: a validated palette that holds up under
  colour-vision deficiency, one value axis, no legend for a single series,
  hairline gridlines, thin bars with a 2 px gap. More than 8 series, pies
  over 6 slices and mixed units on one axis are refused.
- **`render`** draws the sheet the way Excel shows it: number formats,
  fonts, fills, borders, merged and wrapped cells, banding, hidden rows,
  charts. It frames it with column letters and row numbers, so an agent
  can check its work visually.

### Undo

Every change keeps the previous version of the file (20 per file) in the
user cache directory (`SHEETS_CACHE_DIR` overrides it; `SHEETS_UNDO=off`
disables it). `sheets undo` restores the last one. It refuses if the file
was changed by something else since.

### Large files

A workbook over 1 MB is kept parsed in the user cache after its first
read, keyed by a hash of its contents and the `sheets` build. Later
commands on the same file skip the parse: `info` on an 11 MB,
200,000-row file goes from 1.4 s to 0.05 s. `SHEETS_CACHE=off` disables
the cache. It holds at most 32 files and 2 GB, dropping the least
recently used.

Edits stream. Writing values, inserting and deleting rows and appending
to a table rewrite only the sheets that change, row by row, and copy every
other part of the file as it is. On the 200,000-row file a one-cell edit
takes 1.2 s with the parse cached, 3 s on a first read (10 s through
excelize), and deleting 1,000 rows 3.1 s (4.6 s); on a 20,000-row file
deleting 1,000 rows takes 0.8 s instead of 19 s. The first edit that
streaming does not cover (styling, columns, sheets, copy/move/sort, notes;
also listing styling) hands the session to the full editor: what was done
so far is rendered in memory and the editor opens that.
`SHEETS_STREAM=never` turns streaming off, or gives a file size in bytes
from which to stream. A randomized test makes the same edits both ways on
a set of workbooks and checks that a reader sees the same files.

## The house style

`sheets fmt` throws away existing cell formatting and rebuilds it from the
contents:

- a dark header row with filter buttons, frozen in place; labels
  left-aligned so the buttons (at the right edge) never cover them
- alternate data rows banded light gray, using conditional formatting so
  the banding survives sorting
- thin light borders, gridlines off
- a bold total row with a heavier rule above it
- titles above a table in bold 14 pt
- two-row headers styled as one, group labels centred across their columns
- column widths fitted to their contents, long text wrapped
- number formats chosen per column:
  - integers and decimals: thousands separators, negatives in parentheses
  - currency: whole amounts, or two decimals when there are cents
  - percentages: up to two decimals
  - dates: ISO
  - ids and years: no separators

The `finance` preset adds the banking color code: blue for typed-in
numbers, black for formulas, green for references to other sheets.

Formatting depends only on the cell contents and the recorded styling
(below), so `fmt` is idempotent and `fmt --check` works in CI. Column types
are inferred, and `--as Revenue=currency` overrides them. The choice is
saved in the file, so later runs keep it.

### Styling on request

`style`, `cond` and `view` change what the user asks for on top of the
house style, and record it in the workbook so that `fmt` keeps it: each
entry is a hidden defined name whose range follows inserted and deleted
rows and columns (in Excel too), with what was applied as its comment
(`_sheets.000003` → `Sales!$A$1:$F$1`, "style fill=lightblue bold"). A
recorded freeze moves as its pane does, and settings that are not cells
(gridlines) stay put. Every
`fmt` strips formatting as before, then lays the recorded styling over the
house style; recorded conditional formats stay where they are.
`fmt --reset` drops the record.

- A fill without a text colour gets black or white text, whichever has the
  contrast to read.
- A number format given a table column (`Sales[Revenue]`) reaches its total
  row. Rows added with `write --append` take the look, conditional formats
  and validation of the rows above, as Excel extends a list's formats.
- Conditional format formulas follow inserted and deleted rows and columns,
  as in Excel, and highlights take priority over the banding.
- `style --show` describes cells in the same properties `style` takes, so
  an agent can check a result without rendering it.

Formula columns take the type of what they compute from: `=Units*Price` is
currency when Price is, and `=Start+30` is a date. Divisions and date
differences stay plain numbers.

## Design

- `internal/grid` is a model of a workbook that doesn't depend on any file
  format. Analysis and output only see this model. `internal/book` picks the
  implementation for a file's extension. CSV (`internal/delim`) is read
  straight into the model, and edited through an .xlsx workbook in memory
  that is written back in the file's own dialect: rows whose values no edit
  changed are written as they were, byte for byte.
- `internal/xlsx` reads .xlsx with its own streaming XML parser. It is fast
  (`peek` on 20k rows with 20k formulas takes about half a second) and
  sees shared formulas and cached values as stored. Editing and styling go
  through [excelize](https://github.com/qax-os/excelize), with workarounds
  for places where excelize would corrupt or lose data:
  - filled-down formulas are kept when their first cell is overwritten
  - the sheet `dimension` is kept accurate
  - sheet properties are restored after adding a filter
  - formatting is stripped only from cells that actually carry a style, so
    a stray value at row 1,000,000 doesn't make `fmt` touch a million rows
  - Excel tables are found even when their path is absolute, which openpyxl
    writes. Their ranges follow moved and appended rows, and their totals
    rows and column names (which Excel requires to match the header cells)
    are kept right. Formulas naming a renamed column follow it, as in
    Excel. excelize skips such tables and resets totals rows.
  - chart ranges follow inserted and deleted rows and columns, and rows
    appended to their table
  - renaming a sheet updates formulas, chart series and pivot sources, not
    just names
  - deleting rows or columns turns references into them into `#REF!` (as
    Excel does) instead of pointing them at the neighbouring cells
  - moving rows keeps the namespace declarations of chart drawings, which
    excelize drops
  - notes move with their cells (excelize leaves them behind), comment
    threads from Excel 365 too; notes parts named otherwise than excelize
    expects (openpyxl's) are renamed first, and each drawing keeps one note
    shape per note
  - error values typed into cells (`#N/A`) stay error values when copied,
    moved or sorted
  - conditional formats are never handed to excelize, which moves them
    wrongly when rows go (a range losing its first row slides onto the
    header) and has no fast way to remove one. They are taken out of the
    sheets when a file is opened, kept with the ones the session adds, and
    written back when saving; data validations are taken off around each
    row or column edit. Both move as in Excel: a range keeps its cells, and
    its formulas, written for its first cell, are rewritten for the cell
    that becomes first. Excel 2010 data bars keep their extension.
  - a value written into a cell that merged cells hide is refused, where
    excelize would write it over the merge's first cell
  - dynamic arrays get the cell metadata Excel 365 expects
- Formula results: a file `sheets` saves never stores a stale result.
  When saving, it follows the dependencies from every changed cell, and
  from rows, columns or sheets that moved. It recomputes each formula
  reached with excelize's calculator and stores the result, so pandas,
  previews and other tools read what Excel would show.
  - Table references and LET are computed by translating them for the
    calculator (into cell ranges, and LET's names inlined), in a copy of the
    workbook.
  - SUM, COUNT, COUNTA, AVERAGE, MIN and MAX of plain ranges are computed
    by `sheets` itself, as Excel computes them (excelize's SUM skips an
    error Excel would give), and fast: a column of running totals is summed
    in one pass.
  - A formula the calculator can't evaluate (LAMBDA, FILTER, XMATCH, ...;
    array arithmetic over ranges, such as `SUMPRODUCT((A2:A9="East")*B2:B9)`,
    which it gets wrong; or one holding an error like `#REF!`, which it
    misreads) loses its stored result instead of keeping a wrong one, and
    Excel computes it on open.
  - The calculator recomputes everything a formula reads, silently: so a
    result is trusted only when everything under it is computable, free of
    circular references, and agrees with every result Excel stored on the
    way. A formula the calculator can't evaluate, whose stored result
    still holds, stands at that result for what reads it.
  - excelize gives every result as text, so a small type checker works out
    whether "005" is text (`TEXT(...)`) or a number, and whether "1" is
    TRUE; a result it can't type is left to Excel.
  - Results that only Excel can update, from links to other workbooks, are
    kept.
  - Charts drawing changed cells drop their cached copy of the data. The
    file is still flagged for Excel to recalculate. A fingerprint stamp
    tells `sheets` itself that the stored results are current; an edit by
    another tool voids the stamp. Calls to functions that don't exist show
    `#NAME?`.
- `internal/analyze` finds the structure: blocks of cells, titles, header
  and total rows, column profiles, and formula patterns
  (`=D{r}-E{r}`).

## Limitations

- .xlsx, .xlsm, .csv and .tsv. Legacy .xls and .ods are not read.
- A CSV file is edited through a workbook in memory, so a large one takes
  as long to edit as an .xlsx of the same size.
- `fmt` leaves merged cells alone. Excel tables (Insert › Table) keep their
  own style; `fmt` only sets their number formats and widths.
- Custom row heights from other tools are kept.
- Charts, pivot tables and images are kept. Chart ranges follow row and
  column edits; pivot table sources follow sheet renames only.
- Columns cannot be inserted into or deleted from an Excel table.
- Inserted rows and columns are not formatted like their neighbours, as
  Excel does; `fmt` styles them. Conditional formats and banding reach
  them.
- `render` approximates Excel: other fonts stand in for Calibri, and
  conditional formats other than banding and Excel table styles are not
  drawn.
- `do` re-reads and re-writes the workbook in memory for each command: fine
  for ordinary files, slow for a batch of many commands on a file with
  hundreds of thousands of rows.
- Styling, column edits, sheet edits and range operations load the whole
  workbook into excelize: several seconds and about 2.5 GB for a
  200,000-row sheet. Value and row edits stream (see Large files).

## Development

```bash
go test ./...
go test ./cmd/sheets -update                     # rewrite expected CLI output in cmd/sheets/testdata
STREAM_ROUNDS=500 go test ./internal/xlsx -run TestStreamMatchesFullRandom   # a longer comparison of the two editors
go run ./cmd/sheets skill --install --dir skills  # after changing the guide: regenerate skills/sheets
```

End-to-end CLI tests are [testscript](https://pkg.go.dev/github.com/rogpeppe/go-internal/testscript)
files in `cmd/sheets/testdata`. `internal/xlsx/testdata` holds workbooks
made by openpyxl and XlsxWriter (see `scripts/fixtures.py`), so tests cover
files from other producers.

`scripts/excel-check.sh` builds a set of workbooks and a checklist for a
manual look in real Excel. That's the one check the test suite can't do.

[docs/deferred.md](docs/deferred.md) lists work we chose not to do yet, and
why.
