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
| `set <file> <cell> <value>...` | Sets cells or formulas; a range fills down with shifted references |
| `rows insert\|delete`, `cols insert\|delete` | Insert or delete rows and columns; references follow, and broken ones become `#REF!` and are reported |
| `sheet add\|rename\|delete\|copy\|move\|hide\|show` | Manage sheets; renaming updates every reference, deleting refuses while it is referenced |
| `fmt <file>` | Strips formatting and applies the house style; `--preset finance`, `--as COL=type`, `--check` |
| `undo <file>` | Restores the file to before the last change sheets made; `--list` |
| `agent-docs` | Prints the guide for agents (also in [internal/cli/agent-docs.md](internal/cli/agent-docs.md)) |
| `skill` | Prints or installs (`--install`) the Claude Code skill |

Every command takes `--json`, and every command that changes a file takes
`--dry-run`.

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

Formatting depends only on the cell contents, so `fmt` is idempotent and
`fmt --check` works in CI. Column types are inferred, and `--as Revenue=currency`
overrides them. The choice is saved in the file, so later runs keep it.

Formula columns take the type of what they compute from: `=Units*Price` is
currency when Price is, and `=Start+30` is a date. Divisions and date
differences stay plain numbers.

## Design

- `internal/grid` is a model of a workbook that doesn't depend on any file
  format. Analysis and output only see this model. `internal/book` picks the
  implementation for a file's extension, so adding CSV or ODS means adding
  a reader and editor there.
- `internal/xlsx` reads .xlsx with its own streaming XML parser. It is fast
  (`peek` on 20k rows with 20k formulas takes about half a second) and
  sees shared formulas and cached values as stored. Editing and styling go
  through
  [excelize](https://github.com/qax-os/excelize), with workarounds for
  places where excelize would corrupt or lose data:
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
  - dynamic arrays get the cell metadata Excel 365 expects
- Formula values: after `sheets` writes a file, it marks the file for
  recalculation, and its cached results are treated as stale. `sheets`
  computes them with excelize's calculator and Excel recalculates on open.
  Calls to functions that don't exist show `#NAME?`. Real Excel functions
  the calculator lacks (LET, LAMBDA, FILTER, ...) show `?`: Excel computes
  them.
- `internal/analyze` finds the structure: blocks of cells, titles, header
  and total rows, column profiles, and formula patterns
  (`=D{r}-E{r}`).

## Limitations

- .xlsx and .xlsm only, for now.
- `fmt` leaves merged cells alone. Excel tables (Insert › Table) keep their
  own style; `fmt` only sets their number formats and widths.
- Custom row heights from other tools are kept.
- Charts, pivot tables and images are kept. Chart ranges follow row and
  column edits; pivot table sources follow sheet renames only.
- Columns cannot be inserted into or deleted from an Excel table.
- `render` approximates Excel: other fonts stand in for Calibri, and
  conditional formats other than banding and Excel table styles are not
  drawn.
- Deleting rows goes through excelize one row at a time. A few rows take
  about a second, but a thousand rows from a 20,000-row sheet take about 15
  seconds.

## Development

```bash
go test ./...
go test ./cmd/sheets -update   # rewrite expected CLI output in cmd/sheets/testdata
```

End-to-end CLI tests are [testscript](https://pkg.go.dev/github.com/rogpeppe/go-internal/testscript)
files in `cmd/sheets/testdata`. `internal/xlsx/testdata` holds workbooks
made by openpyxl (see `scripts/fixtures.py`), so tests cover files from
another producer.

`scripts/excel-check.sh` builds a set of workbooks and a checklist for a
manual look in real Excel. That's the one check the test suite can't do.
