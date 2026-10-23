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
| `write <file> [sheet!cell]` | Writes TSV/CSV/JSON from stdin with type inference; `--append` (matches headers), `--clear` |
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

### Undo

Every change keeps the previous version of the file (20 per file) in the
user cache directory (`SHEETS_CACHE_DIR` overrides it; `SHEETS_UNDO=off`
disables it). `sheets undo` restores the last one. It refuses if the file
was changed by something else since.

## The house style

`sheets fmt` throws away existing cell formatting and rebuilds it from the
contents:

- a dark header row with filter buttons, frozen in place
- alternate data rows banded light gray, using conditional formatting so
  the banding survives sorting
- thin light borders, gridlines off
- a bold total row with a heavier rule above it
- titles above a table in bold 14 pt
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
    writes
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
- Charts, pivot tables and images are kept, but chart data ranges do not
  move when rows or columns are inserted or deleted (renaming a sheet does
  update them).
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
