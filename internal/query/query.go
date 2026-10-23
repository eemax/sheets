// Package query runs SQL over the tables of a workbook, using an embedded
// SQLite (pure Go, so sheets stays a single binary).
package query

import (
	"database/sql"
	"errors"
	"fmt"
	"math"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/grid"
	_ "modernc.org/sqlite"
)

// Table is a block of a sheet exposed to SQL.
type Table struct {
	Name    string   `json:"name"`
	Sheet   string   `json:"sheet"`
	Range   string   `json:"range"` // the whole table, header included
	Rows    int      `json:"rows"`
	Total   bool     `json:"total_row_excluded,omitempty"`
	Columns []Column `json:"columns"`

	sheet  *grid.Sheet
	body   grid.Range
	header int
}

// Body is the table's data rows (header and total row excluded).
func (t *Table) Body() grid.Range { return t.body }

// HeaderRow is the row holding the column labels.
func (t *Table) HeaderRow() int { return t.header }

// Index is the sheet column a table column comes from.
func (c Column) Index() int { return c.col }

// Column is one SQL column of a table.
type Column struct {
	Name   string `json:"name"`
	Type   string `json:"type"` // INTEGER, REAL, TEXT, or "" for mixed values
	Kind   string `json:"kind"` // what the cells hold: integer, currency, date, ...
	Symbol string `json:"symbol,omitempty"`
	col    int
}

// RowColumn holds the sheet row of each record.
const RowColumn = "_row"

// Extra defines a table by range, for blocks sheets does not detect as
// tables. The first row of the range is the header.
type Extra struct {
	Name  string
	Sheet string
	Range grid.Range
}

// Tables lists the tables of a workbook: every detected table, named after
// its Excel table, or its sheet ("Sheet", then "Sheet_2", ...).
func Tables(wb *grid.Workbook, extras []Extra) ([]*Table, error) {
	var out []*Table
	used := map[string]bool{}
	unique := func(name string) string {
		n := name
		for i := 2; used[strings.ToLower(n)]; i++ {
			n = fmt.Sprintf("%s_%d", name, i)
		}
		used[strings.ToLower(n)] = true
		return n
	}
	for _, e := range extras {
		s := wb.Sheet(e.Sheet)
		if s == nil {
			return nil, fmt.Errorf("--table %s: sheet %q not found (sheets: %s)", e.Name, e.Sheet, strings.Join(wb.SheetNames(), ", "))
		}
		if e.Range.Rows() < 1 {
			return nil, fmt.Errorf("--table %s: range %s is empty", e.Name, e.Range)
		}
		reg := analyze.Region{Range: e.Range, HeaderRow: e.Range.R1, Body: grid.Range{R1: e.Range.R1 + 1, C1: e.Range.C1, R2: e.Range.R2, C2: e.Range.C2}}
		if strings.ContainsAny(e.Name, "\"") || e.Name == "" {
			return nil, fmt.Errorf("--table %q: give a plain name", e.Name)
		}
		out = append(out, build(wb, s, reg, unique(e.Name)))
	}
	for _, s := range wb.Sheets {
		n := 0
		for _, reg := range analyze.Regions(wb, s) {
			if !reg.IsTable() {
				continue
			}
			name := s.Name
			if t := s.TableAt(reg.Range); t != nil {
				name = t.Name
			} else if n > 0 {
				name = fmt.Sprintf("%s_%d", s.Name, n+1)
			}
			n++
			out = append(out, build(wb, s, reg, unique(name)))
		}
	}
	return out, nil
}

func build(wb *grid.Workbook, s *grid.Sheet, reg analyze.Region, name string) *Table {
	t := &Table{Name: name, Sheet: s.Name, Range: reg.Range.String(), Rows: max(reg.Body.Rows(), 0), Total: reg.TotalRow > 0, sheet: s, body: reg.Body, header: reg.HeaderRow}
	cols := analyze.Profile(wb, s, reg)
	seen := map[string]bool{strings.ToLower(RowColumn): true}
	for _, c := range cols {
		base := strings.TrimSpace(c.Header)
		if base == "" {
			base = grid.ColName(c.Col)
		}
		label := base
		for i := 2; seen[strings.ToLower(label)]; i++ {
			label = fmt.Sprintf("%s_%d", base, i)
		}
		seen[strings.ToLower(label)] = true
		col := Column{Name: label, Kind: c.Type, Symbol: c.Symbol, col: c.Col}
		switch c.Type {
		case "integer", "bool":
			col.Type = "INTEGER"
		case "decimal", "currency", "percent":
			col.Type = "REAL"
		case "date", "datetime", "time", "text":
			col.Type = "TEXT"
		}
		if col.Kind == "" {
			col.Kind = "empty"
		}
		t.Columns = append(t.Columns, col)
	}
	return t
}

// Quote quotes an SQL identifier when it needs it.
func Quote(name string) string {
	plain := name != "" && !(name[0] >= '0' && name[0] <= '9')
	for _, r := range name {
		if !(r == '_' || r >= 'a' && r <= 'z' || r >= 'A' && r <= 'Z' || r >= '0' && r <= '9') {
			plain = false
			break
		}
	}
	if plain && !reserved[strings.ToUpper(name)] {
		return name
	}
	return `"` + strings.ReplaceAll(name, `"`, `""`) + `"`
}

var reserved = map[string]bool{}

func init() {
	for _, w := range strings.Fields(`ADD ALL ALTER AND AS ASC BETWEEN BY CASE CHECK COLLATE COLUMN CREATE CROSS
		CURRENT_DATE CURRENT_TIME CURRENT_TIMESTAMP DEFAULT DELETE DESC DISTINCT DROP ELSE END ESCAPE EXCEPT
		EXISTS FROM FULL GLOB GROUP HAVING IN INDEX INNER INSERT INTERSECT INTO IS ISNULL JOIN KEY LEFT LIKE
		LIMIT MATCH NATURAL NOT NOTNULL NULL OFFSET ON OR ORDER OUTER PRIMARY REFERENCES REGEXP RIGHT ROW
		ROWS SELECT SET TABLE THEN TO UNION UNIQUE UPDATE USING VALUES WHEN WHERE WINDOW WITH`) {
		reserved[w] = true
	}
}

// Result holds the rows a query returned.
type Result struct {
	Columns []string `json:"columns"`
	Rows    [][]any  `json:"rows"`
	More    int      `json:"more,omitempty"` // rows beyond the limit
}

// mentioned reports whether a query may refer to a table: its name appears
// in the text. Only those tables are loaded.
func mentioned(q, name string) bool {
	return strings.Contains(strings.ToLower(q), strings.ToLower(name))
}

// Run executes one SQL statement against the tables. At most limit rows are
// returned (all when limit <= 0); More counts the rest.
func Run(wb *grid.Workbook, tables []*Table, q string, limit int) (*Result, error) {
	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		return nil, err
	}
	defer db.Close()
	db.SetMaxOpenConns(1) // one connection: :memory: databases are per connection
	for _, t := range tables {
		if mentioned(q, t.Name) {
			if err := load(db, wb, t); err != nil {
				return nil, fmt.Errorf("loading %s: %w", t.Name, err)
			}
		}
	}
	rows, err := db.Query(q)
	if err != nil {
		return nil, explain(err, q, tables)
	}
	defer rows.Close()
	cols, err := rows.Columns()
	if err != nil {
		return nil, err
	}
	res := &Result{Columns: cols, Rows: [][]any{}}
	for rows.Next() {
		if limit > 0 && len(res.Rows) >= limit {
			res.More++
			continue
		}
		vals := make([]any, len(cols))
		ptrs := make([]any, len(cols))
		for i := range vals {
			ptrs[i] = &vals[i]
		}
		if err := rows.Scan(ptrs...); err != nil {
			return nil, err
		}
		for i, v := range vals {
			if b, ok := v.([]byte); ok {
				vals[i] = string(b)
			}
		}
		res.Rows = append(res.Rows, vals)
	}
	if err := rows.Err(); err != nil {
		return nil, explain(err, q, tables)
	}
	return res, nil
}

func load(db *sql.DB, wb *grid.Workbook, t *Table) error {
	defs := []string{Quote(RowColumn) + " INTEGER"}
	marks := []string{"?"}
	for _, c := range t.Columns {
		defs = append(defs, strings.TrimSpace(Quote(c.Name)+" "+c.Type))
		marks = append(marks, "?")
	}
	if _, err := db.Exec(fmt.Sprintf("CREATE TABLE %s (%s)", Quote(t.Name), strings.Join(defs, ", "))); err != nil {
		return err
	}
	tx, err := db.Begin()
	if err != nil {
		return err
	}
	defer tx.Rollback()
	stmt, err := tx.Prepare(fmt.Sprintf("INSERT INTO %s VALUES (%s)", Quote(t.Name), strings.Join(marks, ", ")))
	if err != nil {
		return err
	}
	defer stmt.Close()
	args := make([]any, len(t.Columns)+1)
	for r := t.body.R1; r <= t.body.R2; r++ {
		args[0] = int64(r)
		for i, c := range t.Columns {
			args[i+1] = value(wb, t.sheet, r, c)
		}
		if _, err := stmt.Exec(args...); err != nil {
			return err
		}
	}
	return tx.Commit()
}

// value converts a cell for SQL: dates as ISO text (SQLite's date functions
// read it), percentages as fractions, errors and blanks as NULL.
func value(wb *grid.Workbook, s *grid.Sheet, row int, c Column) any {
	cell := wb.Value(s, row, c.col)
	if cell == nil || cell.Pending {
		return nil
	}
	switch cell.Kind {
	case grid.Number:
		f := wb.FormatOf(cell)
		if f.Kind.IsTemporal() {
			return grid.FormatTemporal(cell.Num, f.Kind, wb.Date1904)
		}
		if c.Type == "INTEGER" && cell.Num == math.Trunc(cell.Num) && math.Abs(cell.Num) < 1<<53 {
			return int64(cell.Num)
		}
		return cell.Num
	case grid.Text:
		return cell.Text
	case grid.Bool:
		if cell.Bool {
			return int64(1)
		}
		return int64(0)
	}
	return nil
}

var sqliteNoise = regexp.MustCompile(`^SQL logic error: | \(\d+\)$`)

// explain adds the available tables or columns to SQL errors about them.
func explain(err error, q string, tables []*Table) error {
	msg := sqliteNoise.ReplaceAllString(err.Error(), "")
	switch {
	case strings.Contains(msg, "no such table"):
		var names []string
		for _, t := range tables {
			names = append(names, Quote(t.Name))
		}
		if len(names) == 0 {
			return fmt.Errorf("%s; the workbook has no tables (blocks with a header row); define one with --table name=Sheet!A1:D20", msg)
		}
		return fmt.Errorf("%s (tables: %s)", msg, strings.Join(names, ", "))
	case strings.Contains(msg, "no such column"):
		var parts []string
		for _, t := range tables {
			if !mentioned(q, t.Name) {
				continue
			}
			var cols []string
			for _, c := range t.Columns {
				cols = append(cols, Quote(c.Name))
			}
			parts = append(parts, Quote(t.Name)+"("+strings.Join(cols, ", ")+")")
		}
		sort.Strings(parts)
		return fmt.Errorf("%s (columns: %s)", msg, strings.Join(parts, "; "))
	}
	return errors.New(msg)
}
