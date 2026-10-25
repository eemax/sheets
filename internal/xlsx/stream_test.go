package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"math"
	"math/rand"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/xuri/excelize/v2"
)

// The streaming editor must leave a file as the full editor does. These
// tests make the same edits through both and compare what a reader sees.

type editOp struct {
	kind   string // set, insert, delete, extend
	sheet  string
	row    int
	col    int
	n      int
	value  string
	body   grid.Range
	stream bool // the streaming editor must do it without the full editor
}

func (o editOp) String() string {
	switch o.kind {
	case "set":
		return fmt.Sprintf("set %s!%s %q", o.sheet, grid.CellName(o.row, o.col), o.value)
	case "insert":
		return fmt.Sprintf("insert %s!%d -n %d", o.sheet, o.row, o.n)
	case "delete":
		return fmt.Sprintf("delete %s!%d:%d", o.sheet, o.row, o.row+o.n-1)
	}
	return fmt.Sprintf("extend %s %s by %d", o.sheet, o.body, o.n)
}

type editor interface {
	Set(sheet string, row, col int, v input.Value) error
	InsertRows(sheet string, at, n int) error
	DeleteRows(sheet string, from, to int) error
	ExtendTable(sheet string, body grid.Range, n int)
	Save(dryRun bool) error
}

func apply(ed editor, ops []editOp) []string {
	var errs []string
	for _, o := range ops {
		var err error
		switch o.kind {
		case "set":
			err = ed.Set(o.sheet, o.row, o.col, input.Parse(o.value, false))
		case "insert":
			err = ed.InsertRows(o.sheet, o.row, o.n)
		case "delete":
			err = ed.DeleteRows(o.sheet, o.row, o.row+o.n-1)
		case "extend":
			ed.ExtendTable(o.sheet, o.body, o.n)
		}
		if err != nil {
			errs = append(errs, o.String()+": "+err.Error())
		}
	}
	return errs
}

// both runs ops through the full editor and the streaming one, on copies of
// path, and returns the two files.
func both(t *testing.T, path string, ops []editOp) (full, stream string, upgraded bool) {
	t.Helper()
	data, err := os.ReadFile(path)
	must(t, err)
	dir := t.TempDir()
	full, stream = filepath.Join(dir, "full.xlsx"), filepath.Join(dir, "stream.xlsx")
	must(t, os.WriteFile(full, data, 0o644))
	must(t, os.WriteFile(stream, data, 0o644))

	e, err := OpenEditor(full, false)
	must(t, err)
	fullErrs := apply(e, ops)
	must(t, e.Save(false))
	e.Close()

	a, err := OpenAuto(stream, data)
	must(t, err)
	streamErrs := apply(a, ops)
	upgraded = !a.Streaming()
	must(t, a.Save(false))
	a.Close()
	if strings.Join(fullErrs, "\n") != strings.Join(streamErrs, "\n") {
		t.Fatalf("errors differ:\nfull:   %q\nstream: %q", fullErrs, streamErrs)
	}
	return full, stream, upgraded
}

// sameFile compares what a reader sees in two workbooks.
func sameFile(t *testing.T, full, stream string, label string) {
	t.Helper()
	validPackage(t, stream)
	a, b := read(t, full), read(t, stream)
	if strings.Join(a.SheetNames(), ",") != strings.Join(b.SheetNames(), ",") {
		t.Fatalf("%s: sheets %v vs %v", label, a.SheetNames(), b.SheetNames())
	}
	diffs := 0
	report := func(format string, args ...any) {
		diffs++
		if diffs <= 12 {
			t.Errorf("%s: %s", label, fmt.Sprintf(format, args...))
		}
	}
	for _, sa := range a.Sheets {
		sb := b.Sheet(sa.Name)
		ua, ub := sa.Used(), sb.Used()
		if ua != ub {
			report("%s used %s vs %s", sa.Name, ua, ub)
		}
		u := ua.Union(ub)
		for r := u.R1; r <= u.R2 && !u.Empty(); r++ {
			for c := u.C1; c <= u.C2; c++ {
				x, y := sa.Get(r, c), sb.Get(r, c)
				ref := sa.Name + "!" + grid.CellName(r, c)
				if x.Formula != y.Formula {
					report("%s formula %q vs %q", ref, x.Formula, y.Formula)
					continue
				}
				// Whether the file stores a result, before reading
				// computes one.
				if x.Pending != y.Pending {
					report("%s stored a result %v vs %v", ref, !x.Pending, !y.Pending)
				}
				vx, vy := a.Value(sa, r, c), b.Value(sb, r, c)
				if vx == nil || vy == nil {
					if (vx == nil) != (vy == nil) {
						report("%s %v vs %v", ref, vx, vy)
					}
					continue
				}
				if vx.Pending != vy.Pending {
					report("%s pending %v vs %v", ref, vx.Pending, vy.Pending)
					continue
				}
				if vx.Kind != vy.Kind || vx.Text != vy.Text || vx.Bool != vy.Bool ||
					vx.Kind == grid.Number && math.Abs(vx.Num-vy.Num) > 1e-9*math.Max(1, math.Abs(vx.Num)) {
					report("%s value %s vs %s", ref, show(*vx), show(*vy))
				}
				fx, fy := a.FormatOf(&x), b.FormatOf(&y)
				if fx.Kind != fy.Kind || fx.Decimals != fy.Decimals || fx.Symbol != fy.Symbol {
					report("%s format %+v vs %+v", ref, fx, fy)
				}
			}
		}
		check := func(what string, x, y any) {
			if fmt.Sprint(x) != fmt.Sprint(y) {
				report("%s %s %v vs %v", sa.Name, what, x, y)
			}
		}
		check("merges", sortedRanges(sa.Merges), sortedRanges(sb.Merges))
		check("notes", sa.Notes, sb.Notes)
		check("links", sa.Links, sb.Links)
		check("tables", sa.Tables, sb.Tables)
		check("hidden rows", sa.HiddenRows, sb.HiddenRows)
		check("filter", sa.Filter, sb.Filter)
		var ca, cb []string
		for _, ch := range sa.Charts {
			ca = append(ca, fmt.Sprint(ch.Shown().Series))
		}
		for _, ch := range sb.Charts {
			cb = append(cb, fmt.Sprint(ch.Shown().Series))
		}
		check("charts", ca, cb)
	}
	names := func(wb *grid.Workbook) []string {
		var out []string
		for _, n := range wb.Names {
			out = append(out, n.Name+"="+n.Ref)
		}
		sort.Strings(out)
		return out
	}
	if fmt.Sprint(names(a)) != fmt.Sprint(names(b)) {
		report("names %v vs %v", names(a), names(b))
	}
	// What the full editor reads of formatting: records, conditional
	// formats, validations, panes.
	ea, err := OpenEditor(full, false)
	must(t, err)
	defer ea.Close()
	eb, err := OpenEditor(stream, false)
	must(t, err)
	defer eb.Close()
	var ra, rb []string
	for _, r := range ea.Records() {
		ra = append(ra, r.Ref+" "+r.Text())
	}
	for _, r := range eb.Records() {
		rb = append(rb, r.Ref+" "+r.Text())
	}
	if fmt.Sprint(ra) != fmt.Sprint(rb) {
		report("records %v vs %v", ra, rb)
	}
	for _, sa := range a.Sheets {
		cfa, _ := ea.CondFormats(sa.Name)
		cfb, _ := eb.CondFormats(sa.Name)
		if fmt.Sprint(cfa) != fmt.Sprint(cfb) {
			report("%s conditional formats %v vs %v", sa.Name, cfa, cfb)
		}
		dva, _ := ea.Validations(sa.Name)
		dvb, _ := eb.Validations(sa.Name)
		// A2:A2 and A2 are the same range.
		for _, dvs := range [][]Validation{dva, dvb} {
			for i := range dvs {
				dvs[i].Range = strings.Join(sortedRanges(sqrefRanges(dvs[i].Range)), " ")
			}
		}
		if fmt.Sprint(dva) != fmt.Sprint(dvb) {
			report("%s validations %v vs %v", sa.Name, dva, dvb)
		}
		va, _ := ea.ViewOf(sa.Name)
		vb, _ := eb.ViewOf(sa.Name)
		if va.Freeze != vb.Freeze {
			report("%s frozen at %q vs %q", sa.Name, va.Freeze, vb.Freeze)
		}
		for r := 1; r <= 3; r++ {
			for c := 1; c <= 4; c++ {
				la, _ := ea.Looks(sa.Name, grid.Range{R1: r, C1: c, R2: r, C2: c})
				lb, _ := eb.Looks(sa.Name, grid.Range{R1: r, C1: c, R2: r, C2: c})
				if fmt.Sprint(la) != fmt.Sprint(lb) {
					report("%s!%s look %v vs %v", sa.Name, grid.CellName(r, c), la, lb)
				}
			}
		}
	}
	if diffs > 12 {
		t.Errorf("%s: … %d differences in all", label, diffs)
	}
}

func sortedRanges(rs []grid.Range) []string {
	var out []string
	for _, r := range rs {
		out = append(out, r.String())
	}
	sort.Strings(out)
	return out
}

// styledBook is a workbook with what row edits must carry along:
// formulas, totals, recorded styling, conditional formats, validation,
// merges, notes, links, a frozen header and a second sheet reading the
// first.
func styledBook(t *testing.T) string {
	path := writeBook(t, [][]string{
		{"Region", "Units", "Price", "Revenue", "Share"},
		{"East", "5", "30", "=B{r}*C{r}", "=D{r}/D$9"},
		{"West", "2", "40", "=B{r}*C{r}", "=D{r}/D$9"},
		{"North", "9", "10", "=B{r}*C{r}", "=D{r}/D$9"},
		{"South", "1", "100", "=B{r}*C{r}", "=D{r}/D$9"},
		{"East", "3", "30", "=B{r}*C{r}", "=D{r}/D$9"},
		{"West", "7", "40", "=B{r}*C{r}", "=D{r}/D$9"},
		{"North", "4", "10", "=B{r}*C{r}", "=D{r}/D$9"},
		{"Total", "=SUM(B2:B8)", "", "=SUM(D2:D8)", ""},
	})
	format(t, path)
	change(t, path, func(e *Editor) {
		_, err := e.Style("Data", rng("D2:D8"), look(t, "format=currency decimals=0"))
		must(t, err)
		_, err = e.Style("Data", rng("A1:E1"), look(t, "fill=lightblue bold"))
		must(t, err)
		_, err = e.Cond("Data", rng("A2:E8"), []string{`=$B2>5`, "fill=lightyellow"})
		must(t, err)
		_, err = e.Validate("Data", rng("B2:B8"), ValidationRule{Words: []string{"whole", "between", "0", "and", "1000"}})
		must(t, err)
		must(t, e.SetNote("Data", 3, 1, "West office", "sheets"))
		must(t, e.SetLink("Data", 4, 1, &grid.Link{URL: "https://example.com/north"}, ""))
		must(t, e.View("Data", ViewOptions{Freeze: &grid.Range{R1: 2, C1: 1, R2: 2, C2: 1}}))
		_, _, err = e.EnsureSheet("Summary")
		must(t, err)
		for cell, f := range map[[2]int]string{{1, 1}: "Units", {1, 2}: "=Data!B9", {2, 1}: "Top", {2, 2}: "=MAX(Data!D2:D8)", {3, 1}: "Fourth", {3, 2}: "=Data!D5"} {
			must(t, e.Set("Summary", cell[0], cell[1], input.Parse(f, false)))
		}
	})
	change(t, path, func(e *Editor) { must(t, e.f.MergeCell("Summary", "D1", "E1")) })
	return path
}

func TestStreamMatchesFull(t *testing.T) {
	base := styledBook(t)
	cases := map[string][]editOp{
		"one value": {{kind: "set", sheet: "Data", row: 3, col: 2, value: "20", stream: true}},
		"values of each kind": {
			{kind: "set", sheet: "Data", row: 2, col: 3, value: "$45.50", stream: true},
			{kind: "set", sheet: "Data", row: 4, col: 1, value: "Northeast", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 1, value: "2026-03-01", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 2, value: "12%", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 3, value: "TRUE", stream: true},
			{kind: "set", sheet: "Data", row: 11, col: 6, value: "_x0041_ & <tag>", stream: true},
			{kind: "set", sheet: "Data", row: 5, col: 2, value: "", stream: true},
		},
		"formulas": {
			{kind: "set", sheet: "Data", row: 10, col: 4, value: "=SUM(D2:D8)/COUNT(D2:D8)", stream: true},
			{kind: "set", sheet: "Summary", row: 4, col: 2, value: "=AVERAGE(Data!B2:B8)", stream: true},
			{kind: "set", sheet: "Data", row: 2, col: 4, value: "=B2*C2*2", stream: true},
		},
		"rows inserted": {
			{kind: "insert", sheet: "Data", row: 5, n: 2, stream: true},
			{kind: "set", sheet: "Data", row: 5, col: 1, value: "Central", stream: true},
			{kind: "set", sheet: "Data", row: 5, col: 2, value: "6", stream: true},
		},
		"rows deleted": {
			{kind: "delete", sheet: "Data", row: 3, n: 2, stream: true},
			{kind: "set", sheet: "Data", row: 2, col: 2, value: "11", stream: true},
		},
		"rows above everything": {
			{kind: "insert", sheet: "Data", row: 1, n: 3, stream: true},
			{kind: "set", sheet: "Data", row: 1, col: 1, value: "Title", stream: true},
		},
		"the header deleted": {
			{kind: "delete", sheet: "Data", row: 1, n: 1, stream: true},
		},
		"the total deleted": {
			{kind: "delete", sheet: "Data", row: 9, n: 1, stream: true},
		},
		"rows appended": {
			{kind: "insert", sheet: "Data", row: 9, n: 2, stream: true},
			{kind: "set", sheet: "Data", row: 9, col: 1, value: "Central", stream: true},
			{kind: "set", sheet: "Data", row: 9, col: 2, value: "8", stream: true},
			{kind: "set", sheet: "Data", row: 9, col: 3, value: "20", stream: true},
			{kind: "set", sheet: "Data", row: 9, col: 4, value: "=B9*C9", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 1, value: "Island", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 2, value: "1", stream: true},
			{kind: "set", sheet: "Data", row: 10, col: 4, value: "=B10*C10", stream: true},
			{kind: "extend", sheet: "Data", body: rng("A2:E8"), n: 2, stream: true},
			{kind: "set", sheet: "Data", row: 11, col: 2, value: "=SUM(B2:B10)", stream: true},
			{kind: "set", sheet: "Data", row: 11, col: 4, value: "=SUM(D2:D10)", stream: true},
		},
		"a sheet added": {
			{kind: "set", sheet: "Data", row: 2, col: 2, value: "50", stream: true},
		},
	}
	names := make([]string, 0, len(cases))
	for n := range cases {
		names = append(names, n)
	}
	sort.Strings(names)
	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			full, stream, upgraded := both(t, base, cases[name])
			if upgraded {
				t.Errorf("the streaming editor handed over to the full editor")
			}
			sameFile(t, full, stream, name)
		})
	}
}

// Random edits on workbooks from other programs and from sheets.
func TestStreamMatchesFullRandom(t *testing.T) {
	books := map[string]string{
		"styled":  styledBook(t),
		"rich":    fixture(t, "rich.xlsx"),
		"tables":  fixture(t, "tables.xlsx"),
		"cached":  fixture(t, "cached.xlsx"),
		"threads": fixture(t, "threads.xlsx"),
		"modern":  fixture(t, "styled.xlsx"),
		"rules":   fixture(t, "rules.xlsx"),
	}
	rounds := 6
	if testing.Short() {
		rounds = 2
	}
	if n, err := strconv.Atoi(os.Getenv("STREAM_ROUNDS")); err == nil {
		rounds = n
	}
	var keys []string
	for k := range books {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	values := []string{"7", "-3.5", "Text", "12%", "$9.99", "2026-05-17", "", "=1+1", "FALSE", "1,234"}
	for _, name := range keys {
		wb := read(t, books[name])
		for round := 0; round < rounds; round++ {
			rnd := rand.New(rand.NewSource(int64(round*7919 + len(name))))
			var ops []editOp
			for i := 0; i < 6; i++ {
				s := wb.Sheets[rnd.Intn(len(wb.Sheets))]
				u := s.Used()
				if u.Empty() {
					u = grid.Range{R1: 1, C1: 1, R2: 3, C2: 3}
				}
				row := u.R1 + rnd.Intn(u.Rows()+2)
				col := u.C1 + rnd.Intn(u.Cols()+1)
				switch rnd.Intn(5) {
				case 0:
					ops = append(ops, editOp{kind: "insert", sheet: s.Name, row: row, n: 1 + rnd.Intn(3)})
				case 1:
					ops = append(ops, editOp{kind: "delete", sheet: s.Name, row: row, n: 1 + rnd.Intn(2)})
				default:
					v := values[rnd.Intn(len(values))]
					if rnd.Intn(4) == 0 && u.Rows() > 1 {
						// A formula over a column of the sheet.
						v = fmt.Sprintf("=SUM(%s%d:%s%d)", grid.ColName(u.C1), u.R1, grid.ColName(u.C1), u.R2)
					}
					ops = append(ops, editOp{kind: "set", sheet: s.Name, row: row, col: col, value: v})
				}
			}
			label := fmt.Sprintf("%s/%d %v", name, round, ops)
			t.Run(fmt.Sprintf("%s-%d", name, round), func(t *testing.T) {
				full, stream, _ := both(t, books[name], ops)
				sameFile(t, full, stream, label)
			})
		}
	}
}

// Sequences the random comparison found, kept as they are.
func TestStreamMatchesFullFound(t *testing.T) {
	cases := []struct {
		book string
		ops  []editOp
	}{
		// Rows inserted at the top of a rule's range, then its first row
		// deleted: the rule's formulas follow the inserted rows.
		{"rules.xlsx", []editOp{{kind: "insert", sheet: "Data", row: 4, n: 3}, {kind: "set", sheet: "Data", row: 2, col: 2, value: "=SUM(A1:A7)"},
			{kind: "delete", sheet: "Data", row: 2, n: 2}}},
		// The used range losing its first rows: a sum over them.
		{"threads.xlsx", []editOp{{kind: "insert", sheet: "Data", row: 1, n: 1}, {kind: "set", sheet: "Data", row: 6, col: 1, value: "2026-05-17"},
			{kind: "delete", sheet: "Data", row: 2, n: 2}, {kind: "set", sheet: "Data", row: 3, col: 1, value: "=1+1"},
			{kind: "delete", sheet: "Data", row: 1, n: 2}, {kind: "set", sheet: "Data", row: 5, col: 3, value: "=SUM(A1:A4)"}}},
		// Every formula deleted, then more rows edited.
		{"tables.xlsx", []editOp{{kind: "delete", sheet: "Sales", row: 2, n: 1}, {kind: "delete", sheet: "Sales", row: 3, n: 2},
			{kind: "set", sheet: "Sales", row: 4, col: 6, value: "=SUM(A1:A8)"}, {kind: "delete", sheet: "Sales", row: 6, n: 1},
			{kind: "insert", sheet: "Sales", row: 6, n: 1}, {kind: "insert", sheet: "Sales", row: 3, n: 3}}},
	}
	for _, c := range cases {
		t.Run(c.book, func(t *testing.T) {
			full, stream, _ := both(t, fixture(t, c.book), c.ops)
			sameFile(t, full, stream, fmt.Sprint(c.ops))
		})
	}
}

// Results taken as stored are errors where they were errors, not blanks,
// for every one, not only those computed again.
func TestStreamFrozenErrors(t *testing.T) {
	rows := [][]string{{"k", "v"}}
	for i := 0; i < 40; i++ {
		rows = append(rows, []string{"x", "=NA()"})
	}
	path := writeBook(t, rows)
	change(t, path, func(e *Editor) {
		must(t, e.Set("Data", 1, 4, input.Parse("=COUNTBLANK(B2:B41)+C1", false)))
	})
	full, stream, upgraded := both(t, path, []editOp{{kind: "set", sheet: "Data", row: 1, col: 3, value: "1", stream: true}})
	if upgraded {
		t.Fatal("the edit left the streaming editor")
	}
	for _, p := range []string{full, stream} {
		if c := read(t, p).Sheet("Data").Get(1, 4); c.Kind != grid.Number || c.Num != 1 {
			t.Errorf("%s: D1 = %+v, want 1 (errors are not blank)", filepath.Base(p), c)
		}
	}
}

// A recorded freeze moves as its pane does, and view entries that are not
// cells (gridlines) stay through row edits: fmt lays them again.
func TestViewRecordsFollowPanes(t *testing.T) {
	cases := []struct {
		ops  []editOp
		want string // the freeze after the edits, and after fmt
	}{
		{[]editOp{{kind: "delete", sheet: "Data", row: 2, n: 1}}, "B2"},
		{[]editOp{{kind: "insert", sheet: "Data", row: 2, n: 2}}, "B2"},
		{[]editOp{{kind: "insert", sheet: "Data", row: 1, n: 2}}, "B4"},
		{[]editOp{{kind: "delete", sheet: "Data", row: 1, n: 1}}, "B1"},
	}
	for _, c := range cases {
		path := writeBook(t, [][]string{{"Region", "Units"}, {"East", "5"}, {"West", "2"}, {"North", "9"}})
		change(t, path, func(e *Editor) {
			must(t, e.View("Data", ViewOptions{Freeze: &grid.Range{R1: 2, C1: 2, R2: 2, C2: 2}, Gridlines: ptr(false)}))
		})
		full, stream, _ := both(t, path, c.ops)
		sameFile(t, full, stream, fmt.Sprint(c.ops))
		for _, p := range []string{full, stream} {
			format(t, p)
			e, err := OpenEditor(p, false)
			must(t, err)
			v, err := e.ViewOf("Data")
			must(t, err)
			e.Close()
			if v.Freeze != c.want || v.Gridlines {
				t.Errorf("%v %s: frozen at %q, gridlines %v; want %s, off", c.ops, filepath.Base(p), v.Freeze, v.Gridlines, c.want)
			}
		}
	}
}

// Built-in names follow as in the full editor: a print area that loses
// its cells goes, and a filter's name goes with its header row.
func TestBuiltInNames(t *testing.T) {
	for _, ops := range [][]editOp{
		{{kind: "delete", sheet: "Data", row: 2, n: 2}},
		{{kind: "delete", sheet: "Data", row: 1, n: 1}},
		{{kind: "insert", sheet: "Data", row: 1, n: 1}},
	} {
		path := writeBook(t, [][]string{{"Region", "Units"}, {"East", "5"}, {"West", "2"}, {"North", "9"}})
		format(t, path) // a filter, with its hidden name
		change(t, path, func(e *Editor) {
			must(t, e.f.SetDefinedName(&excelize.DefinedName{Name: "_xlnm.Print_Area", RefersTo: "Data!$A$2:$B$3", Scope: "Data"}))
		})
		full, stream, _ := both(t, path, ops)
		names := func(p string) string {
			x := sheetXML(t, p, "xl/workbook.xml")
			var out []string
			for _, m := range regexp.MustCompile(`<definedName\b[^>]*\bname="(_xlnm[^"]*)"[^>]*>([^<]*)<`).FindAllStringSubmatch(x, -1) {
				out = append(out, m[1]+"="+m[2])
			}
			sort.Strings(out)
			return strings.Join(out, " ")
		}
		if a, b := names(full), names(stream); a != b {
			t.Errorf("%v: names %q (full) vs %q (stream)", ops, a, b)
		}
	}
}

func TestRowSpans(t *testing.T) {
	var sp rowSpans
	for _, c := range []struct {
		r1, r2  int
		missing string
		after   string
	}{
		{5, 9, "[[5 9]]", "[[5 9]]"},
		{1, 3, "[[1 3]]", "[[1 3] [5 9]]"},
		{2, 12, "[[4 4] [10 12]]", "[[1 12]]"},
		{20, 20, "[[20 20]]", "[[1 12] [20 20]]"},
		{13, 19, "[[13 19]]", "[[1 20]]"},
		{3, 8, "[]", "[[1 20]]"},
	} {
		if got := fmt.Sprint(sp.missing(c.r1, c.r2)); got != c.missing {
			t.Errorf("missing(%d, %d) = %s, want %s", c.r1, c.r2, got, c.missing)
		}
		sp.add(c.r1, c.r2)
		if got := fmt.Sprint(sp.spans); got != c.after {
			t.Errorf("after add(%d, %d): %s, want %s", c.r1, c.r2, got, c.after)
		}
	}
}

// Formulas reading a range that holds an error compute with it, in the
// streaming editor as in the full one.
func TestStreamErrorInRange(t *testing.T) {
	path := writeBook(t, [][]string{{"v"}, {"1"}, {"x"}, {"3"}})
	change(t, path, func(e *Editor) {
		must(t, e.f.SetCellDefault("Data", "A3", "#REF!"))
		e.patches.errorCells["Data"] = map[string]bool{"A3": true}
		must(t, e.Set("Data", 1, 3, input.Parse("=COUNTBLANK(A1:A4)+B1", false)))
		must(t, e.Set("Data", 2, 3, input.Parse("=COUNTBLANK(A1:A4)+B1*2", false)))
	})
	if c := read(t, path).Sheet("Data").Get(3, 1); c.Kind != grid.Error {
		t.Fatalf("A3 = %+v, want the error #REF!", c)
	}
	data, err := os.ReadFile(path)
	must(t, err)
	a, err := OpenAuto(path, data)
	must(t, err)
	must(t, a.Set("Data", 1, 2, input.Parse("1", false)))
	must(t, a.Save(false))
	if !a.Streaming() {
		t.Fatal("left the streaming editor")
	}
	// The engine holds the error like any value: COUNTBLANK reads it.
	s := read(t, path).Sheet("Data")
	for r, want := range map[int]float64{1: 1, 2: 2} {
		if c := s.Get(r, 3); c.Pending || c.Num != want {
			t.Errorf("C%d = %+v, want %v", r, c, want)
		}
	}
}

// failingReader gives data, then fails.
type failingReader struct{ data []byte }

func (f *failingReader) Read(p []byte) (int, error) {
	if len(f.data) == 0 {
		return 0, errors.New("flate: corrupt input")
	}
	n := copy(p, f.data)
	f.data = f.data[n:]
	return n, nil
}

// A sheet that cannot be read to its end is not written cut short.
func TestRewriteSheetReadError(t *testing.T) {
	part := `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`
	for _, cut := range []int{len(part) / 3, len(part) - 20} {
		var out bytes.Buffer
		err := rewriteSheet(&failingReader{[]byte(part[:cut])}, &out, &sheetJob{dimension: &grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}})
		if err == nil || !strings.Contains(err.Error(), "corrupt") {
			t.Errorf("cut at %d: err = %v", cut, err)
		}
	}
}
