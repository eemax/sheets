package xlsx

import (
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"sort"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// TestEngineCorpus measures the sheets formula engine against Excel, one
// formula at a time: every other formula cell reads as the result Excel
// stored, so a difference is the formula's own. Off unless
// SHEETS_CALC_CORPUS names a folder of Excel-saved workbooks;
// SHEETS_CALC_CORPUS_SHOW sets the examples listed per function, and
// SHEETS_CALC_CORPUS_FN lists only formulas calling that function.
func TestEngineCorpus(t *testing.T) {
	dir := os.Getenv("SHEETS_CALC_CORPUS")
	if dir == "" {
		t.Skip("SHEETS_CALC_CORPUS names a folder of Excel-saved workbooks")
	}
	show := 3
	fmt.Sscan(os.Getenv("SHEETS_CALC_CORPUS_SHOW"), &show)
	only := strings.ToUpper(os.Getenv("SHEETS_CALC_CORPUS_FN"))
	var files []string
	filepath.WalkDir(dir, func(p string, d os.DirEntry, err error) error {
		if err == nil && !d.IsDir() && strings.EqualFold(filepath.Ext(p), ".xlsx") {
			files = append(files, p)
		}
		return nil
	})
	sort.Strings(files)
	var agree, differ, space, volatile int
	differBy := map[string][]string{}
	unsupBy := map[string]int{}
	unsup := 0
	for _, p := range files {
		data, err := os.ReadFile(p)
		if err != nil {
			t.Fatal(err)
		}
		wb, _, _, err := parse(data)
		if err != nil {
			continue
		}
		rs := wbResolver(wb)
		book := &gridBook{wb: wb, text: func(sheet string, row, col int, f string) string {
			return rs.calcText(&formulaAt{sheet: sheet, row: row, col: col, text: f})
		}}
		eng := engine.New(book)
		eng.Date1904 = wb.Date1904
		eng.Override = func(sheet string, row, col int) (engine.Value, bool) {
			s := book.sheet(sheet)
			if s == nil {
				return engine.Value{}, false
			}
			c := s.Get(row, col)
			if c.Formula != "" && !c.Pending && c.Kind != grid.Empty {
				return cellToCalc(c), true
			}
			return engine.Value{}, false
		}
		for _, s := range wb.Sheets {
			s.Each(s.Used(), func(r, col int, cell *grid.Cell) {
				if cell.Formula == "" || cell.Pending || cell.Kind == grid.Empty {
					return
				}
				fnsIn := formula.CalledFunctions(cell.Formula)
				if only != "" && !slices.Contains(fnsIn, only) {
					return
				}
				up := strings.ToUpper(cell.Formula)
				if strings.Contains(up, "TODAY(") || strings.Contains(up, "NOW(") || strings.Contains(up, "RAND") {
					volatile++
					return
				}
				text := book.text(s.Name, r, col, cell.Formula)
				v := eng.Eval(s.Name, r, col, text)
				if _, _, _, ok := book.ArrayAt(s.Name, r, col); ok {
					// Its computed array would stand in for the results
					// Excel stored in the cells it fills.
					eng.Reset()
				}
				want := *cell
				if v.Kind == engine.Error && v.Str == engine.ErrUnsupported {
					unsup++
					unsupBy[reasonKind(v.Why())]++
					return
				}
				got := calcToCell(v)
				if agreesCell(got, want) {
					agree++
					return
				}
				if got.Kind == grid.Text && want.Kind == grid.Text && strings.TrimSpace(got.Text) == strings.TrimSpace(want.Text) {
					space++
					if os.Getenv("SHEETS_CALC_CORPUS_SPACES") != "" && space <= 12 {
						t.Logf("spaces: %s %s!%s =%s: got %q, Excel %q", filepath.Base(p), s.Name, grid.CellName(r, col), cell.Formula, got.Text, want.Text)
					}
					return
				}
				differ++
				if len(fnsIn) == 0 {
					fnsIn = []string{"(no function)"}
				}
				ex := fmt.Sprintf("%s %s!%s =%s: got %s, Excel %s", filepath.Base(p), s.Name, grid.CellName(r, col), cell.Formula, show1(got), show1(want))
				for _, fn := range fnsIn {
					differBy[fn] = append(differBy[fn], ex)
				}
			})
		}
	}
	total := agree + differ + space + unsup
	t.Logf("%d workbooks, %d formulas with a result Excel stored (%d volatile left out)", len(files), total, volatile)
	t.Logf("agrees %d (%.2f%%) · differs %d (%.2f%%) · differs in spaces only %d · not supported %d",
		agree, 100*float64(agree)/float64(max(total, 1)), differ, 100*float64(differ)/float64(max(total, 1)), space, unsup)
	fns := make([]string, 0, len(differBy))
	for fn := range differBy {
		fns = append(fns, fn)
	}
	sort.Slice(fns, func(a, b int) bool { return len(differBy[fns[a]]) > len(differBy[fns[b]]) })
	for _, fn := range fns {
		t.Logf("differs in %s: %d", fn, len(differBy[fn]))
		for _, ex := range differBy[fn][:min(show, len(differBy[fn]))] {
			t.Logf("    %s", ex)
		}
	}
	reasons := make([]string, 0, len(unsupBy))
	for r := range unsupBy {
		reasons = append(reasons, r)
	}
	sort.Slice(reasons, func(a, b int) bool { return unsupBy[reasons[a]] > unsupBy[reasons[b]] })
	for _, r := range reasons[:min(25, len(reasons))] {
		t.Logf("not supported, %s: %d", r, unsupBy[r])
	}
}

func show1(c grid.Cell) string {
	if c.Kind == grid.Text {
		return strconv.Quote(c.Text)
	}
	return show(c)
}

// reasonKind shortens a reason to group by: the function named, or the
// reason's first words.
func reasonKind(why string) string {
	if i := strings.Index(why, " is not supported"); i > 0 {
		return why[:i] + " unsupported"
	}
	if i := strings.Index(why, " is a LAMBDA"); i > 0 {
		return "a LAMBDA name"
	}
	words := strings.Fields(why)
	return strings.Join(words[:min(6, len(words))], " ")
}
