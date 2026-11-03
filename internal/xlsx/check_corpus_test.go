package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsxtest"
)

// TestCheckCorpus runs sheets check --fix over a folder of workbooks and
// holds it to its promises, checked here on its own rather than through
// the comparison the fix makes of itself:
//
//   - the fixed file opens, in sheets and (with SHEETS_CHECK_CORPUS_SOFFICE
//     naming LibreOffice) in LibreOffice;
//   - no result Excel stored and vouched for changes, nor any value but
//     formula results;
//   - a second fix changes nothing.
//
// Off unless SHEETS_CHECK_CORPUS names the folder (a copy: the test only
// reads it). SHEETS_CHECK_CORPUS_OUT keeps the fixed files there.
// SHEETS_CHECK_CORPUS_EXCEL names a folder holding the same workbooks as
// Excel saved them, under the same paths (the corpus being, say, their
// openpyxl round trips): each result the fix stores is compared with the
// one Excel stored.
// SHEETS_STREAM=never runs the fixes through the full editor.
func TestCheckCorpus(t *testing.T) {
	dir := os.Getenv("SHEETS_CHECK_CORPUS")
	if dir == "" {
		t.Skip("SHEETS_CHECK_CORPUS names a folder of workbooks")
	}
	soffice := os.Getenv("SHEETS_CHECK_CORPUS_SOFFICE")
	excelDir := os.Getenv("SHEETS_CHECK_CORPUS_EXCEL")
	var asExcel, unlikeExcel, notInExcel int
	var unlike []string
	unlikeBy, unlikeEx := map[string]int{}, map[string]string{}
	out := os.Getenv("SHEETS_CHECK_CORPUS_OUT")
	if out == "" {
		out = t.TempDir()
	}
	var files []string
	filepath.WalkDir(dir, func(p string, d os.DirEntry, err error) error {
		if err == nil && !d.IsDir() && strings.EqualFold(filepath.Ext(p), ".xlsx") {
			files = append(files, p)
		}
		return nil
	})
	sort.Strings(files)
	opts := CheckOptions{Budget: 2 * time.Minute, Compare: 10 * time.Second}
	var (
		badOrder                                                                                               int
		unreadable, nothing, fixed, refused, failedOpen, loFailed, secondChanged, vouchedChanged, otherChanged int
		formulas, vouched, stored, replaced, dropped, prefixed, unknown, compared, agreed, differ              int
		slowest                                                                                                time.Duration
		slowFile                                                                                               string
	)
	profile := filepath.Join(out, "lo-profile")
	for _, p := range files {
		data, err := os.ReadFile(p)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := ReadBytes(data); err != nil {
			unreadable++ // sheets cannot read it as given: not the fix's
			continue
		}
		start := time.Now()
		res, err := CheckResults(p, data, opts)
		if d := time.Since(start); d > slowest {
			slowest, slowFile = d, filepath.Base(p)
		}
		if err == nil && res.FixError != "" {
			err = errors.New(res.FixError)
		}
		if err != nil {
			refused++
			t.Errorf("%s: %v", filepath.Base(p), err)
			continue
		}
		formulas += res.Formulas
		vouched += res.Vouched
		unknown += len(res.Unknown)
		compared += res.Compared
		agreed += res.Agreed
		differ += len(res.Differ)
		for _, c := range res.Changes {
			switch c.Kind {
			case "stored":
				stored++
			case "replaced":
				replaced++
			case "dropped":
				dropped++
			case "prefixed":
				prefixed++
			}
		}
		if res.Data == nil {
			nothing++
			continue
		}
		fixed++
		before, _ := ReadBytes(data)
		after, err := ReadBytes(res.Data)
		if err != nil {
			failedOpen++
			t.Errorf("%s: sheets cannot open the fixed file: %v", filepath.Base(p), err)
			continue
		}
		if err := xlsxtest.CheckOrder(res.Data); err != nil {
			badOrder++
			t.Errorf("%s: cells out of order: %v", filepath.Base(p), err)
		}
		v, o := promisesKept(before, after)
		if len(v) > 0 {
			vouchedChanged++
			t.Errorf("%s: results Excel vouched for changed: %s", filepath.Base(p), strings.Join(v[:min(3, len(v))], "; "))
		}
		if len(o) > 0 {
			otherChanged++
			t.Errorf("%s: values changed: %s", filepath.Base(p), strings.Join(o[:min(3, len(o))], "; "))
		}
		if excelDir != "" {
			if orig, err := os.ReadFile(filepath.Join(excelDir, strings.TrimPrefix(p, dir))); err == nil {
				if ex, err := ReadBytes(orig); err == nil {
					for _, c := range res.Changes {
						if c.Kind != "stored" && c.Kind != "replaced" {
							continue
						}
						s := ex.Sheet(c.Sheet)
						if s == nil {
							notInExcel++
							continue
						}
						e := s.Get(c.Row, c.Col)
						switch {
						case e.Formula == "" || e.Pending:
							notInExcel++
						case agreesCell(c.After, valueOnly(e)):
							asExcel++
						default:
							unlikeExcel++
							line := fmt.Sprintf("%s %s =%s: sheets %s, Excel %s", filepath.Base(p),
								grid.FormatRef(c.Sheet, grid.Range{R1: c.Row, C1: c.Col, R2: c.Row, C2: c.Col}), e.Formula, show1(c.After), show1(valueOnly(e)))
							unlike = append(unlike, line)
							fns := formula.CalledFunctions(e.Formula)
							key := "(no function)"
							if len(fns) > 0 {
								key = strings.Join(fns, " ")
							}
							if unlikeBy[key]++; unlikeBy[key] == 1 {
								unlikeEx[key] = line
							}
						}
					}
				}
			}
		}
		name := filepath.Join(out, strings.ReplaceAll(strings.TrimPrefix(p, dir), string(filepath.Separator), "_"))
		if err := os.WriteFile(name, res.Data, 0o644); err != nil {
			t.Fatal(err)
		}
		again, err := CheckResults(name, res.Data, opts)
		if err == nil && again.FixError != "" {
			err = errors.New(again.FixError)
		}
		if err != nil {
			secondChanged++
			t.Errorf("%s: second run: %v", filepath.Base(p), err)
		} else if again.Data != nil {
			secondChanged++
			c := again.Changes[0]
			t.Errorf("%s: a second fix changes %d cells (%s %s)", filepath.Base(p), len(again.Changes), c.Kind, grid.FormatRef(c.Sheet, grid.Range{R1: c.Row, C1: c.Col, R2: c.Row, C2: c.Col}))
		}
		if soffice != "" {
			if err := libreOfficeOpens(soffice, profile, name, filepath.Join(out, "lo")); err != nil {
				loFailed++
				t.Errorf("%s: LibreOffice: %v", filepath.Base(p), err)
			}
		}
	}
	t.Logf("%d workbooks: %d unreadable as given, %d needing nothing, %d fixed, %d refused", len(files), unreadable, nothing, fixed, refused)
	t.Logf("formulas %d, results vouched for %d; changed cells: %d stored, %d stale replaced, %d stale dropped, %d prefixed; %d left unknown",
		formulas, vouched, stored, replaced, dropped, prefixed, unknown)
	t.Logf("stored results against the calculator: %d compared, %d agree, %d differ", compared, agreed, differ)
	t.Logf("promises: %d fixed files sheets cannot open, %d LibreOffice cannot open (checked: %v), %d with a vouched result changed, %d with other values changed, %d changed by a second fix or failing on it, %d with cells out of order",
		failedOpen, loFailed, soffice != "", vouchedChanged, otherChanged, secondChanged, badOrder)
	t.Logf("slowest: %s, %s", slowFile, slowest.Round(time.Millisecond))
	if excelDir != "" {
		t.Logf("results stored, against Excel's: %d the same, %d different, %d Excel did not store", asExcel, unlikeExcel, notInExcel)
		keys := make([]string, 0, len(unlikeBy))
		for k := range unlikeBy {
			keys = append(keys, k)
		}
		sort.Slice(keys, func(i, j int) bool { return unlikeBy[keys[i]] > unlikeBy[keys[j]] })
		for _, k := range keys[:min(len(keys), 30)] {
			t.Logf("    %5d  %s   e.g. %s", unlikeBy[k], k, unlikeEx[k])
		}
		os.WriteFile(filepath.Join(out, "unlike-excel.txt"), []byte(strings.Join(unlike, "\n")+"\n"), 0o644)
	}
}

// promisesKept compares two readings of a workbook: results Excel stored
// and vouched for that changed, and other values that changed (but for
// those of array formulas).
func promisesKept(before, after *grid.Workbook) (vouched, other []string) {
	for i, sb := range before.Sheets {
		if i >= len(after.Sheets) {
			other = append(other, "sheet "+sb.Name+" gone")
			continue
		}
		sa := after.Sheets[i]
		sb.Each(sb.Used(), func(r, c int, b *grid.Cell) {
			a := sa.Get(r, c)
			ref := grid.FormatRef(sb.Name, grid.Range{R1: r, C1: c, R2: r, C2: c})
			same := b.Kind == a.Kind && b.Text == a.Text && b.Bool == a.Bool && (b.Num == a.Num || b.Num != b.Num && a.Num != a.Num)
			switch {
			case b.Formula != "" && !b.Pending:
				if !same || a.Pending || a.Formula != b.Formula {
					vouched = append(vouched, fmt.Sprintf("%s =%s: %v → %v", ref, b.Formula, *b, a))
				}
			case b.Formula == "" && !same && !inArray(sb, r, c) && !inArray(sa, r, c):
				other = append(other, fmt.Sprintf("%s: %v → %v", ref, *b, a))
			case b.Formula != "" && a.Formula != b.Formula:
				other = append(other, fmt.Sprintf("%s: formula =%s → =%s", ref, b.Formula, a.Formula))
			}
		})
	}
	return vouched, other
}

// libreOfficeOpens converts a workbook with LibreOffice, which fails on a
// file it cannot open.
func libreOfficeOpens(soffice, profile, path, dir string) error {
	os.MkdirAll(dir, 0o755)
	want := filepath.Join(dir, strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))+".ods")
	os.Remove(want)
	cmd := exec.Command(soffice, "-env:UserInstallation=file://"+profile, "--headless", "--calc", "--convert-to", "ods", "--outdir", dir, path)
	var buf bytes.Buffer
	cmd.Stdout, cmd.Stderr = &buf, &buf
	done := make(chan error, 1)
	if err := cmd.Start(); err != nil {
		return err
	}
	go func() { done <- cmd.Wait() }()
	select {
	case err := <-done:
		if err != nil {
			return fmt.Errorf("%v: %s", err, buf.String())
		}
	case <-time.After(3 * time.Minute):
		cmd.Process.Kill()
		return fmt.Errorf("timed out")
	}
	if st, err := os.Stat(want); err != nil || st.Size() == 0 {
		return fmt.Errorf("no output: %s", strings.TrimSpace(buf.String()))
	}
	return nil
}
