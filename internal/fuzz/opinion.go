package fuzz

import (
	"bytes"
	"context"
	"fmt"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// LibreOffice's second opinion, in long mode only: the workbooks some
// sequences leave are recalculated by LibreOffice, and the values sheets
// stored for plain formulas are compared with LibreOffice's. It is reported,
// never a finding: LibreOffice is not Excel, and differs from it in places.
//
// Only workbooks that started from an openpyxl file are taken: openpyxl
// stores no results, so every result in them is one sheets computed.
// (XlsxWriter stores a result of its own for every formula, 0 for a number,
// which sheets keeps where no edit reaches.)

// Sample is a workbook a sequence left, for the second opinion.
type Sample struct {
	Seq      uint64
	Pool     uint64
	Workbook string
	Data     []byte
}

// Opinion is LibreOffice's second opinion on a run's samples.
type Opinion struct {
	Files     int // workbooks sent
	Converted int // workbooks LibreOffice recalculated and saved
	Compared  int // plain formulas with a result sheets stored
	Agreed    int
	Differ    int
	// ByFunctions counts the differences by the functions the formula calls;
	// Examples are the first few.
	ByFunctions map[string]int
	Examples    []string
	Err         string // why LibreOffice could not give one
}

// String is the opinion as the log prints it.
func (o Opinion) String() string {
	if o.Err != "" && o.Converted == 0 {
		return fmt.Sprintf("none: %s", o.Err)
	}
	var b strings.Builder
	fmt.Fprintf(&b, "%d workbooks sent, %d recalculated; %d plain formulas with a result sheets stored: %d agree, %d differ", o.Files, o.Converted, o.Compared, o.Agreed, o.Differ)
	if o.Err != "" {
		fmt.Fprintf(&b, " (%s)", o.Err)
	}
	if len(o.ByFunctions) > 0 {
		var keys []string
		for k := range o.ByFunctions {
			keys = append(keys, k)
		}
		sort.Slice(keys, func(i, j int) bool {
			if o.ByFunctions[keys[i]] != o.ByFunctions[keys[j]] {
				return o.ByFunctions[keys[i]] > o.ByFunctions[keys[j]]
			}
			return keys[i] < keys[j]
		})
		var by []string
		for _, k := range keys {
			by = append(by, fmt.Sprintf("%s %d", k, o.ByFunctions[k]))
		}
		fmt.Fprintf(&b, "; differences by the functions called: %s", strings.Join(by, ", "))
		for _, e := range o.Examples {
			b.WriteString("\n    " + e)
		}
	}
	return b.String()
}

// opinionTimeout bounds LibreOffice's work on all the samples.
const opinionTimeout = 10 * time.Minute

// SecondOpinion has LibreOffice (soffice) recalculate the samples, in dir,
// and compares the values of their plain formulas.
func SecondOpinion(soffice string, samples []Sample, dir string) Opinion {
	o := Opinion{ByFunctions: map[string]int{}}
	in, out, profile := filepath.Join(dir, "in"), filepath.Join(dir, "out"), filepath.Join(dir, "profile")
	for _, d := range []string{in, out, filepath.Join(profile, "user")} {
		if err := os.MkdirAll(d, 0o755); err != nil {
			o.Err = err.Error()
			return o
		}
	}
	// Recalculate every file on load (OOXMLRecalcMode 0: always): else
	// LibreOffice shows the results stored.
	xcu := `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
`
	if err := os.WriteFile(filepath.Join(profile, "user", "registrymodifications.xcu"), []byte(xcu), 0o644); err != nil {
		o.Err = err.Error()
		return o
	}
	var names []string
	byName := map[string]Sample{}
	for _, s := range samples {
		name := fmt.Sprintf("%d.xlsx", s.Seq)
		if err := os.WriteFile(filepath.Join(in, name), s.Data, 0o644); err != nil {
			o.Err = err.Error()
			return o
		}
		names = append(names, filepath.Join(in, name))
		byName[name] = s
	}
	o.Files = len(names)
	if len(names) == 0 {
		o.Err = "no workbooks to send"
		return o
	}
	ctx, cancel := context.WithTimeout(context.Background(), opinionTimeout)
	defer cancel()
	args := append([]string{"-env:UserInstallation=file://" + profile, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out}, names...)
	cmd := exec.CommandContext(ctx, soffice, args...)
	var buf bytes.Buffer
	cmd.Stdout, cmd.Stderr = &buf, &buf
	if err := cmd.Run(); err != nil {
		o.Err = fmt.Sprintf("soffice: %v: %s", err, lastLines(buf.String(), 2))
	}
	sort.Strings(names)
	for _, n := range names {
		name := filepath.Base(n)
		lo, err := os.ReadFile(filepath.Join(out, name))
		if err != nil {
			continue
		}
		mine, err1 := readBytes(byName[name].Data)
		theirs, err2 := readBytes(lo)
		if err1 != nil || err2 != nil {
			continue
		}
		o.Converted++
		s := byName[name]
		o.compare(mine, theirs, fmt.Sprintf("sequence %d (pool %d, %s)", s.Seq, s.Pool, s.Workbook))
	}
	return o
}

// compare counts the plain formulas of a workbook whose result sheets stored,
// against LibreOffice's.
func (o *Opinion) compare(mine, theirs *grid.Workbook, where string) {
	for _, s := range mine.Sheets {
		t := theirs.Sheet(s.Name)
		if t == nil {
			continue
		}
		arrays := append(append([]grid.Range{}, s.Spills...), s.Arrays...)
		u := s.Used()
		if u.R2 < 1 {
			continue
		}
		s.Each(u, func(r, c int, cl *grid.Cell) {
			if cl.Formula == "" || cl.Pending || cl.Unprefixed || marker(cl.Formula) != Plain {
				return
			}
			for _, a := range arrays {
				if a.Contains(r, c) {
					return
				}
			}
			o.Compared++
			lo := t.Get(r, c)
			if sameResult(*cl, lo) {
				o.Agreed++
				return
			}
			o.Differ++
			fns := formula.CalledFunctions(cl.Formula)
			key := "(no function)"
			if len(fns) > 0 {
				key = strings.Join(fns, " ")
			}
			if o.ByFunctions[key]++; len(o.Examples) < 10 {
				o.Examples = append(o.Examples, fmt.Sprintf("%s %s!%s =%s: sheets %s, LibreOffice %s", where, s.Name, grid.CellName(r, c), cl.Formula, value(*cl), value(lo)))
			}
		})
	}
}

// sameResult compares two results: numbers to nine significant digits, an
// empty text and an empty cell alike (LibreOffice may store either), and
// sheets' #REF! as LibreOffice's #NAME? (LibreOffice reads a #REF! in a
// formula's text as a name it does not know, so the formulas reading one give
// #NAME? there).
func sameResult(a, b grid.Cell) bool {
	empty := func(c grid.Cell) bool { return c.Kind == grid.Empty || c.Kind == grid.Text && c.Text == "" }
	switch {
	case empty(a) || empty(b):
		return empty(a) && empty(b)
	case a.Kind == grid.Error && b.Kind == grid.Error && a.Text == "#REF!" && b.Text == "#NAME?":
		return true
	case a.Kind != b.Kind:
		return false
	case a.Kind == grid.Number:
		return a.Num == b.Num || math.Abs(a.Num-b.Num) <= 1e-9*math.Max(math.Abs(a.Num), math.Abs(b.Num))
	case a.Kind == grid.Bool:
		return a.Bool == b.Bool
	}
	return a.Text == b.Text
}
