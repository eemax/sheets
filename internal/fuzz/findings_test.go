package fuzz

import (
	"bytes"
	"io"
	"math/rand/v2"
	"os"
	"regexp"
	"slices"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
)

// afterCommand puts a command in cliMain's place that runs the real one and
// then does something to the workbook, for the commands match picks. The
// workbook is the first word of the command line that names a file.
func afterCommand(t *testing.T, match func(args []string) bool, do func(path string)) {
	t.Helper()
	old := cliMain
	t.Cleanup(func() { cliMain = old })
	cliMain = func(args []string, stdin io.Reader, out, errw io.Writer) int {
		code := old(args, stdin, out, errw)
		if match(args) {
			for _, a := range args {
				if strings.HasSuffix(a, ".xlsx") {
					do(a)
					break
				}
			}
		}
		return code
	}
}

func is(names ...string) func([]string) bool {
	return func(args []string) bool { return len(args) > 0 && slices.Contains(names, args[0]) }
}

// TestMissingWorkbookIsAFinding: an edit that leaves no workbook at its path
// (deleted or moved: data loss, the worst slip) is a finding like any other,
// with the seed, the workbook and the edits, shrunk, and the run goes on. (It
// was an error: the test stopped with the one line `open .../book.xlsx: no
// such file or directory`, and the other sequences never ran.)
func TestMissingWorkbookIsAFinding(t *testing.T) {
	pool := loadPool(t)

	t.Run("after an edit", func(t *testing.T) {
		afterCommand(t, is("clear"), func(path string) { os.Remove(path) })
		s, err := newSession(t.TempDir(), pool[0].Data)
		if err != nil {
			t.Fatal(err)
		}
		defer s.close()
		o, err := s.step(Edit{Class: "clear", Args: []string{"clear", FileArg, "Data!A1:B2", "--all"}}, 0)
		if err != nil {
			t.Fatalf("the edit is a finding, not an error: %v", err)
		}
		if len(o.broken) != 2 || o.broken[0].invariant != "missing" {
			t.Fatalf("want both editors' missing workbook, got %+v", o.broken)
		}
		for k, ed := range editors {
			d := o.broken[k].detail
			if !strings.HasPrefix(d, ed.name+" editor: the edit (exit ") || !strings.Contains(d, ") left no workbook at book.xlsx: open book.xlsx: no such file") {
				t.Errorf("the finding for the %s editor reads %q", ed.name, d)
			}
		}
	})

	t.Run("a refused edit that deleted it", func(t *testing.T) {
		old := cliMain
		t.Cleanup(func() { cliMain = old })
		cliMain = func(args []string, stdin io.Reader, out, errw io.Writer) int {
			os.Remove(args[1])
			return 1
		}
		s, err := newSession(t.TempDir(), pool[0].Data)
		if err != nil {
			t.Fatal(err)
		}
		defer s.close()
		o, err := s.step(Edit{Class: "set", Args: []string{"set", FileArg, "Data!B2", "5"}}, 0)
		if err != nil || len(o.broken) == 0 || o.broken[0].invariant != "missing" || !strings.Contains(o.broken[0].detail, "(exit 1)") {
			t.Fatalf("a refused edit that removed the file: %+v, %v", o.broken, err)
		}
	})

	t.Run("after undo", func(t *testing.T) {
		s, err := newSession(t.TempDir(), pool[0].Data)
		if err != nil {
			t.Fatal(err)
		}
		defer s.close()
		e := Edit{Class: "set", Args: []string{"set", FileArg, "Data!B2", "5"}}
		before := pool[0].Data
		if code, msg := run(editors[0].env, true, e, s.paths[0]); code != 0 {
			t.Fatal(msg)
		}
		after, err := os.ReadFile(s.paths[0])
		if err != nil {
			t.Fatal(err)
		}
		afterCommand(t, is("undo"), func(path string) { os.Remove(path) })
		v, err := checkUndo(0, s.paths[0], before, after)
		if err != nil || v == nil || v.invariant != "undo" || !strings.Contains(v.detail, "left no workbook at book.xlsx") {
			t.Fatalf("an undo that removed the file: %v, %v", v, err)
		}
		if got, _ := os.ReadFile(s.paths[0]); !bytes.Equal(got, after) {
			t.Error("the check did not put the edit's bytes back")
		}
	})

	t.Run("in the repeat", func(t *testing.T) {
		s, err := newSession(t.TempDir(), pool[0].Data)
		if err != nil {
			t.Fatal(err)
		}
		defer s.close()
		e := Edit{Class: "set", Args: []string{"set", FileArg, "Data!B2", "5"}}
		var before, after [2][]byte
		for k, ed := range editors {
			before[k] = pool[0].Data
			if code, msg := run(ed.env, false, e, s.paths[k]); code != 0 {
				t.Fatalf("%s: %s", ed.name, msg)
			}
			if after[k], err = os.ReadFile(s.paths[k]); err != nil {
				t.Fatal(err)
			}
		}
		afterCommand(t, is("set"), func(path string) { os.Remove(path) })
		v, err := s.repeat(e, before, after)
		if err != nil || v == nil || v.invariant != "repeat" || !strings.Contains(v.detail, "left no workbook at book.xlsx") {
			t.Fatalf("a repeat that removed the file: %v, %v", v, err)
		}
		for k := range editors {
			if got, _ := os.ReadFile(s.paths[k]); !bytes.Equal(got, after[k]) {
				t.Errorf("the check did not put the first run's bytes back (%s)", editors[k].name)
			}
		}
	})

	t.Run("in a run", func(t *testing.T) {
		if raceOn {
			t.Skip("takes 5 s under -race; the steps above judge the same code, and the plain run makes this one")
		}
		// Commands that delete the workbook, and others that move it away.
		removes := []string{"clear", "copy", "move", "sort"}
		afterCommand(t, is(removes...), func(path string) { os.Remove(path) })
		const edits = 40
		res, err := Run(pool, Options{Seed: 1, Edits: edits, Dir: t.TempDir()})
		if err != nil {
			t.Fatalf("the run stopped with an error, not a finding: %v", err)
		}
		if res.Edits != edits {
			t.Errorf("the run stopped after %d of %d edits", res.Edits, edits)
		}
		// The run's other findings are the real ones in known.txt.
		var missing []Finding
		for _, f := range res.Findings {
			if f.Invariant == "missing" {
				missing = append(missing, f)
			}
		}
		if len(missing) < 2 {
			t.Fatalf("%d findings of missing in %d edits: the run should go on past the first", len(missing), res.Edits)
		}
		for _, f := range missing {
			if !slices.Contains(removes, f.Class) {
				t.Errorf("finding %s, want an edit of one of %v that left no workbook", f.Signature, removes)
			}
			if last := f.Edits[len(f.Edits)-1]; last.Class != f.Class {
				t.Errorf("the sequence should end at the edit that did it: %s", last)
			}
		}
		f := missing[0]
		g, err := Shrink(pool, f, t.TempDir())
		if err != nil {
			t.Fatal(err)
		}
		if g.Invariant != "missing" || len(g.Edits) > len(f.Edits) || g.Edits[len(g.Edits)-1].Class != f.Class {
			t.Errorf("shrunk to %s with %d of %d edits", g.Signature, len(g.Edits), len(f.Edits))
		}
		report := g.Report()
		for _, want := range []string{"finding " + f.Class + " missing " + strconv.FormatUint(f.Seed, 10), "workbook: testdata/pool/" + f.Workbook, "sheets ", "left no workbook at book.xlsx"} {
			if !strings.Contains(report, want) {
				t.Errorf("the report has no %q:\n%s", want, report)
			}
		}
		if strings.Contains(report, os.TempDir()) {
			t.Errorf("the report names a temporary directory:\n%s", report)
		}
		// known.txt can list one (it is a bug in sheets like another).
		if _, err := ParseKnown(f.Signature.String() + "\n"); err != nil {
			t.Error(err)
		}
		t.Log(report)
	})
}

// poolSheetNames are the names of the sheets the pool's workbooks hold.
func poolSheetNames(t *testing.T) map[string]bool {
	t.Helper()
	names := map[string]bool{}
	for i := range loadPool(t) {
		for _, n := range poolBook(t, i).SheetNames() {
			names[n] = true
		}
	}
	return names
}

// TestPanicInJudgingIsAFinding: a panic in the step goroutine outside the
// commands (in sheets' reader on the file an editor saved, or in the fuzzer's
// own checks) is a finding of the invariant panic, with the seed, the
// workbook and the edits, and the run goes on. (It was a crash of the test
// binary: a dump of goroutines, and every later test skipped.)
func TestPanicInJudgingIsAFinding(t *testing.T) {
	pool := loadPool(t)
	known := poolSheetNames(t)
	// A reader that cannot read a workbook that holds a sheet the pool's do not.
	oldRead := readBytes
	t.Cleanup(func() { readBytes = oldRead })
	readBytes = func(data []byte) (*grid.Workbook, error) {
		wb, err := oldRead(data)
		if wb != nil {
			for _, n := range wb.SheetNames() {
				if !known[n] {
					panic("reader bug on " + n)
				}
			}
		}
		return wb, err
	}

	t.Run("a step", func(t *testing.T) {
		s, err := newSession(t.TempDir(), pool[0].Data)
		if err != nil {
			t.Fatal(err)
		}
		defer s.close()
		o, err := s.step(Edit{Class: "sheet-add", Args: []string{"sheet", "add", FileArg, "Backup"}}, 0)
		if err != nil || len(o.broken) != 1 || o.broken[0].invariant != "panic" {
			t.Fatalf("a panic in the reader: %+v, %v", o.broken, err)
		}
		d := o.broken[0].detail
		for _, want := range []string{"judging the files the editors saved: reader bug on Backup", "compareEditors"} {
			if !strings.Contains(d, want) {
				t.Errorf("the finding has no %q:\n%s", want, d)
			}
		}
	})

	t.Run("in a run", func(t *testing.T) {
		if raceOn {
			t.Skip("takes 7 s under -race; the step above judges the same code, and the plain run makes this one")
		}
		const edits = 40
		res, err := Run(pool, Options{Seed: 1, Edits: edits, Dir: t.TempDir()})
		if err != nil {
			t.Fatalf("the run stopped with an error, not a finding: %v", err)
		}
		if res.Edits != edits {
			t.Errorf("the run stopped after %d of %d edits", res.Edits, edits)
		}
		if len(res.Findings) < 2 {
			t.Fatalf("%d findings in %d edits: the run should go on past the first", len(res.Findings), res.Edits)
		}
		for _, f := range res.Findings {
			if f.Invariant != "panic" || !strings.Contains(f.Detail, "reader bug") {
				t.Errorf("finding %s: %s", f.Signature, f.Detail)
			}
		}
		f := res.Findings[0]
		g, err := Shrink(pool, f, t.TempDir())
		if err != nil {
			t.Fatal(err)
		}
		if g.Invariant != "panic" || len(g.Edits) > len(f.Edits) {
			t.Errorf("shrunk to %s with %d of %d edits", g.Signature, len(g.Edits), len(f.Edits))
		}
		report := g.Report()
		for _, want := range []string{"finding " + f.Class + " panic " + strconv.FormatUint(f.Seed, 10), "workbook: testdata/pool/" + f.Workbook, "reader bug"} {
			if !strings.Contains(report, want) {
				t.Errorf("the report has no %q:\n%s", want, report)
			}
		}
		t.Log(strings.SplitN(report, "goroutine", 2)[0])
	})
}

// TestGeneratorPanicIsAFinding: the generator panics on the workbook the last
// edit saved. The finding is that edit's, with the edits so far.
func TestGeneratorPanicIsAFinding(t *testing.T) {
	pool := loadPool(t)
	old := makeEdit
	t.Cleanup(func() { makeEdit = old })
	calls := 0
	makeEdit = func(rng *rand.Rand, wb *grid.Workbook) Edit {
		if calls++; calls > 1 {
			panic("generator bug")
		}
		return old(rng, wb)
	}
	// The first sequence that asks for a second edit.
	var f *Finding
	var seed uint64
	for i := 0; f == nil && i < 20; i++ {
		var err error
		calls, seed = 0, SequenceSeed(1, i)
		if f, _, _, err = sequence(pool, seed, t.TempDir(), 12, 12, nil, nil); err != nil {
			t.Fatal(err)
		}
	}
	if f == nil {
		t.Fatal("no sequence asked for a second edit")
	}
	if f.Invariant != "panic" || f.Seed != seed || len(f.Edits) != 1 || f.Class != f.Edits[0].Class || !strings.Contains(f.Detail, "generator bug") {
		t.Errorf("finding %s after %d edits: %s", f.Signature, len(f.Edits), f.Detail)
	}
	// A panic on the pool file itself is no edit's: it is the pool's, an error.
	calls = 1
	if _, _, _, err := sequence(pool, seed, t.TempDir(), 12, 12, nil, nil); err == nil || !strings.Contains(err.Error(), "the pool file") {
		t.Errorf("a panic on a pool file: %v", err)
	}
}

// TestNoPlaceholderTyped: no edit types a literal {r} or {s} (the places of a
// formula's row and sheet): sheets accepts the formula, Excel cannot read it,
// and the edits after it run on a workbook Excel would repair. A range set
// types the formula of the range's first cell, which sheets fills down and
// across with its references shifted: that is the relative formula the class
// is for.
func TestNoPlaceholderTyped(t *testing.T) {
	// The formulas with a row in them, as patterns for the rows they were typed with.
	var relative []*regexp.Regexp
	for _, f := range plainFormulas {
		if strings.Contains(f, "{r}") {
			re := regexp.QuoteMeta(f)
			re = strings.ReplaceAll(re, `\{r\}`, `(\d+)`)
			relative = append(relative, regexp.MustCompile("^"+strings.ReplaceAll(re, `\{s\}`, `.+`)+"$"))
		}
	}
	rows := func(formula string) ([]int, bool) {
		for _, re := range relative {
			if m := re.FindStringSubmatch(formula); m != nil {
				var rs []int
				for _, d := range m[1:] {
					n, _ := strconv.Atoi(d)
					rs = append(rs, n)
				}
				return rs, true
			}
		}
		return nil, false
	}

	draws := 150
	if raceOn {
		draws = 40
	}
	var set, ranges, filled int
	// check judges the pairs of cell and value a set types.
	check := func(where string, pairs []string) {
		for i := 0; i+1 < len(pairs); i += 2 {
			target, value := pairs[i], pairs[i+1]
			set++
			rs, rel := rows(value)
			if !rel {
				continue
			}
			rg, err := grid.ParseRange(target[strings.LastIndexByte(target, '!')+1:])
			if err != nil {
				t.Errorf("%s: %q is no range: %v", where, target, err)
				continue
			}
			if rg.R2 > rg.R1 || rg.C2 > rg.C1 {
				ranges++
			}
			for _, r := range rs {
				if r != rg.R1 {
					t.Errorf("%s: %s is typed with row %d: the formula of the range's first row is %d", where, target, r, rg.R1)
				}
			}
			if rg.R2 > rg.R1 {
				filled++
			}
		}
	}
	for i := range loadPool(t) {
		wb := poolBook(t, i)
		rng := rand.New(rand.NewPCG(uint64(i), 7))
		for range draws {
			e := generate(rng, wb)
			text := strings.Join(e.Args, "\n") + "\n" + e.Stdin
			for _, ph := range []string{"{r}", "{s}"} {
				if strings.Contains(text, ph) {
					t.Fatalf("an edit types a literal %s:\n%s", ph, e)
				}
			}
			switch e.Class {
			case "set":
				check(e.String(), e.Args[2:])
			case "do":
				for _, line := range strings.Split(strings.TrimSuffix(e.Stdin, "\n"), "\n") {
					w, err := words.Texts(line)
					if err != nil {
						t.Fatalf("%q: %v", line, err)
					}
					if len(w) > 0 && w[0] == "set" {
						check(line, w[1:])
					}
				}
			}
		}
	}
	t.Logf("of %d cells and ranges a set typed, %d were ranges with a relative formula, %d of them over several rows", set, ranges, filled)
	if filled < 10 {
		t.Errorf("only %d range sets typed a relative formula over several rows: the class is not exercised", filled)
	}
}
