package fuzz

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/undo"
)

// The default mode, in every go test ./...: a fixed seed and a fixed number
// of edits, and under -race a time cap. See README.md.
const (
	defaultSeed  = 1
	defaultEdits = 2000
	// Under -race each edit takes about ten times as long: the run stops
	// at the cap, once raceMinEdits are made. Without -race there is no
	// cap: a run always makes its defaultEdits, so two runs make the same
	// edits and a slow machine takes longer instead of making fewer (go
	// test's own timeout is the backstop).
	raceCap      = 40 * time.Second
	raceMinEdits = 200
	// New findings past this many are printed unshrunk.
	maxShrunk = 5
	// A plain run must put this many edits to the streaming editor's own code
	// (see streamed), or the editors invariant compares the full editor with
	// itself.
	minStreamed = 300
)

func TestMain(m *testing.M) {
	// The undo journal and the parse cache in a directory of the run's own,
	// and undo on whatever the environment says.
	dir, err := os.MkdirTemp("", "sheets-fuzz-")
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	os.Setenv("SHEETS_CACHE_DIR", dir)
	os.Setenv("SHEETS_CACHE", "off")
	os.Unsetenv("SHEETS_UNDO")
	if raceOn {
		repeats = 3
	}
	code := m.Run()
	os.RemoveAll(dir)
	os.Exit(code)
}

func loadPool(t *testing.T) []Workbook {
	t.Helper()
	pool, err := LoadPool(filepath.Join("testdata", "pool"))
	if err != nil {
		t.Fatal(err)
	}
	return pool
}

// TestPool: every pool file holds the invariants before any edit, and is
// small (the pool is about 1 MB in all).
func TestPool(t *testing.T) {
	pool := loadPool(t)
	total := 0
	producers := map[string]int{}
	for _, wb := range pool {
		total += len(wb.Data)
		producers[strings.SplitN(wb.Name, "-", 2)[0]]++
		for _, v := range checkFile(wb.Data) {
			t.Errorf("%s: %s: %s", wb.Name, v.invariant, v.detail)
		}
		if _, d := compareEditors(wb.Data, wb.Data); d != "" {
			t.Errorf("%s: differs from itself: %s", wb.Name, d)
		}
		cells, _, err := formulas(mustZip(t, wb.Data))
		if err != nil {
			t.Fatal(err)
		}
		for _, f := range cells {
			if marker(f.text) == "" {
				t.Errorf("%s: %s!%s: =%s names no kind", wb.Name, f.sheet, f.cell, f.text)
			}
		}
	}
	if len(pool) > 45 || total > 1<<20 {
		t.Errorf("pool: %d files, %d bytes; keep it to about 40 files and 1 MB", len(pool), total)
	}
	if producers["openpyxl"] == 0 || producers["xlsxwriter"] == 0 {
		t.Errorf("pool: want files from both producers, have %v", producers)
	}
}

// TestPoolResults: the pool holds a formula result of every stored form but
// an error (XlsxWriter stores results, openpyxl none), and shared formulas,
// which neither producer writes: fuzzgen.py rewrites a filled-down column.
func TestPoolResults(t *testing.T) {
	seen := map[string]int{}
	for _, wb := range loadPool(t) {
		_, cells := inspect(wb.Data)
		files := map[string]bool{}
		for _, f := range cells {
			files[f.result()] = true
			if f.shared {
				files[resShared] = true
			}
		}
		for form := range files {
			seen[form]++
		}
	}
	for _, f := range []string{resNumber, resText, resEmptyText, resBoolean, resNone, resShared} {
		if seen[f] == 0 {
			t.Errorf("no pool file holds a %s", f)
		}
	}
	t.Logf("pool files holding each form: %v", seen)
}

// TestPoolKinds: the pool holds every kind, for the kinds invariant to
// have something to keep.
func TestPoolKinds(t *testing.T) {
	seen := map[string]int{}
	for _, wb := range loadPool(t) {
		cells, _, err := formulas(mustZip(t, wb.Data))
		if err != nil {
			t.Fatal(err)
		}
		for _, f := range cells {
			seen[f.kind]++
		}
	}
	for _, k := range []string{Plain, Legacy, Dynamic} {
		if seen[k] == 0 {
			t.Errorf("no %s in the pool", kindName[k])
		}
	}
}

// TestRefuse: the must-refuse cases, pool files with one defect each, are
// caught by the invariant their name starts with.
func TestRefuse(t *testing.T) {
	names, err := filepath.Glob(filepath.Join("testdata", "refuse", "*.xlsx"))
	if err != nil || len(names) == 0 {
		t.Fatalf("no refuse cases: %v", err)
	}
	for _, name := range names {
		base := filepath.Base(name)
		t.Run(base, func(t *testing.T) {
			data, err := os.ReadFile(name)
			if err != nil {
				t.Fatal(err)
			}
			invariant := strings.SplitN(base, "-", 2)[0]
			switch invariant {
			case "xml", "kinds":
				vs := checkFile(data)
				if len(vs) == 0 {
					t.Fatalf("not refused")
				}
				if vs[0].invariant != invariant {
					t.Fatalf("refused by %s, not %s: %s", vs[0].invariant, invariant, vs[0].detail)
				}
				t.Logf("%s: %s", vs[0].invariant, vs[0].detail)
			case "editors":
				// editors-<what>--<pool file>.xlsx differs from that file.
				i := strings.Index(base, "--")
				if i < 0 {
					t.Fatalf("name the pool file after --")
				}
				orig, err := os.ReadFile(filepath.Join("testdata", "pool", base[i+2:]))
				if err != nil {
					t.Fatal(err)
				}
				if vs := checkFile(data); len(vs) > 0 {
					t.Fatalf("should hold the file's invariants: %s: %s", vs[0].invariant, vs[0].detail)
				}
				_, d := compareEditors(orig, data)
				if d == "" {
					t.Fatalf("not refused")
				}
				t.Log(d)
				if p := partsDiffer(orig, data); !strings.Contains(p, "xl/worksheets/sheet1.xml differs") {
					t.Errorf("undo's comparison names %q, not the sheet", p)
				}
			default:
				t.Fatalf("no invariant %q", invariant)
			}
		})
	}
}

// TestUndoRefuses: the undo check catches an undo that restores other
// bytes than the edit's before, and one that fails.
func TestUndoRefuses(t *testing.T) {
	pool := loadPool(t)
	before, after := pool[0].Data, pool[1].Data
	path := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(path, after, 0o644); err != nil {
		t.Fatal(err)
	}
	if v, err := checkUndo(0, path, before, after); err != nil || v == nil || v.invariant != "undo" {
		t.Fatalf("an undo with nothing recorded: %v, %v", v, err)
	} else {
		t.Log(v.detail)
	}
	// The journal records a third file as the state before.
	if err := undo.Record(path, pool[2].Data, after, "sheets test"); err != nil {
		t.Fatal(err)
	}
	if v, err := checkUndo(1, path, before, after); err != nil || v == nil || v.invariant != "undo" {
		t.Fatalf("an undo restoring other bytes: %v, %v", v, err)
	} else {
		t.Log(v.detail)
	}
	if got, _ := os.ReadFile(path); !bytes.Equal(got, after) {
		t.Error("the check did not put the edit's bytes back")
	}
	// And the true before passes.
	if err := undo.Record(path, before, after, "sheets test"); err != nil {
		t.Fatal(err)
	}
	if v, err := checkUndo(0, path, before, after); err != nil || v != nil {
		t.Fatalf("a true undo: %v, %v", v, err)
	}
}

// TestStrictXML: strictXML refuses the XML Go's decoder lets through and
// expat refuses (each case's verdict is expat's, with namespaces on), and
// passes what is well-formed.
func TestStrictXML(t *testing.T) {
	for _, c := range []struct {
		name, doc string
		ok        bool
	}{
		{"declaration and namespaces", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<a xmlns=\"u\" xmlns:r=\"v\" r:id=\"1\"><b/></a>", true},
		{"two prefixes for one namespace, distinct names", `<a xmlns:x="u" xmlns:y="u"><b x:p="1" y:q="1" q="1"/></a>`, true},
		{"xml:space", `<a xml:space="preserve"><b/></a>`, true},
		{"a prefix of an ancestor", `<a xmlns:x="u"><b><x:c/></b></a>`, true},
		{"a prefix redeclared", `<a xmlns:x="u"><b xmlns:x="v"><x:c/></b></a>`, true},
		{"a byte order mark", "\ufeff<?xml version=\"1.0\"?><a/>", true},
		{"a comment after the root", "<a/>\n<!-- c -->\n", true},
		{"one local name, prefixed and not", `<a xmlns:x="u" x:p="1" p="2"/>`, true},
		{"an attribute twice", `<c r="B2" r="B2"/>`, false},
		{"an attribute twice by namespace", `<a xmlns:x="u" xmlns:y="u" x:p="1" y:p="2"/>`, false},
		{"a declaration twice", `<a xmlns:x="u" xmlns:x="v"/>`, false},
		{"an element with an undeclared prefix", `<a><zz:x/></a>`, false},
		{"an attribute with an undeclared prefix", `<a zz:x="1"/>`, false},
		{"a prefix used past its element", `<a><b xmlns:x="u"/><x:c/></a>`, false},
		{"text after the root", `<a/>junk`, false},
		{"two roots", `<a/><b/>`, false},
		{"mismatched tags", `<a></b>`, false},
		{"an element left open", `<a>`, false},
		{"a late declaration", `<a/><?xml version="1.0"?>`, false},
		{"space before the declaration", ` <?xml version="1.0"?><a/>`, false},
		{"an end tag with no start", `</a>`, false},
		{"nothing", ``, false},
	} {
		err := strictXML(strings.NewReader(c.doc))
		if c.ok && err != nil {
			t.Errorf("%s: refused: %v", c.name, err)
		}
		if !c.ok && err == nil {
			t.Errorf("%s: not refused", c.name)
		}
	}
}

// TestRepeatRefuses: the repeat check passes an edit that gives the same
// workbook each time, and refuses one whose repeat gives another; and
// arraysGone counts arrays removed, not moved.
func TestRepeatRefuses(t *testing.T) {
	pool := loadPool(t)
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
	if v, err := s.repeat(e, before, after); err != nil || v != nil {
		t.Fatalf("an edit that gives the same workbook: %v, %v", v, err)
	}
	// The first run "saved" another workbook than the repeats do.
	after[0] = pool[1].Data
	v, err := s.repeat(e, before, after)
	if err != nil || v == nil || v.invariant != "repeat" {
		t.Fatalf("a repeat that gives another workbook: %v, %v", v, err)
	}
	t.Log(v.detail)
	if got, _ := os.ReadFile(s.paths[0]); !bytes.Equal(got, after[0]) {
		t.Error("the check did not put the first run's bytes back")
	}

	arrays := func(rs ...grid.Range) *grid.Workbook {
		return &grid.Workbook{Sheets: []*grid.Sheet{{Spills: rs}}}
	}
	a, b, c := grid.Range{R1: 2, C1: 13, R2: 4, C2: 13}, grid.Range{R1: 2, C1: 14, R2: 2, C2: 14}, grid.Range{R1: 5, C1: 13, R2: 7, C2: 13}
	for _, tc := range []struct {
		before, after *grid.Workbook
		want          int
	}{
		{arrays(a, b), arrays(a, b), 0},
		{arrays(a, b), arrays(c, b), 0}, // one moved
		{arrays(a, b, c), arrays(), 3},
		{arrays(a, b), arrays(b), 1},
		{arrays(a), arrays(a, b), 0}, // one added
	} {
		if got := arraysGone(tc.before, tc.after); got != tc.want {
			t.Errorf("arraysGone(%v, %v) = %d, want %d", tc.before.Sheets[0].Spills, tc.after.Sheets[0].Spills, got, tc.want)
		}
	}
}

func TestKnownParses(t *testing.T) {
	text, err := os.ReadFile("known.txt")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ParseKnown(string(text)); err != nil {
		t.Fatal(err)
	}
	if _, err := ParseKnown("move kinds\n"); err == nil {
		t.Error("a line without a seed parsed")
	}
	if _, err := ParseKnown("move typo 3\n"); err == nil {
		t.Error("an unknown invariant parsed")
	}
}

// TestEditsDeterministic: a seed makes the same edits each time.
func TestEditsDeterministic(t *testing.T) {
	pool := loadPool(t)
	dir := t.TempDir()
	var logs [2][]string
	for i := range logs {
		for seed := uint64(7); seed < 10; seed++ {
			_, err := Replay(pool, seed, dir, 4, func(s string) { logs[i] = append(logs[i], s) })
			if err != nil {
				t.Fatal(err)
			}
		}
	}
	if a, b := strings.Join(logs[0], "\n"), strings.Join(logs[1], "\n"); a != b {
		t.Fatalf("two runs differ:\n%s\n---\n%s", a, b)
	}
}

// defaultOptions are the default mode's: a fixed number of edits, and under
// -race (and only then) a time cap.
func defaultOptions(dir string) (Options, string) {
	o := Options{Seed: defaultSeed, Edits: defaultEdits, Dir: dir}
	if raceOn {
		o.Time, o.Min = raceCap, raceMinEdits
		return o, "race"
	}
	return o, "plain"
}

// TestDefaultOptions: a plain run is never cut short by time, which would
// pass it with fewer edits on a slow machine; only -race caps by time, and
// still makes raceMinEdits.
func TestDefaultOptions(t *testing.T) {
	o, mode := defaultOptions(t.TempDir())
	switch {
	case !raceOn && (o.Time != 0 || o.Min != 0 || o.Edits != 2000 || mode != "plain"):
		t.Errorf("a plain run is %d edits with no cap, not %+v", 2000, o)
	case raceOn && (o.Time <= 0 || o.Min < 200 || mode != "race"):
		t.Errorf("a race run is capped by time and makes 200 edits at least, not %+v", o)
	}
}

// TestFuzz is the default mode. SHEETS_FUZZ_REPRO=<seed> replays one
// sequence instead, printing every edit.
func TestFuzz(t *testing.T) {
	t.Setenv("SHEETS_STREAM", "")
	pool := loadPool(t)
	text, err := os.ReadFile("known.txt")
	if err != nil {
		t.Fatal(err)
	}
	known, err := ParseKnown(string(text))
	if err != nil {
		t.Fatal(err)
	}
	if v := os.Getenv("SHEETS_FUZZ_REPRO"); v != "" {
		seed, err := strconv.ParseUint(v, 10, 64)
		if err != nil {
			t.Fatalf("SHEETS_FUZZ_REPRO=%q: %v", v, err)
		}
		f, err := Replay(pool, seed, t.TempDir(), 0, func(s string) { t.Log(s) })
		if err != nil {
			t.Fatal(err)
		}
		if f == nil {
			t.Logf("seed %d: no finding", seed)
			return
		}
		g, err := Shrink(pool, *f, t.TempDir())
		if err != nil {
			t.Fatal(err)
		}
		t.Errorf("%s(known: %v)", g.Report(), known[f.Signature])
		return
	}

	o, mode := defaultOptions(t.TempDir())
	// Another seed or count is not the default mode: known.txt holds the
	// default seed's findings only.
	if v := os.Getenv("SHEETS_FUZZ_SEED"); v != "" {
		if o.Seed, err = strconv.ParseUint(v, 10, 64); err != nil {
			t.Fatalf("SHEETS_FUZZ_SEED=%q: %v", v, err)
		}
		mode += " (SHEETS_FUZZ_SEED)"
	}
	if v := os.Getenv("SHEETS_FUZZ_EDITS"); v != "" {
		if o.Edits, err = strconv.Atoi(v); err != nil {
			t.Fatalf("SHEETS_FUZZ_EDITS=%q: %v", v, err)
		}
		o.Min = min(o.Min, o.Edits)
		mode += " (SHEETS_FUZZ_EDITS)"
	}
	custom := strings.Contains(mode, "(")
	res, err := Run(pool, o)
	if err != nil {
		t.Fatal(err)
	}
	seen := map[Signature]bool{}
	var lines []string
	newFindings := 0
	for _, f := range res.Findings {
		seen[f.Signature] = true
		if known[f.Signature] {
			lines = append(lines, "known "+f.Signature.String())
			continue
		}
		newFindings++
		lines = append(lines, "NEW   "+f.Signature.String())
		// Shrinking replays the edits many times: the first few are shrunk,
		// the rest printed as the run made them.
		g, how := f, "as made"
		if newFindings <= maxShrunk {
			if g, err = Shrink(pool, f, o.Dir); err != nil {
				t.Fatal(err)
			}
			how = "shrunk"
		}
		t.Errorf("a finding not in known.txt, %s (replay: SHEETS_FUZZ_REPRO=%d go test ./internal/fuzz -run 'TestFuzz$' -v):\n%s", how, f.Seed, g.Report())
	}
	sort.Strings(lines)
	t.Logf("fuzz %s: %d edits (%d applied) in %d sequences, %.1fs, seed %d; %d findings, %d new",
		mode, res.Edits, res.Applied, res.Sequences, res.Elapsed.Seconds(), o.Seed, len(res.Findings), newFindings)
	if res.Capped {
		t.Logf("fuzz %s: the time cap (%s) stopped the run after %d of %d edits", mode, o.Time, res.Edits, o.Edits)
	}
	for _, l := range lines {
		t.Log(l)
	}
	var forms []string
	for _, f := range Results {
		forms = append(forms, fmt.Sprintf("%s %d", f, res.Saved[f]))
	}
	t.Logf("fuzz %s: saved files holding a formula result of each form: %s; %d edits removed a dynamic array and were made again", mode, strings.Join(forms, ", "), res.Repeated)
	if !custom {
		// The invariants must have been put to every type of result and to
		// shared formulas, or the pool and the edits have lost them.
		for _, f := range []string{resNumber, resText, resEmptyText, resBoolean, resShared} {
			if res.Saved[f] == 0 {
				t.Errorf("no file the editors saved held a %s: the pool and the edits no longer put the invariants to it", f)
			}
		}
	}
	streamTotal, takenTotal := 0, 0
	var by []string
	for _, c := range Classes {
		streamTotal += res.Streamed[c.Name]
		takenTotal += res.Taken[c.Name]
		if res.Streamed[c.Name] > 0 {
			by = append(by, fmt.Sprintf("%s %d of %d", c.Name, res.Streamed[c.Name], res.Taken[c.Name]))
		}
	}
	t.Logf("fuzz %s: the streaming editor could take %d of %d edits itself (%s); the others are full-editor code in both copies", mode, streamTotal, takenTotal, strings.Join(by, ", "))
	if !custom && !raceOn && streamTotal < minStreamed {
		t.Errorf("only %d edits could be taken by the streaming editor, want %d: the pool or the edits hand too much to the full editor (see README, The streaming editor)", streamTotal, minStreamed)
	}
	var stale []string
	for s := range known {
		if !seen[s] {
			stale = append(stale, s.String())
		}
	}
	sort.Strings(stale)
	if len(stale) > 0 && !res.Capped && !custom {
		t.Logf("known.txt lists %d findings this run did not make (fixed? then remove them): %s", len(stale), strings.Join(stale, "; "))
	}
	if !custom && !raceOn && res.Edits < defaultEdits {
		t.Errorf("made %d edits, want %d", res.Edits, defaultEdits)
	}
	if !custom && raceOn && res.Edits < raceMinEdits {
		t.Errorf("made %d edits under -race, want at least %d", res.Edits, raceMinEdits)
	}
}

func mustZip(t *testing.T, data []byte) *zip.Reader {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	return zr
}
