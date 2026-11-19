package fuzz

import (
	"archive/zip"
	"bytes"
	"fmt"
	"math/rand/v2"
	"os"
	"path/filepath"
	"slices"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/undo"
	"github.com/eemax/sheets/internal/words"
	"github.com/eemax/sheets/internal/xlsx"
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
	// A plain run must put the streaming editor's own text writer at least this
	// many texts of each kind that needs escaping (see escapeKinds), and make
	// the sorts it judges: at least minSorts applied, and at least as many
	// applied as refused. (An earlier version's sorts were refused by both
	// editors 97 times in 137.)
	minEscaped = 5
	minSorts   = 60
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
		// The file's validations, conditional formats and looks cost 2 ms an
		// edit read, and 40 ms under -race, which makes at least 200 edits:
		// check full would pass 135 s (148 s measured). The plain run and long
		// mode judge them.
		fileChecks = false
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
// sequence instead, printing every edit; SHEETS_FUZZ=long runs long mode (see
// fuzzLong).
func TestFuzz(t *testing.T) {
	t.Setenv("SHEETS_STREAM", "")
	pool := loadPool(t)
	text, err := os.ReadFile("known.txt")
	if err != nil {
		t.Fatal(err)
	}
	kn, err := ReadKnown(string(text))
	if err != nil {
		t.Fatal(err)
	}
	known := kn.Seeds
	long := false
	switch v := os.Getenv("SHEETS_FUZZ"); v {
	case "":
	case "long":
		long = true
	default:
		t.Fatalf("SHEETS_FUZZ=%q: the one mode it names is long", v)
	}
	if v := os.Getenv("SHEETS_FUZZ_REPRO"); v != "" {
		seed, err := strconv.ParseUint(v, 10, 64)
		if err != nil {
			t.Fatalf("SHEETS_FUZZ_REPRO=%q: %v", v, err)
		}
		files := pool
		var p uint64
		if long {
			// The sequence's pool, made again from its seed.
			ps := &Pools{Root: repoRoot, Dir: t.TempDir(), Committed: pool}
			p = PoolSeed(seed)
			if files = ps.Files(p); ps.Failed != "" {
				t.Fatalf("cannot make the pool of sequence %d: %s", seed, ps.Failed)
			}
			t.Logf("sequence %d starts from pool %d: %s", seed, p, GenerateCommand(p, "DIR"))
		}
		f, err := Replay(files, seed, t.TempDir(), 0, func(s string) { t.Log(s) })
		if err != nil {
			t.Fatal(err)
		}
		if f == nil {
			t.Logf("seed %d: no finding", seed)
			return
		}
		f.Pool, f.Long = p, p != 0
		g, err := Shrink(files, *f, t.TempDir())
		if err != nil {
			t.Fatal(err)
		}
		how := "not listed by its seed"
		if known[f.Signature] {
			how = "listed by its seed"
		}
		if b := kn.Bug(*f); b != nil {
			how += "; it shows the bug " + b.Name + " (known.txt:" + strconv.Itoa(b.Line) + ")"
		}
		t.Errorf("%s(known.txt: %s)", g.Report(), how)
		return
	}
	if long {
		fuzzLong(t, pool, kn)
		return
	}

	o, mode := defaultOptions(t.TempDir())
	// SHEETS_FUZZ_UNTIL=new: the mutants, on the committed pool when long
	// mode's cannot be made (scripts/fuzz-mutants.sh).
	until := os.Getenv("SHEETS_FUZZ_UNTIL")
	if until != "" && until != "new" {
		t.Fatalf("SHEETS_FUZZ_UNTIL=%q: the one it takes is new", until)
	}
	var theirs map[Signature]string
	o.Stop, theirs = untilNew(t, kn, until == "new")
	// What a mutant's run takes as known: what known.txt lists, a bug it
	// names, and what the tree without the mutant makes too.
	isKnown := func(f Finding) bool { return known[f.Signature] }
	if o.Stop != nil {
		mode += " (SHEETS_FUZZ_UNTIL)"
		isKnown = func(f Finding) bool { return known[f.Signature] || kn.Bug(f) != nil || theirs[f.Signature] != "" }
	}
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
	newFindings := judgeFindings(t, res, isKnown, func(uint64) []Workbook { return pool }, o.Dir, "")
	t.Logf("fuzz %s: %d edits (%d applied) in %d sequences, %.1fs, seed %d; %d findings, %d new",
		mode, res.Edits, res.Applied, res.Sequences, res.Elapsed.Seconds(), o.Seed, len(res.Findings), newFindings)
	if res.Capped {
		t.Logf("fuzz %s: the time cap (%s) stopped the run after %d of %d edits", mode, o.Time, res.Edits, o.Edits)
	}
	logFindings(t, res, func(f Finding) string {
		if isKnown(f) {
			return "known"
		}
		return ""
	})
	logStop(t, mode, res, theirs)
	if res.Hung {
		return // the run stopped at the edit that hung: the counts below are no judge
	}
	logCounts(t, mode, res, !custom && !raceOn)
	if !custom {
		// The invariants must have been put to every type of result and to
		// shared formulas, or the pool and the edits have lost them.
		for _, f := range []string{resNumber, resText, resEmptyText, resBoolean, resShared} {
			if res.Saved[f] == 0 {
				t.Errorf("no file the editors saved held a %s: the pool and the edits no longer put the invariants to it", f)
			}
		}
	}
	if !raceOn {
		// Under -race the repeats are 3, and may miss what 20 see.
		logFixed(t, known, res, o.Seed, false)
	}
	if !custom && !raceOn && res.Edits < defaultEdits {
		t.Errorf("made %d edits, want %d", res.Edits, defaultEdits)
	}
	if !custom && raceOn && res.Edits < raceMinEdits {
		t.Errorf("made %d edits under -race, want at least %d", res.Edits, raceMinEdits)
	}
}

// repoRoot is the repository, from the package's directory (go test runs a
// package's tests there).
const repoRoot = "../.."

// judgeFindings fails the test on each finding known.txt does not list
// (known tells), shrinking the first maxShrunk (pool gives the files of a
// pool, by seed), and returns how many there are. env goes before the line
// that replays one (long mode's SHEETS_FUZZ=long).
func judgeFindings(t *testing.T, res *Result, known func(Finding) bool, pool func(uint64) []Workbook, dir, env string) int {
	t.Helper()
	newFindings := 0
	for _, f := range res.Findings {
		if known(f) {
			continue
		}
		newFindings++
		// Shrinking replays the edits many times: the first few are shrunk,
		// the rest printed as the run made them.
		g, how := f, "as made"
		if f.Invariant == "hang" {
			how = "as made: a hang is not shrunk, and the run stopped there"
		} else if newFindings <= maxShrunk {
			var err error
			if g, err = Shrink(pool(f.Pool), f, dir); err != nil {
				t.Fatal(err)
			}
			how = "shrunk"
		}
		t.Errorf("a finding not in known.txt, %s, after %.0fs (replay: %sSHEETS_FUZZ_REPRO=%d go test ./internal/fuzz -run 'TestFuzz$' -v):\n%s", how, f.After.Seconds(), env, f.Seed, g.Report())
	}
	return newFindings
}

// logFindings logs each finding's signature, known or new, in order: known
// says how a finding is known ("known", "known: the bug escape"), "" when it
// is new.
func logFindings(t *testing.T, res *Result, known func(Finding) string) {
	t.Helper()
	var lines []string
	for _, f := range res.Findings {
		if how := known(f); how != "" {
			lines = append(lines, "known "+f.Signature.String()+strings.TrimPrefix(how, "known"))
		} else {
			lines = append(lines, "NEW   "+f.Signature.String())
		}
	}
	sort.Strings(lines)
	for _, l := range lines {
		t.Log(l)
	}
}

// logCounts logs what the run put the invariants to, and fails a run held to
// the minimums (strict) that falls short of one.
func logCounts(t *testing.T, mode string, res *Result, strict bool) {
	t.Helper()
	var forms []string
	for _, f := range Results {
		forms = append(forms, fmt.Sprintf("%s %d", f, res.Saved[f]))
	}
	t.Logf("fuzz %s: saved files holding a formula result of each form: %s; %d edits removed a dynamic array and were made again", mode, strings.Join(forms, ", "), res.Repeated)
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
	if strict && streamTotal < minStreamed {
		t.Errorf("only %d edits could be taken by the streaming editor, want %d: the pool or the edits hand too much to the full editor (see README, The streaming editor)", streamTotal, minStreamed)
	}
	var escaped []string
	for _, k := range EscapeKinds {
		escaped = append(escaped, fmt.Sprintf("%s %d", k, res.Escaped[k]))
		if strict && res.Escaped[k] < minEscaped {
			t.Errorf("the streaming editor wrote only %d texts with %s in them, want %d: the edits no longer put its text writer to what it must escape", res.Escaped[k], k, minEscaped)
		}
	}
	t.Logf("fuzz %s: texts the streaming editor could write itself, by what they hold: %s", mode, strings.Join(escaped, ", "))
	t.Logf("fuzz %s: sorts: %d made, %d taken by both editors, %d refused by both", mode, res.Made["sort"], res.Taken["sort"], res.Made["sort"]-res.Taken["sort"])
	if strict && (res.Taken["sort"] < minSorts || res.Taken["sort"] < res.Made["sort"]-res.Taken["sort"]) {
		t.Errorf("of %d sorts, %d were taken by both editors, want at least %d and at least as many as were refused: the sort class is barely judged (see sortable in edits.go)", res.Made["sort"], res.Taken["sort"], minSorts)
	}
}

// fixed are the entries of known.txt for a run's seed (long ones or not)
// whose sequence the run made in full, without making the finding: fixed,
// or no longer reached. They are reported, and never fail a run. A long
// entry whose sequence started from the committed pool (its own could not be
// made: res.Fallback) was not put to the test, and is not one.
func fixed(known map[Signature]bool, res *Result, seed uint64, long bool) []string {
	seen := map[Signature]bool{}
	for _, f := range res.Findings {
		seen[f.Signature] = true
	}
	var out []string
	for s := range known {
		if s.Long != long || s.Seed/100000 != seed || seen[s] {
			continue
		}
		i := int(s.Seed % 100000)
		if long && res.Fallback[i] {
			continue
		}
		if i < res.Sequences-1 || i == res.Sequences-1 && !res.Truncated {
			out = append(out, s.String())
		}
	}
	sort.Strings(out)
	return out
}

func logFixed(t *testing.T, known map[Signature]bool, res *Result, seed uint64, long bool) {
	t.Helper()
	if fx := fixed(known, res, seed, long); len(fx) > 0 {
		t.Logf("fixed: known.txt lists %d findings whose sequences this run made in full without them (fixed? then remove them): %s", len(fx), strings.Join(fx, "; "))
	}
}

func readWorkbook(data []byte) (*grid.Workbook, error) { return xlsx.ReadBytes(data) }

func mustZip(t *testing.T, data []byte) *zip.Reader {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	return zr
}

// TestStreamed: the estimate of what the streaming editor takes itself, and
// the workbooks it takes the edits on.
func TestStreamed(t *testing.T) {
	var plain, arrays *grid.Workbook
	for _, wb := range loadPool(t) {
		g, err := readWorkbook(wb.Data)
		if err != nil {
			t.Fatal(err)
		}
		if streamable(g) && plain == nil {
			plain = g
		}
		if !streamable(g) && arrays == nil {
			arrays = g
		}
	}
	if plain == nil || arrays == nil {
		t.Fatalf("the pool must hold workbooks the streaming editor takes and others: %v, %v", plain != nil, arrays != nil)
	}
	name := plain.Sheets[0].Name
	for _, tc := range []struct {
		name string
		wb   *grid.Workbook
		e    Edit
		want bool
	}{
		{"set", plain, Edit{Class: "set", Args: []string{"set", FileArg, name + "!B2", "5", name + "!B3", "=IF(N(\"p\")=0,B2*2)"}}, true},
		{"set a dynamic array", plain, Edit{Class: "set", Args: []string{"set", FileArg, name + "!B2", "=IF(N(\"d\")=0,SEQUENCE(2))"}}, false},
		{"set on a new sheet", plain, Edit{Class: "set", Args: []string{"set", FileArg, "Nowhere!B2", "5"}}, false},
		{"write", plain, Edit{Class: "write", Args: []string{"write", FileArg, name + "!B2", "--from", "tsv"}, Stdin: "1\t2\n"}, true},
		{"write a dynamic array", plain, Edit{Class: "write", Args: []string{"write", FileArg, name + "!B2", "--from", "tsv"}, Stdin: "=IF(N(\"d\")=0,SEQUENCE(2))\n"}, false},
		{"rows", plain, Edit{Class: "rows-delete", Args: []string{"rows", "delete", FileArg, name + "!3:4"}}, true},
		{"copy", plain, Edit{Class: "copy", Args: []string{"copy", FileArg, name + "!B2", name + "!B9"}}, false},
		{"set with arrays", arrays, Edit{Class: "set", Args: []string{"set", FileArg, arrays.Sheets[0].Name + "!B2", "5"}}, false},
		{"rows with arrays", arrays, Edit{Class: "rows-insert", Args: []string{"rows", "insert", FileArg, arrays.Sheets[0].Name + "!3"}}, false},
	} {
		if got := streamed(tc.e, tc.wb); got != tc.want {
			t.Errorf("%s: streamed = %v, want %v", tc.name, got, tc.want)
		}
	}
}

// TestReadArea: a move or copy often starts at cells a formula reads.
func TestReadArea(t *testing.T) {
	s := grid.NewSheet("Data")
	s.Set(2, 9, grid.Cell{Kind: grid.Number, Formula: `IF(N("p")=0,B3:B4*2)`})
	s.Set(10, 7, grid.Cell{Kind: grid.Number, Formula: `IF(N("p")=0,SUM(B2:B9))`})
	other := grid.NewSheet("Plan 2")
	other.Set(2, 2, grid.Cell{Kind: grid.Number, Formula: `IF(N("p")=0,'Data'!C2:C3*10)`})
	g := &gen{rng: rand.New(rand.NewPCG(3, 4)), wb: &grid.Workbook{Sheets: []*grid.Sheet{s, other}}}
	seen := map[string]int{}
	for range 400 {
		sh, rg, ok := g.readArea()
		if !ok {
			t.Fatal("no cells read")
		}
		seen[sh.Name+"!"+rg.String()]++
		in := func(name string, outer grid.Range) bool {
			return sh.Name == name && rg.R1 >= outer.R1 && rg.R2 <= outer.R2 && rg.C1 >= outer.C1 && rg.C2 <= outer.C2
		}
		if !in("Data", grid.Range{R1: 3, C1: 2, R2: 4, C2: 2}) && !in("Data", grid.Range{R1: 2, C1: 2, R2: 9, C2: 2}) && !in("Data", grid.Range{R1: 2, C1: 3, R2: 3, C2: 3}) {
			t.Fatalf("%s!%s is not read by any formula", sh.Name, rg)
		}
	}
	for _, want := range []string{"Data!B3:B4", "Data!C2:C3", "Data!B2:B9"} {
		if seen[want] == 0 {
			t.Errorf("%s, the range a formula reads, was never picked in 400: %v", want, seen)
		}
	}
	if g.wb.Sheets[0].Name != "Data" {
		t.Error("the workbook changed")
	}
	if seen["Data!B3:B4"]+seen["Data!C2:C3"] < seen["Data!B2:B9"] {
		t.Errorf("formulas that read a range where one value goes should be picked first: %v", seen)
	}
}

// TestBatchesEditNewSheets: a do batch that makes a sheet goes on to edit it
// (the lines are made from the sheets the batch makes, not only from the
// workbook before it).
func TestBatchesEditNewSheets(t *testing.T) {
	var wb *grid.Workbook
	for _, w := range loadPool(t) {
		g, err := readWorkbook(w.Data)
		if err != nil {
			t.Fatal(err)
		}
		wb = g
		break
	}
	rng := rand.New(rand.NewPCG(11, 12))
	made, edited := 0, 0
	for range 300 {
		cp := *wb
		cp.Sheets = slices.Clone(wb.Sheets)
		g := &gen{rng: rng, wb: &cp}
		e := g.edit("do")
		lines := strings.Split(strings.TrimSuffix(e.Stdin, "\n"), "\n")
		first, err := words.Texts(lines[0])
		if err != nil {
			t.Fatal(err)
		}
		// The sheet the first line makes: sheet add NAME, sheet copy FROM NAME,
		// or a set or copy to a name the workbook does not have.
		name := ""
		switch {
		case len(first) == 3 && first[0] == "sheet" && first[1] == "add":
			name = first[2]
		case len(first) == 4 && first[0] == "sheet" && first[1] == "copy":
			name = first[3]
		case first[0] == "set" || first[0] == "copy" || first[0] == "move":
			for _, a := range first[1:] {
				if i := strings.LastIndex(a, "!"); i > 0 && strings.HasSuffix(a[:i], "'") == strings.HasPrefix(a, "'") && wb.Sheet(strings.Trim(a[:i], "'")) == nil {
					name = strings.Trim(a[:i], "'")
				}
			}
		}
		if name == "" {
			continue
		}
		made++
		for _, l := range lines[1:] {
			ws, err := words.Texts(l)
			if err != nil {
				t.Fatal(err)
			}
			if slices.ContainsFunc(ws, func(a string) bool { return a == name || strings.HasPrefix(a, name+"!") }) {
				edited++
				break
			}
		}
	}
	if made < 60 || edited < made/2 {
		t.Errorf("of 300 batches, %d make a sheet, and %d of them go on to edit it", made, edited)
	}
	t.Logf("of 300 batches, %d make a sheet, and %d of them go on to edit it", made, edited)
}
