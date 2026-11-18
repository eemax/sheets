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

	"github.com/eemax/sheets/internal/undo"
)

// The default mode, in every go test ./...: a fixed seed, a fixed number
// of edits, and a time cap that matters only under -race. See README.md.
const (
	defaultSeed  = 1
	defaultEdits = 2000
	// Under -race each edit takes about ten times as long: the run stops
	// at the cap, once raceMinEdits are made.
	raceCap      = 40 * time.Second
	raceMinEdits = 200
	// Without -race the run makes all its edits well inside the cap; the
	// cap guards a slow machine, and the run says when it cut.
	plainCap = 90 * time.Second
	// New findings past this many are printed unshrunk.
	maxShrunk = 5
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

	o := Options{Seed: defaultSeed, Edits: defaultEdits, Time: plainCap, Dir: t.TempDir()}
	mode := "plain"
	if raceOn {
		o.Time, o.Min, mode = raceCap, raceMinEdits, "race"
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
	if !custom && !raceOn && !res.Capped && res.Edits < defaultEdits {
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
