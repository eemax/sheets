package fuzz

import (
	"fmt"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"slices"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/grid"
)

// maxSamples is the most workbooks long mode sends LibreOffice, spread over
// the run.
const maxSamples = 120

// fuzzLong is long mode (SHEETS_FUZZ=long): a run for a time, over pools it
// makes fresh with scripts/fuzzgen.py (one for each round of RoundSequences
// sequences, from seeds the run's seed gives), so its sequences start from
// many more workbooks than the committed 40. See README.md, Long mode.
//
//   - SHEETS_FUZZ_SEED: the run's seed (default: a fresh one, printed, from
//     ObjectsFromRun up: its pools hold objects beside cells);
//   - SHEETS_FUZZ_TIME: how long it runs (default 10m), finishing the
//     sequence it is in;
//   - SHEETS_FUZZ_EDITS: the most edits it makes (default: no limit);
//   - SHEETS_FUZZ_UNTIL=new: stop at the first finding known.txt does not
//     list (scripts/fuzz-mutants.sh), and with SHEETS_FUZZ_BASE, a copy of
//     the tree without the mutant with its test binary built
//     (internal/fuzz/fuzz.test), only at one whose sequence the copy makes
//     with no finding: the mutant's;
//   - SHEETS_FUZZ_SOFFICE: LibreOffice's soffice, when it is not on PATH.
//
// A finding is known when known.txt lists its signature or it shows one of
// the bugs known.txt names (a fresh seed's findings have signatures no list
// holds: the bugs are known by their face). One that is not fails the run, as
// in the default mode. The log gives the rate of findings per 10,000 edits,
// the findings by the bug they show, the bugs whose repro no longer shows
// them and the entries for the run's seed no longer made (fixed: reported,
// never a failure), and LibreOffice's second opinion on the values of plain
// formulas (reported, never a failure).
func fuzzLong(t *testing.T, committed []Workbook, kn *Known) {
	seed := ObjectsFromRun + uint64(time.Now().UnixNano()/int64(time.Millisecond)%1_000_000) + 2
	given := false
	if v := os.Getenv("SHEETS_FUZZ_SEED"); v != "" {
		var err error
		if seed, err = strconv.ParseUint(v, 10, 64); err != nil || seed == 0 {
			t.Fatalf("SHEETS_FUZZ_SEED=%q: want a seed above 0 (%v)", v, err)
		}
		given = true
	}
	dur := 10 * time.Minute
	if v := os.Getenv("SHEETS_FUZZ_TIME"); v != "" {
		var err error
		if dur, err = time.ParseDuration(v); err != nil || dur <= 0 {
			t.Fatalf("SHEETS_FUZZ_TIME=%q: want a duration, like 20m (%v)", v, err)
		}
	}
	edits := math.MaxInt
	if v := os.Getenv("SHEETS_FUZZ_EDITS"); v != "" {
		var err error
		if edits, err = strconv.Atoi(v); err != nil {
			t.Fatalf("SHEETS_FUZZ_EDITS=%q: %v", v, err)
		}
	}
	until := false
	switch v := os.Getenv("SHEETS_FUZZ_UNTIL"); v {
	case "":
	case "new":
		until = true
	default:
		t.Fatalf("SHEETS_FUZZ_UNTIL=%q: the one it takes is new", v)
	}
	dir := t.TempDir()
	ps := &Pools{Root: repoRoot, Dir: dir, Committed: committed, Run: seed}
	o := Options{Seed: seed, Edits: edits, Time: dur, Dir: dir, Pools: ps.For}
	// known: listed by its signature, or a known bug's.
	known := func(f Finding) bool { return kn.Seeds[f.Signature] || kn.Bug(f) != nil }
	var theirs map[Signature]string
	o.Stop, theirs = untilNew(t, kn, until)
	soffice := os.Getenv("SHEETS_FUZZ_SOFFICE")
	if soffice == "" {
		soffice, _ = exec.LookPath("soffice")
	}
	var samples []Sample
	if soffice != "" && !until {
		// Spread over the run: one in each maxSamples-th of its time, at most.
		every := dur / maxSamples
		start, next := time.Now(), time.Duration(0)
		o.Sample = func(seq uint64, p Pool, wb string, data []byte) {
			if p.Seed == 0 || !strings.HasPrefix(wb, "openpyxl") || len(samples) >= maxSamples || time.Since(start) < next {
				return
			}
			next = time.Since(start) + every
			samples = append(samples, Sample{seq, p.Seed, wb, data})
		}
	}
	how := "given"
	if !given {
		how = "fresh"
	}
	t.Logf("fuzz long: seed %d (%s), for %s; a fresh pool of %d workbooks each %d sequences, pool seed = seed*1000 + round: %s",
		seed, how, dur, PoolFiles, RoundSequences, GenerateCommand(PoolSeed(SequenceSeed(seed, 0)), "DIR"))
	res, err := Run(committed, o)
	if err != nil {
		t.Fatal(err)
	}
	newFindings := judgeFindings(t, res, func(f Finding) bool { return known(f) || theirs[f.Signature] != "" }, ps.Files, dir, "SHEETS_FUZZ=long ")
	t.Logf("fuzz long: %d edits (%d applied) in %d sequences, %.0fs, seed %d; %d pools generated (seeds %s); %d findings, %d new",
		res.Edits, res.Applied, res.Sequences, res.Elapsed.Seconds(), seed, len(res.Pools), seedList(res.Pools), len(res.Findings), newFindings)
	if ps.Failed != "" {
		t.Logf("fuzz long: could not generate fresh pools, so the committed pool ran in their place for %d of %d sequences (their findings are not long ones): %s", len(res.Fallback), res.Sequences, ps.Failed)
	}
	t.Logf("fuzz long: %.1f findings per 10,000 edits (%d in %d edits): %s", float64(len(res.Findings))*10000/float64(max(res.Edits, 1)), len(res.Findings), res.Edits, bySignature(res.Findings))
	t.Logf("fuzz long: findings by the bug known.txt names that they show: %s", byBug(kn, res.Findings))
	logFindings(t, res, func(f Finding) string {
		switch b := kn.Bug(f); {
		case kn.Seeds[f.Signature]:
			return "known"
		case b != nil:
			return "known: the bug " + b.Name
		case theirs[f.Signature] != "":
			return "known: not the mutant's"
		}
		return ""
	})
	logStop(t, "long", res, theirs)
	if res.Hung {
		return
	}
	logCounts(t, "long", res, false)
	logFixed(t, kn.Seeds, res, seed, true)
	if until {
		return
	}
	logBugsFixed(t, kn, committed, ps, dir)
	if soffice == "" {
		t.Logf("fuzz long: LibreOffice's second opinion: none (no soffice on PATH; SHEETS_FUZZ_SOFFICE names it)")
		return
	}
	if len(samples) == 0 && len(res.Pools) == 0 {
		t.Logf("fuzz long: LibreOffice's second opinion: none (no fresh pool was made, and the committed pool's workbooks are not sent)")
		return
	}
	op := SecondOpinion(soffice, samples, filepath.Join(dir, "lo"))
	t.Logf("fuzz long: LibreOffice's second opinion on plain formulas, in workbooks from openpyxl files (reported, never a finding): %s", op)
}

// untilNew is the stop of a run with SHEETS_FUZZ_UNTIL=new (until; nil
// without): the first finding known.txt does not list, by its signature or
// as a bug's (the mutants: scripts/fuzz-mutants.sh). With SHEETS_FUZZ_BASE, a
// copy of the tree without the mutant whose test binary is built
// (internal/fuzz/fuzz.test), it is the first whose sequence the copy makes
// with no finding (fuzzer-long's audit, f0.3: a finding the code makes anyway
// was taken for the mutant's); the others are in theirs, with what the copy
// made.
func untilNew(t *testing.T, kn *Known, until bool) (stop func(Finding) bool, theirs map[Signature]string) {
	t.Helper()
	base := os.Getenv("SHEETS_FUZZ_BASE")
	theirs = map[Signature]string{}
	switch {
	case base != "" && !until:
		t.Fatalf("SHEETS_FUZZ_BASE=%s: only with SHEETS_FUZZ_UNTIL=new", base)
	case !until:
		return nil, theirs
	}
	return func(f Finding) bool {
		if kn.Seeds[f.Signature] || kn.Bug(f) != nil {
			return false
		}
		if base == "" {
			return true
		}
		made, err := baseMakes(base, f)
		if err != nil {
			t.Fatalf("SHEETS_FUZZ_BASE=%s: %v", base, err)
		}
		if made != "" {
			theirs[f.Signature] = made
			return false
		}
		return true
	}, theirs
}

// logStop logs the findings the tree without the mutant makes too (theirs)
// and the finding a run with SHEETS_FUZZ_UNTIL=new stopped at: the mutant's
// (scripts/fuzz-mutants.sh reads the line).
func logStop(t *testing.T, mode string, res *Result, theirs map[Signature]string) {
	t.Helper()
	base := os.Getenv("SHEETS_FUZZ_BASE")
	env := ""
	if mode == "long" {
		env = "SHEETS_FUZZ=long "
	}
	for _, f := range res.Findings {
		if made := theirs[f.Signature]; made != "" {
			t.Logf("fuzz %s: %s is not in known.txt, and the tree without the mutant makes its sequence with a finding too (SHEETS_FUZZ_BASE): not the mutant's, but a bug of the code, for known.txt (replay: %sSHEETS_FUZZ_REPRO=%d go test ./internal/fuzz -run 'TestFuzz$' -v):\n%s", mode, f.Signature, env, f.Seed, made)
		}
	}
	if !res.Stopped {
		return
	}
	f := res.Findings[len(res.Findings)-1]
	if base == "" {
		t.Logf("fuzz %s: stopped at the first finding known.txt does not list, %s, after %.0fs (SHEETS_FUZZ_UNTIL=new)", mode, f.Signature, f.After.Seconds())
		return
	}
	t.Logf("fuzz %s: the mutant's finding: %s, after %.0fs: known.txt does not list it, and the tree without the mutant makes its sequence with no finding (SHEETS_FUZZ_UNTIL=new, SHEETS_FUZZ_BASE)", mode, f.Signature, f.After.Seconds())
}

// baseMakes replays a finding's sequence with the tree without the mutant
// (base, its test binary built in internal/fuzz/fuzz.test), and says what
// it makes there: "" when nothing, else the replay's finding.
func baseMakes(base string, f Finding) (string, error) {
	pkg := filepath.Join(base, "internal", "fuzz")
	bin := filepath.Join(pkg, "fuzz.test")
	if _, err := os.Stat(bin); err != nil {
		return "", fmt.Errorf("no test binary of the tree without the mutant (go test -c -o internal/fuzz/fuzz.test ./internal/fuzz): %v", err)
	}
	cmd := exec.Command(bin, "-test.run", "^TestFuzz$", "-test.count=1", "-test.v")
	cmd.Dir = pkg
	for _, kv := range os.Environ() {
		if !strings.HasPrefix(kv, "SHEETS_FUZZ") {
			cmd.Env = append(cmd.Env, kv)
		}
	}
	cmd.Env = append(cmd.Env, "SHEETS_FUZZ_REPRO="+strconv.FormatUint(f.Seed, 10))
	if f.Long {
		cmd.Env = append(cmd.Env, "SHEETS_FUZZ=long")
	}
	out, err := cmd.CombinedOutput()
	text := string(out)
	switch {
	case err == nil && strings.Contains(text, fmt.Sprintf("seed %d: no finding", f.Seed)):
		return "", nil
	case err != nil && strings.Contains(text, "finding "):
		i := strings.Index(text, "finding ")
		return lastLines(text[i:], 40), nil
	}
	return "", fmt.Errorf("the replay of %d gave neither a finding nor none (%v): %s", f.Seed, err, lastLines(text, 5))
}

// byBug counts findings by the bug known.txt names that they show, with
// those it lists by signature and those it does not.
func byBug(kn *Known, fs []Finding) string {
	n := map[string]int{}
	for _, f := range fs {
		switch b := kn.Bug(f); {
		case kn.Seeds[f.Signature]:
			n["listed by signature"]++
		case b != nil:
			n[b.Name]++
		default:
			n["none (new)"]++
		}
	}
	var keys []string
	for k := range n {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(i, j int) bool {
		if n[keys[i]] != n[keys[j]] {
			return n[keys[i]] > n[keys[j]]
		}
		return keys[i] < keys[j]
	})
	var out []string
	for _, k := range keys {
		out = append(out, fmt.Sprintf("%s %d", k, n[k]))
	}
	if len(out) == 0 {
		return "none"
	}
	return strings.Join(out, ", ")
}

// replays is how many times a bug's repro is replayed before it is taken
// for fixed: the map-order bug's editors face shows about one replay in two.
const replays = 10

// bugsFixed replays each known bug's repro, and lists those it no longer
// shows (fixed: reported, never a failure), and those whose repro could not
// be replayed (its pool could not be made), each with its own pool's failure.
func bugsFixed(kn *Known, committed []Workbook, ps *Pools, dir string) (fixed, unchecked []string, err error) {
	for _, b := range kn.Bugs {
		files := committed
		if b.Repro.Long {
			seed := PoolSeed(b.Repro.Seed)
			p := ps.Of(seed)
			if p.Seed == 0 {
				unchecked = append(unchecked, fmt.Sprintf("%s (known.txt:%d, repro long %d, pool %d: %s)", b.Name, b.Line, b.Repro.Seed, seed, ps.Why(seed)))
				continue
			}
			files = p.Files
		}
		// A bug that comes and goes (map order) is replayed until it shows,
		// up to replays times, by any of its faces: map-order's script paste
		// shows its editors face or its kinds face, as the editors' layouts
		// fall, from one repro.
		shows := func(f Finding) bool {
			g := kn.Bug(f)
			return b.Shows(f) || g != nil && g.Name == b.Name
		}
		var f *Finding
		for range replays {
			var err error
			if f, err = Replay(files, b.Repro.Seed, dir, 0, nil); err != nil {
				return nil, nil, err
			}
			if f != nil {
				f.Long = b.Repro.Long
				if shows(*f) {
					break
				}
			}
		}
		switch {
		case f == nil:
			fixed = append(fixed, fmt.Sprintf("%s (known.txt:%d): its repro, %s, makes no finding", b.Name, b.Line, repro(b)))
		case !shows(*f):
			detail, _, _ := strings.Cut(f.Detail, "\n")
			fixed = append(fixed, fmt.Sprintf("%s (known.txt:%d): its repro, %s, makes %s, which it does not show: %s", b.Name, b.Line, repro(b), f.Signature, detail))
		}
	}
	return fixed, unchecked, nil
}

func repro(b Bug) string {
	if b.Repro.Long {
		return fmt.Sprintf("long %d", b.Repro.Seed)
	}
	return strconv.FormatUint(b.Repro.Seed, 10)
}

func logBugsFixed(t *testing.T, kn *Known, committed []Workbook, ps *Pools, dir string) {
	t.Helper()
	fx, unchecked, err := bugsFixed(kn, committed, ps, dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(fx) > 0 {
		t.Logf("fixed: known.txt names %d bugs whose repro no longer shows them (fixed? then remove their lines): %s", len(fx), strings.Join(fx, "; "))
	}
	t.Logf("fuzz long: of known.txt's %d bug lines, %d repros replayed, %d still showing their bug", len(kn.Bugs), len(kn.Bugs)-len(unchecked), len(kn.Bugs)-len(unchecked)-len(fx))
	if len(unchecked) > 0 {
		t.Logf("fuzz long: %d repros could not be replayed, their pools not made: %s", len(unchecked), strings.Join(unchecked, "; "))
	}
}

func seedList(seeds []uint64) string {
	if len(seeds) == 0 {
		return "none"
	}
	if len(seeds) <= 3 {
		var s []string
		for _, x := range seeds {
			s = append(s, strconv.FormatUint(x, 10))
		}
		return strings.Join(s, ", ")
	}
	return fmt.Sprintf("%d to %d", seeds[0], seeds[len(seeds)-1])
}

// bySignature counts findings by edit class and invariant.
func bySignature(fs []Finding) string {
	n := map[string]int{}
	for _, f := range fs {
		n[f.Class+" "+f.Invariant]++
	}
	var keys []string
	for k := range n {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(i, j int) bool {
		if n[keys[i]] != n[keys[j]] {
			return n[keys[i]] > n[keys[j]]
		}
		return keys[i] < keys[j]
	})
	var out []string
	for _, k := range keys {
		out = append(out, fmt.Sprintf("%s %d", k, n[k]))
	}
	if len(out) == 0 {
		return "none"
	}
	return strings.Join(out, ", ")
}

// TestPoolSeed: a long run's sequences start from a pool each round, by a
// seed their own gives.
func TestPoolSeed(t *testing.T) {
	for _, tc := range []struct{ seq, want uint64 }{
		{700000, 7000}, {700012, 7000}, {700000 + RoundSequences - 1, 7000}, {700000 + RoundSequences, 7001}, {100000, 1000},
	} {
		if got := PoolSeed(tc.seq); got != tc.want {
			t.Errorf("PoolSeed(%d) = %d, want %d", tc.seq, got, tc.want)
		}
	}
	if !strings.Contains(GenerateCommand(7000, "DIR"), "fuzzgen.py --seed 7000 --count 40 --pool-only DIR") {
		t.Error(GenerateCommand(7000, "DIR"))
	}
	// Pools of runs from ObjectsFromRun hold objects beside cells; those of
	// every run before it are made as they were, so known.txt's repros replay.
	for _, tc := range []struct {
		pool uint64
		want bool
	}{
		{7000, false}, {718113007, false}, {915488000, false}, {ObjectsFromRun*1000 - 1, false}, {ObjectsFromRun * 1000, true}, {2000031000, true},
	} {
		if got := WithObjects(tc.pool); got != tc.want {
			t.Errorf("WithObjects(%d) = %v, want %v", tc.pool, got, tc.want)
		}
	}
	if c := GenerateCommand(2000031000, "DIR"); !strings.Contains(c, "--seed 2000031000 --count 40 --pool-only --objects DIR") {
		t.Error(c)
	}
}

// TestKnownLong: known.txt lists long mode's findings apart from the default
// mode's: the same class, invariant and seed on another pool.
func TestKnownLong(t *testing.T) {
	known, err := ParseKnown("long move kinds 700012\nmove kinds 700012 # the committed pool\nlong set text 700001 700005\n")
	if err != nil {
		t.Fatal(err)
	}
	long, plain := Signature{"move", "kinds", 700012, true}, Signature{"move", "kinds", 700012, false}
	if !known[long] || !known[plain] || !known[Signature{"set", "text", 700001, true}] || !known[Signature{"set", "text", 700005, true}] || len(known) != 4 {
		t.Errorf("parsed %v", known)
	}
	if long.String() != "long move kinds 700012" || plain.String() != "move kinds 700012" {
		t.Errorf("%s, %s", long, plain)
	}
	for _, bad := range []string{"long move kinds\n", "long move kinds 7 x\n", "long long move kinds 7\n"} {
		if _, err := ParseKnown(bad); err == nil {
			t.Errorf("%q parsed", bad)
		}
	}
}

// TestFixed: an entry of known.txt is reported fixed when the run made its
// sequence in full without it; never one of another seed or mode, one the run
// did not reach, or one in the last sequence when the edit count cut it short.
func TestFixed(t *testing.T) {
	known := map[Signature]bool{
		{"move", "kinds", 700003, true}:   true, // made: not fixed
		{"copy", "kinds", 700004, true}:   true, // not made: fixed
		{"set", "reach", 700009, true}:    true, // the last sequence, cut short
		{"set", "reach", 700020, true}:    true, // not reached
		{"set", "reach", 700004, false}:   true, // the committed pool
		{"set", "reach", 800004, true}:    true, // another seed
		{"sort", "editors", 700000, true}: true, // not made: fixed
	}
	res := &Result{Sequences: 10, Truncated: true, Findings: []Finding{{Signature: Signature{"move", "kinds", 700003, true}}}}
	got := strings.Join(fixed(known, res, 7, true), "; ")
	if want := "long copy kinds 700004; long sort editors 700000"; got != want {
		t.Errorf("fixed: %s, want %s", got, want)
	}
	res.Truncated = false
	if got := fixed(known, res, 7, true); len(got) != 3 {
		t.Errorf("with the last sequence made in full: %v", got)
	}
	// A sequence that ran on the committed pool, its own not made (no uv:
	// fuzzer-long's audit, f0.4), did not put its long entry to the test.
	res.Fallback = map[int]bool{4: true, 9: true}
	if got := strings.Join(fixed(known, res, 7, true), "; "); got != "long sort editors 700000" {
		t.Errorf("with sequences 4 and 9 on the committed pool: %s", got)
	}
}

// TestKnownBugs: a bug line of known.txt names a bug by its findings' face,
// and takes a finding that shows it, whatever its seed; not one of another
// class, invariant or detail, or whose last edit is not the one it names.
func TestKnownBugs(t *testing.T) {
	kn, err := ReadKnown(`move kinds 100003 # a comment
bug sort-editors: sort,do editors; repro 100047; detail : stream holds .+, full holds (nothing \(pending\)|number 0)$; edit (?m)^sort |^sheets sort
bug broken: rows-delete kinds; repro long 700012; detail holds =#REF! # not a comment
`)
	if err != nil {
		t.Fatal(err)
	}
	if !kn.Seeds[Signature{"move", "kinds", 100003, false}] || len(kn.Seeds) != 1 || len(kn.Bugs) != 2 {
		t.Fatalf("seeds %v, bugs %v", kn.Seeds, kn.Bugs)
	}
	b := kn.Bugs[0]
	if b.Name != "sort-editors" || b.Invariant != "editors" || strings.Join(b.Classes, ",") != "sort,do" || b.Repro != (Signature{"sort", "editors", 100047, false}) || b.Line != 2 {
		t.Errorf("parsed %+v", b)
	}
	if r := kn.Bugs[1]; r.Repro != (Signature{"rows-delete", "kinds", 700012, true}) || r.Detail.String() != "holds =#REF! # not a comment" {
		t.Errorf("parsed %+v (%s)", r, r.Detail)
	}
	sortEdit := Edit{Class: "sort", Args: []string{"sort", FileArg, "Data!A2:C9", "B"}}
	batch := func(lines string) Edit { return Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: lines} }
	for _, tc := range []struct {
		name, class, invariant, detail string
		edit                           Edit
		want                           string
	}{
		{"a sort's", "sort", "editors", "Q 2026!B2: stream holds number 2088, full holds nothing (pending)\nalso editors: …", sortEdit, "sort-editors"},
		{"XlsxWriter's 0", "sort", "editors", "Notes!B2: stream holds number 6, full holds number 0", sortEdit, "sort-editors"},
		{"a batch holding a sort", "do", "editors", "Model!B2: stream holds number 1, full holds number 0", batch("cols delete 'Model!C:C'\nsort 'Model!A2:A6' A A\n"), "sort-editors"},
		{"a batch holding none", "do", "editors", "Model!B2: stream holds number 1, full holds number 0", batch("cols delete 'Model!C:C'\nset Model!A2 sort\n"), ""},
		{"another face", "sort", "editors", "Q 2026!B2: stream holds nothing (pending), full holds number 2088", sortEdit, ""},
		{"another invariant", "sort", "kinds", "Q 2026!B2: stream holds number 2088, full holds nothing (pending)", sortEdit, ""},
		{"another class", "move", "editors", "Q 2026!B2: stream holds number 2088, full holds nothing (pending)", sortEdit, ""},
		{"a rows delete", "rows-delete", "kinds", "Data!I3: holds =#REF! # not a comment", Edit{Class: "rows-delete"}, "broken"},
	} {
		f := Finding{Signature: Signature{tc.class, tc.invariant, 4001600015, true}, Detail: tc.detail, Edits: []Edit{tc.edit}}
		got := ""
		if b := kn.Bug(f); b != nil {
			got = b.Name
		}
		if got != tc.want {
			t.Errorf("%s: the bug %q, want %q", tc.name, got, tc.want)
		}
	}
	for _, bad := range []string{
		"bug x: sort editors; detail a\n",                                                  // no repro
		"bug x: sort editors; repro 7\n",                                                   // no detail
		"bug x: sort editors; repro 7; detail (\n",                                         // not a pattern
		"bug x: sortx editors; repro 7; detail a\n",                                        // no class
		"bug x: sort typo; repro 7; detail a\n",                                            // no invariant
		"bug x: sort editors; repro 7; detail a; why b\n",                                  // no field
		"bug x y: sort editors; repro 7; detail a\n",                                       // a name with a space
		"bug x: sort editors; repro long x; detail a\n",                                    // no seed
		"bug x: sort editors; repro 7; detail a\nbug x: sort editors; repro 8; detail b\n", // a face twice
	} {
		if _, err := ReadKnown(bad); err == nil {
			t.Errorf("%q parsed", bad)
		}
	}
}

// TestKnownFaces: the bugs known.txt names take the findings of each (as long
// mode made them, on seeds 7 and 40016), and not those of others: a known
// bug's face must not hide a new one (fuzzer-long's audit, f0.1: a dynamic
// array set in a batch losing its formula sat among 113 known findings).
func TestKnownFaces(t *testing.T) {
	text, err := os.ReadFile("known.txt")
	if err != nil {
		t.Fatal(err)
	}
	kn, err := ReadKnown(string(text))
	if err != nil {
		t.Fatal(err)
	}
	script := func(code string) Edit {
		return Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Data\");\n" + code}}
	}
	batch := func(lines string) Edit { return Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: lines} }
	plain := func(class string) Edit { return Edit{Class: class} }
	for _, tc := range []struct {
		class, invariant, detail string
		edit                     Edit
		want                     string
	}{
		{"set", "text", `Data!I2: typed the text "_x005F_x0041_", and the cell holds text "_x005FA"`, plain("set"), "escape"},
		{"write", "text", `Bob's Plan!F3: typed the text "_x005F_x0041_", and the cell holds text "_x005FA"`, plain("write"), "escape"},
		{"write", "editors", `Data!L3: stream holds text "_x005F_x0041_-x", full holds text "_x005FA-x"`, plain("write"), "escape"},
		{"script", "editors", `Data!L2: stream holds text "_X005F_X0041_", full holds text "_X005FA"`, script(`ws.getRange("A2").setValue("_x005F_x0041_");`), "escape"},
		{"sort", "editors", `Bob's Plan!B2: stream holds number 0, full holds nothing (pending)`, plain("sort"), "sort-editors"},
		{"sort", "editors", `Data!L2: stream holds text "-x", full holds nothing (pending)`, plain("sort"), "sort-editors"},
		{"sort", "editors", `Summary!B2: stream holds number 6, full holds nothing`, plain("sort"), "sort-editors"},
		{"do", "editors", `Calc!B2: stream holds number 7, full holds number 0`, batch("sort 'Calc!A2:A6' A A\n"), "sort-editors"},
		{"script", "editors", `Data!B2: stream holds boolean true, full holds nothing (pending)`, script(`ws.getRange("A2:C9").getSort().apply([{key: 0, ascending: true}]);`), "sort-editors"},
		{"copy", "repeat", `stream editor: the same edit on the same bytes saved two different workbooks: Data!M3: the first run holds number 86.06, the repeat holds nothing`, plain("copy"), "map-order"},
		{"move", "repeat", `full editor: the same edit on the same bytes saved two different workbooks: Data!M3: the first run has =, the repeat has =IF(N("p")=0,IF(B7>1000,1,""))`, plain("move"), "map-order"},
		{"script", "kinds", `stream editor: sheet1.xml U2: =IF(N("l")=0,UPPER(A2:A4)) is stored as a dynamic array, not a legacy array`, script(`ws.getRange("2:2").insert("Down");`), "script-legacy-dynamic"},
		// A script's column delete (seed 366284).
		{"script", "kinds", `stream editor: sheet1.xml I3: =IF(N("l")=0,SUM(B2:B9*#REF!)) is stored as a dynamic array, not a legacy array`, script(`ws.getRange("I3").setFormula("=IF(N(\"d\")=0,B2:B4*2)");` + "\n" + `ws.getRange("C:C").delete("Left");`), "script-legacy-dynamic"},
		{"set", "xml", `stream editor: xl/tables/table1.xml: XML syntax error on line 2: illegal character code U+001F`, plain("set"), "table-header-control"},
		{"write", "text", `Data!O3: typed the text "_x0041_", and the cell holds number 18, in the results of the dynamic array at O2:O3`, plain("write"), "spill-overwrites"},
		{"do", "kinds", `Calc!E3: the edit should leave a dynamic array there, and it holds no formula (it was lost), in the results of the dynamic array at E2:E5`, batch("set 'Calc!E3' '=IF(N(\"d\")=0,SUM(IF(B2:B6>3,B2:B6)))' 'Calc!B7' '=IF(N(\"p\")=0,UPPER(A7))' 'Calc!E2' '=IF(N(\"d\")=0,B2:B5>3)'\n"), "spill-overwrites"},
		{"script", "editors", `Q 2026!C3: stream holds nothing, full holds text "Apple"`, Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Q 2026\");\nws.getRange(\"C2\").setFormula(\"=IF(N(\\\"d\\\")=0,A2:A4&\\\"!\\\")\");\nws.getRange(\"C2\").copyFrom(\"A1:C4\");\n"}}, "map-order"},
		{"write", "text", `Model!F6: typed the text "5 > 3 & 2 < 4", and the cell holds nothing, in the results of the dynamic array at F4:F7`, plain("write"), "spill-overwrites"},
		{"set", "kinds", `Summary!E3: the edit should leave a plain formula there, and it holds no formula (it was lost), in the results of the dynamic array at E2:E3`, Edit{Class: "set", Args: []string{"set", FileArg, "Summary!E3", `=IF(N("p")=0,COUNTIF(B2:B9,">3"))`, "Summary!E2", `=IF(N("p")=0,'A<B'!B2+1)`}}, "spill-overwrites"},
		{"rows-insert", "kinds", `stream editor: sheet2.xml C5: =IF(N("p")=0,Data!A2&"!") is stored as a dynamic array, not a plain formula`, plain("rows-insert"), "insert-turns-dynamic"},
		{"rows-delete", "kinds", `stream editor: sheet3.xml D2: =IF(N("p")=0,IF(A2="","none",A2)) is stored as a dynamic array, not a plain formula`, plain("rows-delete"), "insert-turns-dynamic"},
		// A range an aggregate reads (seed 548657).
		{"rows-insert", "kinds", `stream editor: sheet2.xml C2: =IF(N("p")=0,MIN(C2:E2)) is stored as a dynamic array, not a plain formula`, plain("rows-insert"), "insert-turns-dynamic"},
		{"cols-delete", "kinds", `stream editor: sheet1.xml I2: =IF(N("p")=0,COUNTIF(B2:B9,">3")) is stored as a dynamic array, not a plain formula`, plain("cols-delete"), "insert-turns-dynamic"},
		{"set", "text", `Data!U3: typed the text "N/A", and the cell holds text "R&D!", in the results of the dynamic array at U2:U4`, plain("set"), "spill-overwrites"},
		{"set", "reach", `Data!M3: outside the edit's reach, it held number 2 shown as #,##0_);(#,##0), and is shown as General now, in the results of the dynamic array at M2:M3`, plain("set"), "spill-overwrites"},
		{"set", "reach", `Data!M3: outside the edit's reach, it held number 2 shown as #,##0_);(#,##0), and is shown as General now`, plain("set"), ""},
		{"set", "reach", `Data!M3: outside the edit's reach, it held number 2 shown as General, and is shown as 0.00 now, in the results of the dynamic array at M2:M3`, plain("set"), ""},
		// A text with a quote in it, quoted as Go quotes it (seed 402252).
		{"write", "text", `Data!U3: typed the text "say \"hi\"", and the cell holds text "Fig!", in the results of the dynamic array at U2:U4`, plain("write"), "spill-overwrites"},
		{"write", "text", `Data!I2: typed the text "say \"_x005F_x0041_\"", and the cell holds text "say \"_x005FA\""`, plain("write"), "escape"},
		{"replace", "reach", `Data!E1: outside the edit's reach, it held text "say \"hi\"", and holds text "Column5" now`, plain("replace"), "table-header-renamed"},
		{"replace", "reach", `Data!E1: outside the edit's reach, it held text "1200", and holds text "Column5" now`, plain("replace"), "table-header-renamed"},
		{"script", "editors", `Q 2026!F4: stream holds number 2, full holds text "!"`, Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Q 2026\");\nws.getRange(\"E4\").setFormula(\"=IF(N(\\\"d\\\")=0,SEQUENCE(1,2))\");\nws.getRange(\"F2\").setFormula(\"=IF(N(\\\"d\\\")=0,A2:A4&\\\"!\\\")\");\n"}}, "map-order"},
		{"write", "editors", `R&D 2!E4: stream holds number 2, full holds text ""`, Edit{Class: "write", Args: []string{"write", FileArg, "R&D 2!D2", "--from", "tsv"}, Stdin: "=IF(N(\"p\")=0,OR(F2,D2>250))\t=IF(N(\"d\")=0,A2:A4&\"!\")\n=IF(N(\"d\")=0,SEQUENCE(1,2))\n"}, "map-order"},
		// A script's paste over a dynamic array an earlier edit set: a formula
		// in one editor's file, none in the other's (seed 720212).
		{"script", "editors", `Data!H7: stream has =IF(N("p")=0,IF(#REF!>1000,1,"")), full has =`, script(`ws.getRange("F4").copyFrom("H5:J8");` + "\n" + `ws.getRange("H5").moveTo("F13");`), "map-order"},
		{"script", "editors", `Data!H6: stream has =, full has =IF(N("p")=0,IF(#REF!>1000,1,""))`, script(`ws.getRange("F4").copyFrom("H5:J8");` + "\n" + `ws.getRange("H5").moveTo("F13");`), "map-order"},
		{"script", "kinds", `Data!H6: the edit should leave a plain formula there, and it holds no formula (it was lost), in the results of the dynamic array at H5:H7`, script(`ws.getRange("F4").copyFrom("H5:J8");` + "\n" + `ws.getRange("H5").moveTo("F13");`), "map-order"},
		{"replace", "kinds", `stream editor: sheet1.xml K12: =IF(N("d")=0,PRODUCT(IF(#REF!:#REF!>3,#REF!:#REF!))) is stored as a plain formula, not a dynamic array`, plain("replace"), "replace-loses-kind"},
		// A batch or a script holding an insert or a delete (the audit's seed
		// 664355, fuzzer-long's audit f1.1).
		{"do", "kinds", `stream editor: sheet3.xml D2: =IF(N("p")=0,OR(E2,C2>250)) is stored as a dynamic array, not a plain formula`, batch("set 'Data 2!E9' '=IF(N(\"p\")=0,IF(B9>1000,1,\"\"))'\ncols insert 'Data 2!A' -n 2\n"), "insert-turns-dynamic"},
		{"script", "kinds", `stream editor: sheet3.xml D2: =IF(N("p")=0,OR(E2,C2>250)) is stored as a dynamic array, not a plain formula`, script(`ws.getRange("E2").setFormula("=IF(N(\"p\")=0,OR(F2,D2>250))");` + "\n" + `ws.getRange("A:A").delete("Left");`), "insert-turns-dynamic"},
		// The full editor keeping a dynamic array's metadata on a plain formula
		// typed where one stood (seed 702553).
		{"write", "editors", `Bob's Plan!dynamic arrays: stream has F5, full has F4 F5 (the full editor's alone: F4)`, Edit{Class: "write", Args: []string{"write", FileArg, "Bob's Plan!F4", "--from", "tsv"}, Stdin: "=IF(N(\"p\")=0,SUM(B2:B6))\n"}, "full-keeps-array-metadata"},
		{"set", "editors", `Data!dynamic arrays: stream has none, full has F7 (the full editor's alone: F7)`, Edit{Class: "set", Args: []string{"set", FileArg, "Data!F7", `=IF(N("p")=0,NOT(F7))`}}, "full-keeps-array-metadata"},
		// A formula read by its row, moved by a script's insert (seed 915488).
		{"script", "editors", `R&D!C3 (=IF(N("p")=0,'Data'!C2:C3*2)): stream holds number 6, full holds number 0.58`, script(`ws.getRange("H7").setFormula("=IF(N(\"p\")=0,NOT(F7))");` + "\n" + `ws.getRange("1:1").insert("Down");`), "intersection-kept"},
		{"script", "editors", `Notes!C1 (=IF(N("p")=0,'Data'!C2:C3*10)): stream holds number 637.4, full holds error "#VALUE!"`, script(`ws.getRange("1:1").delete("Up");`), "intersection-kept"},
		// Notes (fix round 3): the streaming editor's drawing prefix, and a note
		// lost after a sheet was copied.
		{"rows-insert", "xml", `stream editor: xl/drawings/commentsDrawing1.vml: <v:shape>: the prefix "x" is not declared`, plain("rows-insert"), "stream-note-prefix"},
		{"cols-delete", "xml", `stream editor: xl/drawings/commentsDrawing1.vml: <v:shape>: the prefix "v" is not declared`, plain("cols-delete"), "stream-note-prefix"},
		{"rows-insert", "xml", `stream editor: xl/worksheets/sheet1.xml: <row>: the prefix "x" is not declared`, plain("rows-insert"), ""},
		{"set", "xml", `stream editor: xl/drawings/commentsDrawing1.vml: <v:shape>: the prefix "x" is not declared`, plain("set"), ""},
		{"clear", "reach", `Data: outside the edit's reach, the note "check this" on C1 is not there as it was: the edit's shifts give the note "check this" on C1, and the sheet holds nothing of its kind`, plain("clear"), "sheet-copy-note"},
		{"do", "reach", `Q 2027: outside the edit's reach, the note "say \"hi\"" on F1 is not there as it was: the edit's shifts give the note "say \"hi\"" on F1, and the sheet holds nothing of its kind`, plain("do"), "sheet-copy-note"},
		{"set", "editors", `Data!the streaming editor's file holds note D1 "check this" "\"check this\" false" "" "", the full editor's does not`, plain("set"), "sheet-copy-note"},
		{"rows-insert", "editors", `Data!the full editor's file holds note B4 "check this" "\"check this\" false" "" "", the streaming editor's does not`, plain("rows-insert"), "sheet-copy-note"},
		// A dynamic array spilling over a merge replaces the merge's text.
		{"set", "reach", `Data!A18: outside the edit's reach, it held text "Remarks", and holds boolean false now`, plain("set"), "spill-into-merge"},
		{"write", "reach", `Data!A17: outside the edit's reach, it held text "Remarks", and holds text "ÅÄÖ 日本語" now`, plain("write"), "spill-into-merge"},
		{"do", "reach", `Data!A12: outside the edit's reach, it held text "Remarks", and holds number 60 now`, plain("do"), "spill-into-merge"},
		{"set", "reach", `Data!A18: outside the edit's reach, it held text "Apple", and holds boolean false now`, plain("set"), ""},
		{"set", "reach", `Data!B18: outside the edit's reach, it held text "Remarks", and holds boolean false now`, plain("set"), ""},
		// A paste over arrays in a batch (a copy line), seed 2394518.
		{"do", "editors", `Data!O15: stream holds nothing, full holds text "say \"hi\""`, batch("set 'Data!S6' '=IF(N(\"d\")=0,SUM(IF(B2:B6>3,B2:B6)))'\ncopy 'Data!A2:A3' 'Data!O14'\n"), "map-order"},
		{"do", "editors", `Data!O15: stream holds nothing, full holds text "say \"hi\""`, batch("set 'Data!S6' '=IF(N(\"p\")=0,SUM(B2:B6))'\ncopy 'Data!A2:A3' 'Data!O14'\n"), ""},
		// A note's text changed, or a note made, is another bug.
		{"set", "reach", `Data: outside the edit's reach, the note "check this" on C1 is not there as it was: the edit's shifts give the note "check this" on C1, and the sheet holds note "other" on C1`, plain("set"), ""},
		{"set", "reach", `Data: outside the edit's reach, the edit made a note "check this" on C1 it does not name`, plain("set"), ""},
		// New: not a known bug's.
		{"script", "editors", `Notes!C1 (=IF(N("p")=0,'Data'!C2:C3*10)): stream holds number 637.4, full holds error "#REF!"`, script(`ws.getRange("1:1").delete("Up");`), ""},
		{"script", "editors", `R&D!C3 (=IF(N("p")=0,SUM('Data'!C2:C3))): stream holds number 6, full holds number 0.58`, script(`ws.getRange("1:1").insert("Down");`), ""},
		{"script", "editors", `R&D!C3 (=IF(N("p")=0,'Data'!C2:C3*2)): stream holds number 6, full holds number 0.58`, script(`ws.getRange("H7").setFormula("=IF(N(\"p\")=0,NOT(F7))");`), ""},
		{"script", "editors", `R&D!C3 (=IF(N("p")=0,'Data'!C2:C3*2)): stream holds number 6, full holds nothing (pending)`, script(`ws.getRange("1:1").insert("Down");`), ""},
		{"write", "editors", `Bob's Plan!dynamic arrays: stream has F4 F5, full has F5 (the streaming editor's alone: F4)`, Edit{Class: "write", Args: []string{"write", FileArg, "Bob's Plan!F4", "--from", "tsv"}, Stdin: "=IF(N(\"p\")=0,SUM(B2:B6))\n"}, ""},
		{"write", "editors", `Bob's Plan!dynamic arrays: stream has F5, full has F4 (the streaming editor's alone: F5, the full editor's alone: F4)`, Edit{Class: "write", Args: []string{"write", FileArg, "Bob's Plan!F4", "--from", "tsv"}, Stdin: "=IF(N(\"p\")=0,SUM(B2:B6))\n"}, ""},
		{"write", "editors", `Bob's Plan!dynamic arrays: stream has F5, full has F4 F5 (the full editor's alone: F4)`, Edit{Class: "write", Args: []string{"write", FileArg, "Bob's Plan!F4", "--from", "tsv"}, Stdin: "=IF(N(\"d\")=0,SEQUENCE(2))\n"}, ""},
		{"do", "kinds", `stream editor: sheet3.xml D2: =IF(N("p")=0,OR(E2,C2>250)) is stored as a dynamic array, not a plain formula`, batch("set 'Data 2!E9' 5\nsort 'Data 2!A1:A4' A\n"), ""},
		{"do", "kinds", `stream editor: sheet1.xml I2: =IF(N("p")=0,B3:B4*2) is stored as a dynamic array, not a plain formula`, batch("rows insert 'Data!2'\n"), ""},
		// The mutant rewrite-loses-kind: a move rewrites a formula's kind.
		{"move", "kinds", `stream editor: sheet1.xml I2: =IF(N("p")=0,Plan!H6:H7*2) is stored as a dynamic array, not a plain formula`, plain("move"), ""},
		{"replace", "kinds", `stream editor: sheet1.xml K12: =IF(N("d")=0,PRODUCT(IF(B2:B6>3,B2:B6))) is stored as a plain formula, not a dynamic array`, plain("replace"), ""},
		{"write", "editors", `Data!L2: stream holds text "", full holds nothing (pending)`, Edit{Class: "write", Args: []string{"write", FileArg, "Data!B2", "--from", "tsv"}, Stdin: "=IF(N(\"d\")=0,SEQUENCE(2))\n"}, ""},
		// WP-12's mutant: a set on a new sheet loses a formula, in no array.
		{"set", "kinds", `Q1 "draft"!F7: the edit should leave a plain formula there, and it holds no formula (it was lost)`, Edit{Class: "set", Args: []string{"set", FileArg, `Q1 "draft"!F7`, `=IF(N("p")=0,IF(B7>3,"big",""))`, `Q1 "draft"!B9`, "5", `Q1 "draft"!A1`, `=IF(N("d")=0,SEQUENCE(2))`}}, ""},
		{"set", "kinds", `Summary!E3: the edit should leave a plain formula there, and it holds no formula (it was lost)`, Edit{Class: "set", Args: []string{"set", FileArg, "Summary!E2", "1", "Summary!E3", `=IF(N("p")=0,B3)`}}, ""},
		// The mutant rows-insert-arrays: a plain formula reading a range turned dynamic.
		{"cols-insert", "kinds", `stream editor: sheet1.xml I2: =IF(N("p")=0,B3:B4*2) is stored as a dynamic array, not a plain formula`, plain("cols-insert"), ""},
		{"rows-insert", "kinds", `stream editor: sheet2.xml C3: =IF(N("p")=0,'R&D'!C2:C3*10) is stored as a dynamic array, not a plain formula`, plain("rows-insert"), ""},
		{"rows-insert", "kinds", `stream editor: sheet2.xml C3: =IF(N("p")=0,SUM(B2:B6)*B3:B4) is stored as a dynamic array, not a plain formula`, plain("rows-insert"), ""},
		{"do", "kinds", `Data!E4: the edit should leave a dynamic array there, and it holds no formula (it was lost)`, batch("set 'Data!E4' '=IF(N(\"d\")=0,SEQUENCE(2))'\n"), ""},
		{"move", "kinds", `Data!E4: the edit should leave a dynamic array there, and it holds no formula (it was lost)`, plain("move"), ""},
		{"set", "text", `Data!I2: typed the text "Apple", and the cell holds text "APPLE"`, plain("set"), ""},
		{"set", "text", `Data!I2: typed the text "R&D", and the cell holds nothing`, plain("set"), ""},
		{"set", "text", `Data!I2: typed the text "say \"hi\"", and the cell holds text "say "`, plain("set"), ""},
		{"replace", "reach", `Data!E1: outside the edit's reach, it held text "Date", and holds text "1,200" now`, plain("replace"), ""},
		{"set", "text", `Data!I2: typed the text "_x0041_", and the cell holds text "A"`, plain("set"), ""},
		{"move", "editors", `Data!B2: stream holds number 2088, full holds nothing (pending)`, plain("move"), ""},
		{"copy", "repeat", `stream editor: the same edit on the same bytes worked once and failed once (exit 1: no such sheet)`, plain("copy"), ""},
		{"script", "kinds", `stream editor: sheet1.xml U2: =IF(N("l")=0,UPPER(A2:A4)) is stored as a dynamic array, not a legacy array`, script(`ws.getRange("U2").moveTo("U5");`), ""},
		{"set", "xml", `stream editor: xl/worksheets/sheet1.xml: XML syntax error on line 2: illegal character code U+0001`, plain("set"), ""},
		{"script", "editors", `Q 2026!C3: stream holds nothing, full holds text "Apple"`, script(`ws.getRange("C2").copyFrom("A1:C4");`), ""},
		// A formula lost by a script with no paste, or one rewritten otherwise
		// by each editor in a paste.
		{"script", "editors", `Data!H7: stream has =IF(N("p")=0,IF(#REF!>1000,1,"")), full has =`, script(`ws.getRange("H5").moveTo("F13");`), ""},
		{"script", "editors", `Data!I3: stream has =IF(N("p")=0,B4:B5*10), full has =IF(N("p")=0,#REF!*10)`, script(`ws.getRange("F4").copyFrom("H5:J8");`), ""},
		// A formula lost by a script's paste outside any array's results, or by
		// a script with no paste in them.
		{"script", "kinds", `Data!H6: the edit should leave a plain formula there, and it holds no formula (it was lost)`, script(`ws.getRange("F4").copyFrom("H5:J8");`), ""},
		{"script", "kinds", `Data!H6: the edit should leave a plain formula there, and it holds no formula (it was lost), in the results of the dynamic array at H5:H7`, script(`ws.getRange("H6").setFormula("=IF(N(\"p\")=0,B6)");`), ""},
		{"move", "reach", `Data!G6: outside the edit's reach, its formula =B6*2 is rewritten as =E6*2, and the edit's shift of references gives =B6*2`, plain("move"), ""},
	} {
		f := Finding{Signature: Signature{tc.class, tc.invariant, 4001600827, true}, Detail: tc.detail, Edits: []Edit{tc.edit}}
		got := ""
		if b := kn.Bug(f); b != nil {
			got = b.Name
		}
		if got != tc.want {
			t.Errorf("%s %s %q: the bug %q, want %q", tc.class, tc.invariant, tc.detail, got, tc.want)
		}
	}
}

// TestBugRepros replays the repros of the bugs known.txt names, those on the
// committed pool (and with SHEETS_FUZZ=long those on long mode's pools,
// whose pools it makes), and reports each that no longer shows its bug as
// fixed, and each it could not replay with why: a report, never a failure
// (the brief: a fix of a listed bug must not fail the suite until known.txt
// is edited; and with no uv the committed pool still runs: fuzzer-long's
// audit, f1.2 and f1.5). It fails only when the repros cannot be replayed at
// all, or known.txt gives the default tests too few to replay.
func TestBugRepros(t *testing.T) {
	text, err := os.ReadFile("known.txt")
	if err != nil {
		t.Fatal(err)
	}
	kn, err := ReadKnown(string(text))
	if err != nil {
		t.Fatal(err)
	}
	long := os.Getenv("SHEETS_FUZZ") == "long"
	pool := loadPool(t)
	ps := &Pools{Root: repoRoot, Dir: t.TempDir(), Committed: pool}
	committed := 0
	for _, b := range kn.Bugs {
		if !b.Repro.Long {
			committed++
		}
	}
	if committed < 3 {
		t.Errorf("only %d bugs have a repro on the committed pool: the default tests replay those", committed)
	}
	if !long {
		kn.Bugs = slices.DeleteFunc(kn.Bugs, func(b Bug) bool { return b.Repro.Long })
	}
	fx, unchecked, err := bugsFixed(kn, pool, ps, t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range fx {
		t.Logf("fixed: %s (reported, never a failure: remove its line once fixed)", f)
	}
	for _, u := range unchecked {
		t.Logf("not replayed, its pool not made (reported, never a failure): %s", u)
	}
	t.Logf("of %d bug lines, %d repros still show their bug, %d fixed, %d not replayed", len(kn.Bugs), len(kn.Bugs)-len(fx)-len(unchecked), len(fx), len(unchecked))
}

// TestBugsFixedReported: a bug whose repro no longer shows it is listed as
// fixed, and one whose pool cannot be made as not replayed, with its own
// pool's failure; neither is an error (fuzzer-long's audit, f1.2, f1.5).
// Under -race only the second: a fixed bug's repro is replayed 10 times,
// 16 s there (1 s in the plain run, which holds both).
func TestBugsFixedReported(t *testing.T) {
	fixedLine := `bug escape: set,write text; repro 100001; detail typed the text "[^"]*_[xX][0-9A-Fa-f]{4}_[xX][0-9A-Fa-f]{4}_[^"]*", and the cell holds text "` + "\n"
	if raceOn {
		fixedLine = ""
	}
	kn, err := ReadKnown(fixedLine + `bug far: sort editors; repro long 700012; detail x
bug farther: move kinds; repro long 701600; detail x
`)
	if err != nil {
		t.Fatal(err)
	}
	pool := loadPool(t)
	// No fuzzgen.py under this root: no pool can be made.
	ps := &Pools{Root: t.TempDir(), Dir: t.TempDir(), Committed: pool}
	fx, unchecked, err := bugsFixed(kn, pool, ps, t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	if !raceOn && (len(fx) != 1 || !strings.HasPrefix(fx[0], "escape (known.txt:1): its repro, 100001, makes ")) || raceOn && len(fx) != 0 {
		t.Errorf("fixed: %q", fx)
	}
	if len(unchecked) != 2 || !strings.Contains(unchecked[0], "pool 7000: ") || !strings.Contains(unchecked[0], "--seed 7000 ") ||
		!strings.Contains(unchecked[1], "pool 7004: ") || !strings.Contains(unchecked[1], "--seed 7004 ") {
		t.Errorf("not replayed, each with its own pool's failure: %q", unchecked)
	}
}

// TestBugReproAnyFace: a bug line whose repro shows another face of the same
// bug is not fixed (map-order's script paste shows its editors face or its
// kinds face, as the editors' layouts fall), and one whose repro shows only
// another bug's face is.
func TestBugReproAnyFace(t *testing.T) {
	kn, err := ReadKnown(`bug sorts: sort kinds; repro 100047; detail no face of this
bug sorts: sort editors; repro 100047; detail : stream holds .+, full holds (nothing( \(pending\))?|number 0)$
bug other: sort repeat; repro 100047; detail no face of this
`)
	if err != nil {
		t.Fatal(err)
	}
	pool := loadPool(t)
	ps := &Pools{Root: repoRoot, Dir: t.TempDir(), Committed: pool}
	// One replay for each sorts line, and 10 for other's, which never shows:
	// under -race, the sorts lines alone.
	if raceOn {
		kn.Bugs = kn.Bugs[:2]
	}
	fx, unchecked, err := bugsFixed(kn, pool, ps, t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	if len(unchecked) != 0 {
		t.Errorf("not replayed: %q", unchecked)
	}
	want := 1
	if raceOn {
		want = 0
	}
	if len(fx) != want || want == 1 && !strings.HasPrefix(fx[0], "other (known.txt:3): its repro, 100047, makes sort editors 100047, which it does not show") {
		t.Errorf("fixed: %q, want only other's", fx)
	}
}

// TestStopAtNew: a run with Stop ends at the first finding it says (the
// mutants' runs stop at the first one known.txt does not list), and Run
// says when it found it. (Every command here deletes the workbook: each
// sequence ends at its first edit, a finding of the invariant missing.)
func TestStopAtNew(t *testing.T) {
	pool := loadPool(t)
	afterCommand(t, func([]string) bool { return true }, func(path string) { os.Remove(path) })
	stops := 0
	res, err := Run(pool, Options{Seed: 1, Edits: 2000, Dir: t.TempDir(), Stop: func(Finding) bool { stops++; return stops == 3 }})
	if err != nil {
		t.Fatal(err)
	}
	if !res.Stopped || len(res.Findings) != 3 || res.Sequences != 3 || res.Edits >= 2000 {
		t.Fatalf("stopped %v after %d edits in %d sequences with %d findings", res.Stopped, res.Edits, res.Sequences, len(res.Findings))
	}
	if f := res.Findings[0]; f.After <= 0 || f.After > res.Elapsed || f.Long || f.Pool != 0 {
		t.Errorf("the finding %s: after %s of %s, long %v, pool %d", f.Signature, f.After, res.Elapsed, f.Long, f.Pool)
	}
}

// TestPoolsFallBack: when a fresh pool cannot be made (no fuzzgen.py here,
// as when uv cannot fetch openpyxl or XlsxWriter), long mode runs the
// committed pool in its place and says why; its findings are not long ones.
func TestPoolsFallBack(t *testing.T) {
	pool := loadPool(t)
	ps := &Pools{Root: t.TempDir(), Dir: t.TempDir(), Committed: pool, Run: 7}
	p, err := ps.For(0)
	if err != nil {
		t.Fatal(err)
	}
	if p.Seed != 0 || len(p.Files) != len(pool) || ps.Failed == "" {
		t.Fatalf("pool %d of %d files; failed: %q", p.Seed, len(p.Files), ps.Failed)
	}
	if !strings.Contains(ps.Failed, "fuzzgen.py --seed 7000") {
		t.Errorf("the failure does not say what ran: %s", ps.Failed)
	}
	if len(ps.Files(7000)) != len(pool) || len(ps.Files(0)) != len(pool) {
		t.Error("the pool fallen back on is not the committed one")
	}
	if raceOn {
		return
	}
	res, err := Run(pool, Options{Seed: 1, Edits: 12, Dir: t.TempDir(), Pools: ps.For})
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range res.Findings {
		if f.Long || f.Pool != 0 {
			t.Errorf("%s on pool %d", f.Signature, f.Pool)
		}
	}
	if len(res.Pools) != 0 {
		t.Errorf("pools generated: %v", res.Pools)
	}
}

// TestFreshPools: long mode's pools, made fresh by fuzzgen.py (only with
// SHEETS_FUZZ=long: uv may fetch openpyxl and XlsxWriter, and the default
// tests use no network): each seed's pool is the same each time, another
// seed's another; each holds files from both producers, which hold the
// invariants, with every formula naming its kind, as the committed pool does.
//
// When uv cannot run openpyxl and XlsxWriter at all (no uv, or no network to
// fetch them), it is skipped, and says so: long mode runs the committed pool
// in their place, as the brief's failure path has it (fuzzer-long's audit,
// f1.5). A pool that cannot be made with uv working is a failure.
func TestFreshPools(t *testing.T) {
	if os.Getenv("SHEETS_FUZZ") != "long" {
		t.Skip("SHEETS_FUZZ=long: makes pools with uv and fuzzgen.py")
	}
	if err := uvReady(); err != nil {
		t.Skipf("long mode could not generate fresh pools: %v; the committed pool runs in their place (TestFuzz's log says so)", err)
	}
	dir := t.TempDir()
	a, err := Generate(repoRoot, 7000, dir)
	if err != nil {
		t.Fatal(err)
	}
	b, err := Generate(repoRoot, 7000, dir)
	if err != nil {
		t.Fatal(err)
	}
	c, err := Generate(repoRoot, 7001, dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(a.Files) != PoolFiles || len(c.Files) != PoolFiles {
		t.Fatalf("%d and %d files, want %d", len(a.Files), len(c.Files), PoolFiles)
	}
	same, other := 0, 0
	for i := range a.Files {
		if string(a.Files[i].Data) == string(b.Files[i].Data) {
			same++
		}
		if string(a.Files[i].Data) != string(c.Files[i].Data) {
			other++
		}
	}
	if same != PoolFiles || other < PoolFiles-2 {
		t.Errorf("one seed twice: %d of %d files the same; another seed: %d other", same, PoolFiles, other)
	}
	producers := map[string]int{}
	for _, p := range []Pool{a, c} {
		for _, wb := range p.Files {
			producers[strings.SplitN(wb.Name, "-", 2)[0]]++
			for _, v := range checkFile(wb.Data) {
				t.Errorf("pool %d %s: %s: %s", p.Seed, wb.Name, v.invariant, v.detail)
			}
			cells, _, err := formulas(mustZip(t, wb.Data))
			if err != nil {
				t.Fatal(err)
			}
			for _, f := range cells {
				if marker(f.text) == "" {
					t.Errorf("pool %d %s: %s!%s: =%s names no kind", p.Seed, wb.Name, f.sheet, f.cell, f.text)
				}
			}
		}
	}
	if producers["openpyxl"] != PoolFiles || producers["xlsxwriter"] != PoolFiles {
		t.Errorf("files by producer: %v", producers)
	}
}

// TestUntilNew: a mutant's run stops at the first finding known.txt does not
// list (by signature or as a bug's) whose sequence the tree without the
// mutant (SHEETS_FUZZ_BASE) makes with no finding; one the tree makes too is
// the code's, not the mutant's (fuzzer-long's audit, f0.3). The tree's test
// binary is a script here, which replays 1 clean and anything else with a
// finding.
func TestUntilNew(t *testing.T) {
	kn, err := ReadKnown("move kinds 100003\nbug sorts: sort editors; repro 100047; detail full holds nothing\n")
	if err != nil {
		t.Fatal(err)
	}
	if stop, _ := untilNew(t, kn, false); stop != nil {
		t.Fatal("a stop without SHEETS_FUZZ_UNTIL=new")
	}
	base := t.TempDir()
	pkg := filepath.Join(base, "internal", "fuzz")
	if err := os.MkdirAll(pkg, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `#!/bin/sh
if [ "$SHEETS_FUZZ_REPRO" = 1 ]; then echo "    fuzz_test.go:1: seed 1: no finding"; exit 0; fi
echo "    fuzz_test.go:1: finding long copy kinds $SHEETS_FUZZ_REPRO (SHEETS_FUZZ=$SHEETS_FUZZ)"; exit 1
`
	if err := os.WriteFile(filepath.Join(pkg, "fuzz.test"), []byte(script), 0o755); err != nil {
		t.Fatal(err)
	}
	t.Setenv("SHEETS_FUZZ_BASE", base)
	stop, theirs := untilNew(t, kn, true)
	for _, tc := range []struct {
		f    Finding
		want bool
	}{
		{Finding{Signature: Signature{"move", "kinds", 100003, false}}, false},                                    // listed
		{Finding{Signature: Signature{"sort", "editors", 4001600015, true}, Detail: "full holds nothing"}, false}, // a known bug's
		{Finding{Signature: Signature{"copy", "kinds", 7, true}}, false},                                          // the tree makes it too
		{Finding{Signature: Signature{"copy", "kinds", 1, false}}, true},                                          // the mutant's
	} {
		if got := stop(tc.f); got != tc.want {
			t.Errorf("%s: stop %v, want %v", tc.f.Signature, got, tc.want)
		}
	}
	if got := theirs[Signature{"copy", "kinds", 7, true}]; !strings.Contains(got, "finding long copy kinds 7 (SHEETS_FUZZ=long)") || len(theirs) != 1 {
		t.Errorf("the code's findings: %v", theirs)
	}
	if _, err := baseMakes(t.TempDir(), Finding{}); err == nil || !strings.Contains(err.Error(), "go test -c") {
		t.Errorf("a base with no test binary: %v", err)
	}
}

// TestCompareRanges: the editors' arrays, compared, say which ranges one
// editor's file has alone, which full-keeps-array-metadata's face reads.
func TestCompareRanges(t *testing.T) {
	r := func(s string) grid.Range { g, _ := grid.ParseRange(s); return g }
	for _, tc := range []struct {
		a, b []grid.Range
		want string
	}{
		{[]grid.Range{r("F5")}, []grid.Range{r("F4"), r("F5")}, "dynamic arrays: stream has F5, full has F4 F5 (the full editor's alone: F4)"},
		{[]grid.Range{r("F4"), r("F5")}, nil, "dynamic arrays: stream has F4 F5, full has none (the streaming editor's alone: F4 F5)"},
		{[]grid.Range{r("E2:E3")}, []grid.Range{r("E2:E4")}, "dynamic arrays: stream has E2:E3, full has E2:E4 (the streaming editor's alone: E2:E3, the full editor's alone: E2:E4)"},
		{[]grid.Range{r("F5"), r("F4")}, []grid.Range{r("F4"), r("F5")}, ""},
	} {
		if got := compareRanges("dynamic arrays", tc.a, tc.b); got != tc.want {
			t.Errorf("got %q, want %q", got, tc.want)
		}
	}
}
