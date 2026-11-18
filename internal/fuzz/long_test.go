package fuzz

import (
	"fmt"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"
)

// maxSamples is the most workbooks long mode sends LibreOffice, spread over
// the run.
const maxSamples = 120

// fuzzLong is long mode (SHEETS_FUZZ=long): a run for a time, over pools it
// makes fresh with scripts/fuzzgen.py (one for each round of RoundSequences
// sequences, from seeds the run's seed gives), so its sequences start from
// many more workbooks than the committed 40. See README.md, Long mode.
//
//   - SHEETS_FUZZ_SEED: the run's seed (default: a fresh one, printed);
//   - SHEETS_FUZZ_TIME: how long it runs (default 10m), finishing the
//     sequence it is in;
//   - SHEETS_FUZZ_EDITS: the most edits it makes (default: no limit);
//   - SHEETS_FUZZ_UNTIL=new: stop at the first finding known.txt does not
//     list (scripts/fuzz-mutants.sh);
//   - SHEETS_FUZZ_SOFFICE: LibreOffice's soffice, when it is not on PATH.
//
// A finding known.txt does not list fails the run, as in the default mode;
// long mode's entries there start with "long". The log gives the rate of
// findings per 10,000 edits, the entries of known.txt for the run's seed that
// no longer reproduce (fixed: reported, never a failure), and LibreOffice's
// second opinion on the values of plain formulas (reported, never a failure).
func fuzzLong(t *testing.T, committed []Workbook, known map[Signature]bool) {
	seed := uint64(time.Now().UnixNano()/int64(time.Millisecond)%1_000_000) + 2
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
	if until {
		o.Stop = func(f Finding) bool { return !known[f.Signature] }
	}
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
	newFindings := judgeFindings(t, "long", res, known, ps.Files, dir, "SHEETS_FUZZ=long ")
	t.Logf("fuzz long: %d edits (%d applied) in %d sequences, %.0fs, seed %d; %d pools generated (seeds %s); %d findings, %d new",
		res.Edits, res.Applied, res.Sequences, res.Elapsed.Seconds(), seed, len(res.Pools), seedList(res.Pools), len(res.Findings), newFindings)
	if ps.Failed != "" {
		t.Logf("fuzz long: could not generate fresh pools, so the committed pool ran in their place (its findings are not long ones): %s", ps.Failed)
	}
	t.Logf("fuzz long: %.1f findings per 10,000 edits (%d in %d edits): %s", float64(len(res.Findings))*10000/float64(max(res.Edits, 1)), len(res.Findings), res.Edits, bySignature(res.Findings))
	logFindings(t, res, known)
	if res.Stopped {
		f := res.Findings[len(res.Findings)-1]
		t.Logf("fuzz long: stopped at the first finding known.txt does not list, %s, after %.0fs (SHEETS_FUZZ_UNTIL=new)", f.Signature, f.After.Seconds())
	}
	if res.Hung {
		return
	}
	logCounts(t, "long", res, false)
	logFixed(t, known, res, seed, true)
	switch {
	case until:
	case soffice == "":
		t.Logf("fuzz long: LibreOffice's second opinion: none (no soffice on PATH; SHEETS_FUZZ_SOFFICE names it)")
	default:
		op := SecondOpinion(soffice, samples, filepath.Join(dir, "lo"))
		t.Logf("fuzz long: LibreOffice's second opinion on plain formulas, in workbooks from openpyxl files (reported, never a finding): %s", op)
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
}

// TestStopAtNew: a run with Stop ends at the first finding it says (the
// mutants' runs stop at the first one known.txt does not list), and Run
// says when it found it.
func TestStopAtNew(t *testing.T) {
	if raceOn {
		t.Skip("the step judges the same code under -race; this is Run's loop")
	}
	pool := loadPool(t)
	res, err := Run(pool, Options{Seed: 1, Edits: 2000, Dir: t.TempDir(), Stop: func(Finding) bool { return true }})
	if err != nil {
		t.Fatal(err)
	}
	if !res.Stopped || len(res.Findings) != 1 || res.Edits >= 2000 {
		t.Fatalf("stopped %v after %d edits with %d findings", res.Stopped, res.Edits, len(res.Findings))
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
	res, err := Run(pool, Options{Seed: 1, Edits: 60, Dir: t.TempDir(), Pools: ps.For})
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
func TestFreshPools(t *testing.T) {
	if os.Getenv("SHEETS_FUZZ") != "long" {
		t.Skip("SHEETS_FUZZ=long: makes pools with uv and fuzzgen.py")
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
