package fuzz

import (
	"bytes"
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// Long mode (SHEETS_FUZZ=long) runs for a time, not a count, on pools it
// makes fresh with scripts/fuzzgen.py: one for each round of RoundSequences
// sequences, each from its own seed, so the run's sequences start from many
// more workbooks than the committed pool's 40. A pool's seed follows from
// the sequence's own (PoolSeed), so a finding of long mode replays from its
// sequence's seed alone, as one of the default mode does.

// RoundSequences is how many sequences of a long run start from one pool
// (about 2,600 edits, a minute on Max's MacBook), and PoolFiles how many
// workbooks a pool holds (half openpyxl, half XlsxWriter).
const (
	RoundSequences = 400
	PoolFiles      = 40
)

// The releases fuzzgen.py takes (it refuses others).
const (
	openpyxlRelease   = "3.1.5"
	xlsxwriterRelease = "3.3.1"
)

// ObjectsFromRun is the first run seed whose pools hold objects beside cells:
// merges, notes, links, hidden rows and columns and a styled header
// (fuzzgen.py --objects), which the reach invariant judges outside an edit's
// reach. Pools of smaller run seeds are made as they always were, byte for
// byte, so the repros in known.txt (all of runs below it) replay on the pools
// they were found on. A fresh run's default seed is above it.
const ObjectsFromRun = 2_000_000

// WithObjects tells whether the pool of a seed holds those objects.
func WithObjects(poolSeed uint64) bool { return poolSeed/1000 >= ObjectsFromRun }

// objectsFlag is fuzzgen.py's flag for the pool of a seed, with a space before
// it, or "".
func objectsFlag(poolSeed uint64) string {
	if WithObjects(poolSeed) {
		return " --objects"
	}
	return ""
}

// PoolSeed is the seed fuzzgen.py makes the pool of a sequence of long mode
// from: run*1000 plus its round (sequence seed run*100000+i is in round
// i/RoundSequences). Sequence 700012 starts from pool 7000.
func PoolSeed(seq uint64) uint64 {
	run, i := seq/100000, seq%100000
	return run*1000 + i/RoundSequences
}

// GenerateCommand is the command that makes the pool of a seed in dir/pool,
// from the repository's root.
func GenerateCommand(seed uint64, dir string) string {
	return fmt.Sprintf("uv run --no-project --with openpyxl==%s --with xlsxwriter==%s python scripts/fuzzgen.py --seed %d --count %d --pool-only%s %s",
		openpyxlRelease, xlsxwriterRelease, seed, PoolFiles, objectsFlag(seed), dir)
}

// generateTimeout bounds one pool's making: uv may fetch the two releases
// the first time.
const generateTimeout = 5 * time.Minute

// Generate makes the pool of a seed with scripts/fuzzgen.py under root (the
// repository), in a directory of dir, and reads it.
func Generate(root string, seed uint64, dir string) (Pool, error) {
	out := filepath.Join(dir, "pool-"+strconv.FormatUint(seed, 10))
	ctx, cancel := context.WithTimeout(context.Background(), generateTimeout)
	defer cancel()
	args := []string{"run", "--no-project", "--with", "openpyxl==" + openpyxlRelease, "--with", "xlsxwriter==" + xlsxwriterRelease,
		"python", filepath.Join(root, "scripts", "fuzzgen.py"), "--seed", strconv.FormatUint(seed, 10), "--count", strconv.Itoa(PoolFiles), "--pool-only"}
	if WithObjects(seed) {
		args = append(args, "--objects")
	}
	cmd := exec.CommandContext(ctx, "uv", append(args, out)...)
	var buf bytes.Buffer
	cmd.Stdout, cmd.Stderr = &buf, &buf
	if err := cmd.Run(); err != nil {
		return Pool{}, fmt.Errorf("%s: %v: %s", GenerateCommand(seed, out), err, lastLines(buf.String(), 3))
	}
	files, err := LoadPool(filepath.Join(out, "pool"))
	if err != nil {
		return Pool{}, err
	}
	os.RemoveAll(out) // read: the pool is in memory, and made again from its seed
	return Pool{Seed: seed, Files: files}, nil
}

// uvReady tells whether uv can run the two releases fuzzgen.py takes (nil),
// or why not: no uv, or no network to fetch them the first time.
func uvReady() error {
	ctx, cancel := context.WithTimeout(context.Background(), generateTimeout)
	defer cancel()
	cmd := exec.CommandContext(ctx, "uv", "run", "--no-project", "--with", "openpyxl=="+openpyxlRelease, "--with", "xlsxwriter=="+xlsxwriterRelease,
		"python", "-c", "import openpyxl, xlsxwriter")
	var buf bytes.Buffer
	cmd.Stdout, cmd.Stderr = &buf, &buf
	if err := cmd.Run(); err != nil {
		return fmt.Errorf("uv cannot run openpyxl %s and XlsxWriter %s (%v): %s", openpyxlRelease, xlsxwriterRelease, err, lastLines(buf.String(), 3))
	}
	return nil
}

// Pools makes and keeps the pools of a long run, a round at a time, and
// falls back on the committed pool when it cannot make one (uv missing, or
// unable to fetch openpyxl or XlsxWriter): that round's sequences start from
// the committed pool, and their findings are not long ones.
type Pools struct {
	Root      string // the repository, for scripts/fuzzgen.py
	Dir       string
	Committed []Workbook
	Run       uint64 // the run's seed
	made      map[uint64]Pool
	why       map[uint64]string
	// Failed says why a pool could not be made, the first time.
	Failed string
}

// Why says why the pool of a seed could not be made, "" when it was (or was
// not tried): each pool's own failure, where Failed keeps the first
// (fuzzer-long's audit, f1.5: each long repro not replayed named the first
// failing pool's command, not its own).
func (ps *Pools) Why(seed uint64) string { return ps.why[seed] }

// For is the pool sequence i of the run starts from.
func (ps *Pools) For(i int) (Pool, error) {
	return ps.Of(PoolSeed(SequenceSeed(ps.Run, i))), nil
}

// Of is the pool of a seed, made once.
func (ps *Pools) Of(seed uint64) Pool {
	if ps.made == nil {
		ps.made, ps.why = map[uint64]Pool{}, map[uint64]string{}
	}
	if p, ok := ps.made[seed]; ok {
		return p
	}
	p, err := Generate(ps.Root, seed, ps.Dir)
	if err != nil {
		if ps.Failed == "" {
			ps.Failed = err.Error()
		}
		ps.why[seed] = err.Error()
		p = Pool{Files: ps.Committed}
	}
	ps.made[seed] = p
	return p
}

// Files are the workbooks of a pool by seed, 0 for the committed pool: what
// Shrink and Replay need for a finding.
func (ps *Pools) Files(seed uint64) []Workbook {
	if seed == 0 {
		return ps.Committed
	}
	return ps.Of(seed).Files
}

func lastLines(s string, n int) string {
	ls := strings.Split(strings.TrimSpace(s), "\n")
	if len(ls) > n {
		ls = ls[len(ls)-n:]
	}
	return strings.Join(ls, " / ")
}
