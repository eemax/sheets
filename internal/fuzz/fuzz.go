// Package fuzz does by machine what wave 1's audits did by hand. It runs
// the sheets CLI in process (cli.Main) over random workbooks from openpyxl
// and XlsxWriter, makes random edits, and checks fixed invariants on the
// files after each one:
//
//   - xml: every part is well-formed XML, and every worksheet lists its
//     rows and the cells of each row in order, each once (xlsxtest);
//   - kinds: every formula keeps its kind (plain, legacy array, dynamic
//     array) and goes where the edit puts it, by a model of what each edit
//     does to formulas (check.go, model.go);
//   - editors: the streaming and the full editor give the same cells,
//     formulas, kinds and values (and both take the edit, or both refuse);
//   - undo: undo restores the bytes from before the edit, and an edit
//     refused leaves them as they were.
//
// Invariants judge files, never a command's output text. A panic in a
// command is reported too: nothing after it can be judged.
//
// A run is a list of sequences, each with its own seed: the seed picks a
// workbook from the pool and the edits, each made from the state the
// edits before it left. A sequence ends at its first finding. Its
// signature (edit class, invariant, the sequence's seed) is what
// known.txt lists.
package fuzz

import (
	"bytes"
	"fmt"
	"io"
	"math/rand/v2"
	"os"
	"path/filepath"
	"runtime/debug"
	"sort"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/cli"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
	"github.com/eemax/sheets/internal/xlsx"
)

// FileArg stands for the workbook in an Edit's arguments.
const FileArg = "{file}"

// Edit is one sheets command on the workbook.
type Edit struct {
	Class string   // set, write, clear, copy, move, sort, replace, rows-insert, ...
	Args  []string // the command line after "sheets"; FileArg is the workbook
	Stdin string
}

// String renders the edit as a shell command on book.xlsx.
func (e Edit) String() string {
	args := make([]string, len(e.Args))
	for i, a := range e.Args {
		if a == FileArg {
			a = "book.xlsx"
		}
		args[i] = words.Quote(a)
	}
	s := "sheets " + strings.Join(args, " ")
	if e.Stdin != "" {
		s += " <<'EOF'\n" + strings.TrimSuffix(e.Stdin, "\n") + "\nEOF"
	}
	return s
}

// Signature identifies a finding exactly: the class of the edit that broke
// an invariant, the invariant, and the seed of the sequence it ended.
type Signature struct {
	Class     string
	Invariant string
	Seed      uint64
}

func (s Signature) String() string { return fmt.Sprintf("%s %s %d", s.Class, s.Invariant, s.Seed) }

// Finding is an invariant an edit broke.
type Finding struct {
	Signature
	Workbook string // the pool file the sequence started from
	Edits    []Edit // the sequence, ending with the edit that broke it
	Detail   string
}

// Report prints the finding as a repro: seed, workbook, edits, detail.
func (f Finding) Report() string {
	var b strings.Builder
	fmt.Fprintf(&b, "finding %s\n", f.Signature)
	fmt.Fprintf(&b, "  workbook: testdata/pool/%s (copy it to book.xlsx)\n", f.Workbook)
	for i, e := range f.Edits {
		fmt.Fprintf(&b, "  %d. %s\n", i+1, strings.ReplaceAll(e.String(), "\n", "\n     "))
	}
	fmt.Fprintf(&b, "  broke %s: %s\n", f.Invariant, strings.ReplaceAll(f.Detail, "\n", "\n    "))
	return b.String()
}

// Workbook is a pool file.
type Workbook struct {
	Name string
	Data []byte
}

// LoadPool reads every .xlsx file of dir, by name.
func LoadPool(dir string) ([]Workbook, error) {
	names, err := filepath.Glob(filepath.Join(dir, "*.xlsx"))
	if err != nil {
		return nil, err
	}
	sort.Strings(names)
	var pool []Workbook
	for _, n := range names {
		data, err := os.ReadFile(n)
		if err != nil {
			return nil, err
		}
		pool = append(pool, Workbook{filepath.Base(n), data})
	}
	if len(pool) == 0 {
		return nil, fmt.Errorf("no workbooks in %s", dir)
	}
	return pool, nil
}

// Options set a run.
type Options struct {
	Seed  uint64        // the run's seed; sequence i has seed Seed*100000+i
	Edits int           // stop after this many edits
	Time  time.Duration // stop once this much time has passed (0: no cap)...
	Min   int           // ...and at least this many edits are made
	Steps int           // the most edits in a sequence (default 12)
	Dir   string        // scratch directory
}

// Result is what a run did.
type Result struct {
	Edits     int // edits made (each on both editors)
	Applied   int // edits both editors took
	Sequences int
	Capped    bool // stopped by the time cap before Edits
	Elapsed   time.Duration
	Findings  []Finding
	// Saved counts the files the editors saved (an edit, two editors) that
	// hold a formula result of each stored form, and a shared formula (see
	// Results): the forms the invariants were put to.
	Saved map[string]int
	// Repeated counts the edits that removed a dynamic array and were made
	// again, by each editor, on the bytes they first found (see repeatFrom).
	Repeated int
	// Taken counts, by edit class, the edits both editors took; Streamed
	// those of them the streaming editor could do itself (see streamed). The
	// other edits are full-editor code in both copies, so the editors
	// invariant compares that code with itself.
	Taken, Streamed map[string]int
}

// seedMix is the second word of each sequence's random generator, mixed with
// its seed: the golden ratio's bits, and 797, which was picked so that
// sequence 100139 is a paste over dynamic arrays that sheets lays out
// differently from one run to the next (copy repeat 100139 in known.txt).
// That is the sequence wave 2's audit replays hundreds of times, to see
// whether the repeat check ever misses it. Another mix makes other
// sequences; the default run's findings in known.txt are then made again.
var seedMix uint64 = 0x9e3779b97f4a7c15 + 797

// SequenceSeed is the seed of sequence i of a run.
func SequenceSeed(run uint64, i int) uint64 { return run*100000 + uint64(i) }

// Run runs sequences until the edit count or the time cap is reached.
func Run(pool []Workbook, o Options) (*Result, error) {
	start := time.Now()
	r := &Result{Saved: map[string]int{}, Taken: map[string]int{}, Streamed: map[string]int{}}
	for i := 0; r.Edits < o.Edits; i++ {
		if o.Time > 0 && r.Edits >= o.Min && time.Since(start) > o.Time {
			r.Capped = true
			break
		}
		seed := SequenceSeed(o.Seed, i)
		f, made, applied, err := sequence(pool, seed, o.Dir, o.steps(), o.Edits-r.Edits, nil, r)
		if err != nil {
			return r, err
		}
		r.Sequences++
		r.Edits += made
		r.Applied += applied
		if f != nil {
			r.Findings = append(r.Findings, *f)
		}
	}
	r.Elapsed = time.Since(start)
	return r, nil
}

func (o Options) steps() int {
	if o.Steps > 0 {
		return o.Steps
	}
	return 12
}

// Replay runs the sequence of one seed again, logging every edit and its
// outcome, and returns its finding, if any.
func Replay(pool []Workbook, seed uint64, dir string, steps int, log func(string)) (*Finding, error) {
	if steps <= 0 {
		steps = 12
	}
	f, _, _, err := sequence(pool, seed, dir, steps, steps, log, nil)
	return f, err
}

// sequence makes up to steps edits (at most budget) on a workbook the seed
// picks, and stops at the first finding. It counts what it saves and repeats
// into res, when not nil.
func sequence(pool []Workbook, seed uint64, dir string, steps, budget int, log func(string), res *Result) (f *Finding, made, applied int, err error) {
	rng := rand.New(rand.NewPCG(seed, seed^seedMix))
	wb := pool[rng.IntN(len(pool))]
	n := min(1+rng.IntN(steps), budget)
	s, err := newSession(dir, wb.Data)
	if err != nil {
		return nil, 0, 0, err
	}
	defer s.close()
	s.res = res
	var edits []Edit
	for i := 0; i < n; i++ {
		e, err := s.generate(rng)
		if err != nil {
			return nil, made, applied, err
		}
		edits = append(edits, e)
		out, err := s.step(e, i)
		if err != nil {
			return nil, made, applied, err
		}
		made++
		if out.applied {
			applied++
		}
		if log != nil {
			log(fmt.Sprintf("%d. %s\n   -> %s", i+1, e, out.summary()))
		}
		if len(out.broken) > 0 {
			v := out.broken[0]
			detail := v.detail
			for _, w := range out.broken[1:] {
				detail += "\nalso " + w.invariant + ": " + w.detail
			}
			return &Finding{Signature{e.Class, v.invariant, seed}, wb.Name, edits, detail}, made, applied, nil
		}
	}
	return nil, made, applied, nil
}

// Shrink removes the edits a finding does not need, one at a time, keeping
// each removal after which the last edit still breaks the same invariant.
func Shrink(pool []Workbook, f Finding, dir string) (Finding, error) {
	var wb []byte
	for _, w := range pool {
		if w.Name == f.Workbook {
			wb = w.Data
		}
	}
	if wb == nil {
		return f, fmt.Errorf("workbook %s not in the pool", f.Workbook)
	}
	edits := f.Edits
	same := func(try []Edit) (bool, error) {
		g, err := replayEdits(wb, try, dir)
		if err != nil {
			return false, err
		}
		if g != nil && g.edit == len(try)-1 && try[g.edit].Class == f.Class && g.invariant == f.Invariant {
			f.Detail = g.detail
			return true, nil
		}
		return false, nil
	}
	for i := len(edits) - 2; i >= 0; i-- {
		try := append(append([]Edit{}, edits[:i]...), edits[i+1:]...)
		if ok, err := same(try); err != nil {
			return f, err
		} else if ok {
			edits = try
		}
	}
	// Then the lines of batches and scripts.
	for i := range edits {
		for j := len(lines(edits[i])) - 1; j >= 0; j-- {
			e, ok := withoutLine(edits[i], j)
			if !ok {
				continue
			}
			try := append([]Edit{}, edits...)
			try[i] = e
			if ok, err := same(try); err != nil {
				return f, err
			} else if ok {
				edits = try
			}
		}
	}
	f.Edits = edits
	return f, nil
}

// lines are the commands of a do batch, or the statements of a script
// after its first (which names the sheet).
func lines(e Edit) []string {
	switch e.Class {
	case "do":
		return strings.Split(strings.TrimSuffix(e.Stdin, "\n"), "\n")
	case "script":
		return strings.Split(strings.TrimSuffix(e.Args[len(e.Args)-1], "\n"), "\n")[1:]
	}
	return nil
}

// withoutLine is a batch or script without one of its lines; false when it
// would have none left.
func withoutLine(e Edit, j int) (Edit, bool) {
	ls := lines(e)
	if len(ls) < 2 {
		return e, false
	}
	rest := strings.Join(append(append([]string{}, ls[:j]...), ls[j+1:]...), "\n") + "\n"
	switch e.Class {
	case "do":
		e.Stdin = rest
	case "script":
		code := e.Args[len(e.Args)-1]
		head := code[:strings.IndexByte(code, '\n')+1]
		e.Args = append(append([]string{}, e.Args[:len(e.Args)-1]...), head+rest)
	}
	return e, true
}

type broke struct {
	edit      int
	invariant string
	detail    string
}

// replayEdits makes the given edits on a workbook and returns the first
// invariant broken.
func replayEdits(wb []byte, edits []Edit, dir string) (*broke, error) {
	s, err := newSession(dir, wb)
	if err != nil {
		return nil, err
	}
	defer s.close()
	for i, e := range edits {
		out, err := s.step(e, i)
		if err != nil {
			return nil, err
		}
		if len(out.broken) > 0 {
			return &broke{i, out.broken[0].invariant, out.broken[0].detail}, nil
		}
	}
	return nil, nil
}

// editors names the two editors: SHEETS_STREAM's value for each.
var editors = [2]struct{ name, env string }{{"stream", "always"}, {"full", "never"}}

// session holds one workbook twice, a copy per editor.
type session struct {
	root  string
	paths [2]string
	cur   *grid.Workbook // the streaming copy, parsed: what the next edit is made from
	res   *Result        // counts the stored forms saved and the edits repeated, when not nil
}

func newSession(dir string, data []byte) (*session, error) {
	root, err := os.MkdirTemp(dir, "seq-")
	if err != nil {
		return nil, err
	}
	s := &session{root: root}
	for k, ed := range editors {
		d := filepath.Join(root, ed.name)
		if err := os.Mkdir(d, 0o755); err != nil {
			return nil, err
		}
		s.paths[k] = filepath.Join(d, "book.xlsx")
		if err := os.WriteFile(s.paths[k], data, 0o644); err != nil {
			return nil, err
		}
	}
	return s, nil
}

func (s *session) close() { os.RemoveAll(s.root) }

// generate makes the next edit, from the streaming copy.
func (s *session) generate(rng *rand.Rand) (Edit, error) {
	if s.cur == nil {
		data, err := os.ReadFile(s.paths[0])
		if err != nil {
			return Edit{}, err
		}
		if s.cur, err = xlsx.ReadBytes(data); err != nil {
			return Edit{}, fmt.Errorf("the workbook does not read: %w", err)
		}
	}
	return generate(rng, s.cur), nil
}

type violation struct{ invariant, detail string }

type outcome struct {
	codes   [2]int
	errs    [2]string
	applied bool
	broken  []violation
}

func (o outcome) summary() string {
	var b strings.Builder
	for k, ed := range editors {
		if k > 0 {
			b.WriteString("; ")
		}
		fmt.Fprintf(&b, "%s exit %d", ed.name, o.codes[k])
		if o.errs[k] != "" {
			fmt.Fprintf(&b, " (%s)", firstLine(o.errs[k]))
		}
	}
	if len(o.broken) > 0 {
		fmt.Fprintf(&b, "; broke %s: %s", o.broken[0].invariant, o.broken[0].detail)
	}
	return b.String()
}

// step makes one edit with both editors and checks the invariants. Edit i
// checks undo on the streaming copy when even, on the full one when odd;
// the other copy's edit runs with the journal off (SHEETS_UNDO=off), so
// each journal holds the one change undo takes back.
func (s *session) step(e Edit, i int) (outcome, error) {
	var o outcome
	var before, after [2][]byte
	u := i % 2
	for k, ed := range editors {
		var err error
		if before[k], err = os.ReadFile(s.paths[k]); err != nil {
			return o, err
		}
		o.codes[k], o.errs[k] = run(ed.env, k == u, e, s.paths[k])
		if after[k], err = os.ReadFile(s.paths[k]); err != nil {
			return o, err
		}
	}
	for k, ed := range editors {
		if o.codes[k] == panicked {
			o.broken = append(o.broken, violation{"panic", ed.name + " editor: " + o.errs[k]})
		}
	}
	if len(o.broken) > 0 {
		return o, nil
	}
	if (o.codes[0] == 0) != (o.codes[1] == 0) {
		o.broken = append(o.broken, violation{"editors", fmt.Sprintf("stream exit %d (%s), full exit %d (%s)",
			o.codes[0], firstLine(o.errs[0]), o.codes[1], firstLine(o.errs[1]))})
		return o, nil
	}
	if o.codes[0] != 0 {
		for k, ed := range editors {
			if !bytes.Equal(before[k], after[k]) {
				o.broken = append(o.broken, violation{"undo", ed.name + " editor: the edit was refused (" + firstLine(o.errs[k]) + ") but the file changed"})
			}
		}
		return o, nil
	}
	o.applied = true
	wb, d := compareEditors(after[0], after[1])
	prev, err := s.before(before[0])
	if err != nil {
		return o, err
	}
	if s.res != nil {
		s.res.Taken[e.Class]++
		if streamed(e, prev) {
			s.res.Streamed[e.Class]++
		}
	}
	// An edit that removes a dynamic array is made again (see repeatFrom).
	if wb != nil && arraysGone(prev, wb) >= repeatFrom {
		v, err := s.repeat(e, before, after)
		if err != nil {
			return o, err
		}
		if v != nil {
			s.cur = wb
			o.broken = append(o.broken, *v)
			return o, nil
		}
	}
	for k, ed := range editors {
		vs, cells := inspect(after[k])
		for _, v := range vs {
			v.detail = ed.name + " editor: " + v.detail
			o.broken = append(o.broken, v)
		}
		s.count(cells)
	}
	if d != "" {
		o.broken = append(o.broken, violation{"editors", d})
	} else if wb != nil {
		// Where the formulas went: the model of kinds, followed through the edit.
		m := newModel(prev)
		m.apply(e)
		if msg := m.check(wb); msg != "" {
			o.broken = append(o.broken, violation{"kinds", msg})
		}
	}
	s.cur = wb
	if k := u; !bytes.Equal(before[k], after[k]) {
		v, err := checkUndo(k, s.paths[k], before[k], after[k])
		if err != nil {
			return o, err
		}
		if v != nil {
			o.broken = append(o.broken, *v)
		}
	}
	return o, nil
}

// streamed tells whether the streaming editor could take an edit itself, on
// the workbook it was made on: a set or a write on a sheet that exists and
// adds no dynamic array, or the insert or delete of rows, on a workbook that
// holds no dynamic array and no legacy array over several cells (see
// streamable). It is the fuzzer's estimate from the edit and the workbook,
// not what the editor reported (that is the count in README.md): a set or
// rows edit inside a do batch or a script is not counted.
func streamed(e Edit, wb *grid.Workbook) bool {
	switch e.Class {
	case "set", "write":
		if strings.Contains(e.Stdin, `N("d")`) {
			return false
		}
		for i, a := range e.Args[2:] {
			if marker(a) == Dynamic {
				return false
			}
			// A cell to write: a sheet that is new needs the full editor.
			if e.Class == "write" && i == 0 || e.Class == "set" && i%2 == 0 {
				p, err := grid.ParseRef(a, func(n string) bool { return wb.Sheet(n) != nil })
				if err != nil || p.Sheet != "" && wb.Sheet(p.Sheet) == nil {
					return false
				}
			}
			if e.Class == "write" {
				break
			}
		}
	case "rows-insert", "rows-delete":
	default:
		return false
	}
	return streamable(wb)
}

// before is the streaming copy's workbook as the next edit finds it, given
// its bytes.
func (s *session) before(data []byte) (*grid.Workbook, error) {
	if s.cur != nil {
		return s.cur, nil
	}
	wb, err := xlsx.ReadBytes(data)
	if err != nil {
		return nil, fmt.Errorf("the workbook does not read: %w", err)
	}
	return wb, nil
}

// arraysGone counts the dynamic arrays an edit removed from the workbook: how
// many fewer each sheet holds (arrays that only moved are not gone), on the
// sheets at the same position in both.
func arraysGone(a, b *grid.Workbook) int {
	gone := 0
	for i := 0; i < min(len(a.Sheets), len(b.Sheets)); i++ {
		gone += max(0, len(a.Sheets[i].Spills)-len(b.Sheets[i].Spills))
	}
	return gone
}

// An edit that removes a dynamic array (one or more: net of those it adds on
// the sheet) is made again on the bytes it first found, repeats times by each
// editor. Sheets lays the results of dynamic arrays out in map order
// (spillResults), so a paste over arrays can give a different workbook from
// one run to the next (see known.txt), by chance: the shrunk paste of
// copy repeat 100139 gives one workbook 73% of the time and two others about
// 14% each. A finding that came and went would make two runs of one seed
// differ, so the repeats are enough to see such an edit nearly every time:
// the chance that all (1+repeats)*2 runs give the common workbook is below
// 0.0002% (0.73 to the 42nd power).
const repeatFrom = 1

// repeats is the repeats by each editor: 20, and 3 under -race, where each
// takes ten times as long and a miss only drops a finding the run lists as
// stale.
var repeats = 20

// repeatWords name the two workbooks a repeat compares, which compareEditors
// calls the stream's and the full editor's.
var repeatWords = strings.NewReplacer("stream has", "the first run has", "full has", "the repeat has", "stream holds", "the first run holds", "full holds", "the repeat holds")

// repeat makes the edit again on the bytes it first found, repeats times
// with each editor, and reports a file that holds other cells, formulas or
// values than the first one saved. (The bytes may differ: the order of the
// shared strings a spill adds varies, and means nothing.) Both files are
// put back as the first run left them.
func (s *session) repeat(e Edit, before, after [2][]byte) (*violation, error) {
	if s.res != nil {
		s.res.Repeated++
	}
	var v *violation
	for k, ed := range editors {
		for n := 0; n < repeats && v == nil; n++ {
			if err := os.WriteFile(s.paths[k], before[k], 0o644); err != nil {
				return nil, err
			}
			code, msg := run(ed.env, false, e, s.paths[k])
			got, err := os.ReadFile(s.paths[k])
			if err != nil {
				return nil, err
			}
			if code != 0 {
				v = &violation{"repeat", fmt.Sprintf("%s editor: the same edit on the same bytes worked once and failed once (exit %d: %s)", ed.name, code, firstLine(msg))}
			} else if _, d := compareEditors(after[k], got); d != "" {
				v = &violation{"repeat", fmt.Sprintf("%s editor: the same edit on the same bytes saved two different workbooks: %s", ed.name, repeatWords.Replace(d))}
			}
		}
		if err := os.WriteFile(s.paths[k], after[k], 0o644); err != nil {
			return nil, err
		}
	}
	return v, nil
}

// count adds a saved file's formulas to the count of stored forms: each form
// the file holds, once.
func (s *session) count(cells []formulaCell) {
	if s.res == nil {
		return
	}
	held := map[string]bool{}
	for _, f := range cells {
		held[f.result()] = true
		if f.shared {
			held[resShared] = true
		}
	}
	for form := range held {
		s.res.Saved[form]++
	}
}

// checkUndo runs sheets undo on the file an edit with editor k changed
// from before to after, checks it restores before, and puts after back.
func checkUndo(k int, path string, before, after []byte) (*violation, error) {
	code, msg := run(editors[k].env, true, Edit{Args: []string{"undo", FileArg}}, path)
	got, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var v *violation
	switch {
	case code != 0:
		v = &violation{"undo", editors[k].name + " editor: undo failed: " + firstLine(msg)}
	case !bytes.Equal(got, before):
		v = &violation{"undo", editors[k].name + " editor: undo did not restore the bytes: " + partsDiffer(before, got)}
	}
	return v, os.WriteFile(path, after, 0o644)
}

// panicked is the exit code run gives a command that panicked.
const panicked = -1

// run runs one command with an editor, in process, with the undo journal
// on or off.
func run(env string, journal bool, e Edit, path string) (code int, stderr string) {
	args := make([]string, len(e.Args))
	for i, a := range e.Args {
		if a == FileArg {
			a = path
		}
		args[i] = a
	}
	os.Setenv("SHEETS_STREAM", env)
	if journal {
		os.Unsetenv("SHEETS_UNDO")
	} else {
		os.Setenv("SHEETS_UNDO", "off")
	}
	var errb bytes.Buffer
	defer func() {
		if p := recover(); p != nil {
			code, stderr = panicked, fmt.Sprintf("%v\n%s", p, debug.Stack())
		}
	}()
	code = cli.Main(args, strings.NewReader(e.Stdin), io.Discard, &errb)
	return code, strings.TrimSpace(strings.ReplaceAll(errb.String(), path, "book.xlsx"))
}

func firstLine(s string) string {
	s = strings.TrimPrefix(s, "error: ")
	if i := strings.IndexByte(s, '\n'); i >= 0 {
		s = s[:i]
	}
	if len(s) > 160 {
		s = s[:160] + "…"
	}
	return s
}
