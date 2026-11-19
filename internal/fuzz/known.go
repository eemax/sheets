package fuzz

import (
	"fmt"
	"regexp"
	"slices"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/words"
)

// Known is what known.txt lists: findings by their signature (the default
// mode's, and long mode's of one run), and the bugs they show, each by what
// its findings look like (long mode's judge: a fresh seed's findings have
// signatures no list can hold).
type Known struct {
	Seeds map[Signature]bool
	Bugs  []Bug
}

// Bug is a bug known.txt names, by its findings' face: one of its edit
// classes, its invariant, and a detail (the first line of the finding's)
// that its pattern matches; with Edit, the last edit, as a shell command,
// matches that too (a batch holding a sort). Repro is a sequence that shows
// it, which a run replays to see whether it is fixed.
//
// Copied says the bug is one of sheets copied: the sheet the finding names (the
// pattern's group "sheet") must be one an edit of the sequence copied, or made
// as a copy, and not renamed away from since (a bug of a sheet copy shows only
// after one: a note lost by an edit of a workbook where no sheet was copied is
// another bug, fuzzer-long's audit, f3.1).
type Bug struct {
	Name      string
	Classes   []string
	Invariant string
	Repro     Signature
	Detail    *regexp.Regexp
	Edit      *regexp.Regexp
	Copied    bool
	Line      int // in known.txt
}

// Shows tells whether a finding is one of the bug's (copied: a sheet copy of
// the sequence's is among its causes, see Bug).
func (b Bug) Shows(f Finding) bool {
	if f.Invariant != b.Invariant || !slices.Contains(b.Classes, f.Class) {
		return false
	}
	detail, _, _ := strings.Cut(f.Detail, "\n")
	if !b.Detail.MatchString(detail) {
		return false
	}
	if b.Edit != nil && !(len(f.Edits) > 0 && b.Edit.MatchString(f.Edits[len(f.Edits)-1].String())) {
		return false
	}
	if b.Copied {
		m := b.Detail.FindStringSubmatch(detail)
		if m == nil || b.Detail.SubexpIndex("sheet") < 0 {
			return false
		}
		return copiedSheets(f)[strings.ToLower(m[b.Detail.SubexpIndex("sheet")])]
	}
	return true
}

// sheetOp is a sheet copy or rename an edit makes.
type sheetOp struct {
	copy     bool // a copy, else a rename
	from, to string
}

// sheetOps are the sheet copies and renames an edit makes, in order: the
// command itself, or the lines of a batch (a script makes none).
func sheetOps(e Edit) []sheetOp {
	var cmds [][]string
	switch e.Class {
	case "sheet-copy", "sheet-rename":
		var w []string
		for _, a := range e.Args {
			if a != FileArg {
				w = append(w, a)
			}
		}
		cmds = append(cmds, w)
	case "do":
		for _, line := range strings.Split(e.Stdin, "\n") {
			if w, err := words.Texts(line); err == nil {
				cmds = append(cmds, w)
			}
		}
	}
	var ops []sheetOp
	for _, w := range cmds {
		if len(w) == 4 && w[0] == "sheet" && (w[1] == "copy" || w[1] == "rename") {
			ops = append(ops, sheetOp{w[1] == "copy", w[2], w[3]})
		}
	}
	return ops
}

// copiedSheets are the sheets (lower case) a finding's sequence copied, or
// made as a copy, as they are named after its last edit: both the sheet
// copied and its copy, followed through the renames after. An edit both
// editors refused made none (Finding.Taken; a finding that does not say took
// them all), and the last edit's sheet copies count when it is a batch: they
// come before the lines after them.
func copiedSheets(f Finding) map[string]bool {
	out := map[string]bool{}
	for i, e := range f.Edits {
		if len(f.Taken) == len(f.Edits) && !f.Taken[i] {
			continue
		}
		for _, op := range sheetOps(e) {
			from, to := strings.ToLower(op.from), strings.ToLower(op.to)
			if op.copy {
				out[from], out[to] = true, true
			} else if out[from] {
				delete(out, from)
				out[to] = true
			}
		}
	}
	return out
}

func (b Bug) String() string { return b.Name }

// Bug is the first bug a finding shows, nil for none.
func (k *Known) Bug(f Finding) *Bug {
	for i := range k.Bugs {
		if k.Bugs[i].Shows(f) {
			return &k.Bugs[i]
		}
	}
	return nil
}

// ParseKnown reads known.txt's signatures (see ReadKnown).
func ParseKnown(text string) (map[Signature]bool, error) {
	k, err := ReadKnown(text)
	if err != nil {
		return nil, err
	}
	return k.Seeds, nil
}

// ReadKnown reads known.txt:
//
//   - "<edit class> <invariant> <seed>" on a line, or several seeds of one
//     class and invariant; a line that starts with "long" holds findings of
//     long mode, on the pools it generated (see PoolSeed);
//   - "bug <name>: <class>[,<class>...] <invariant>; repro [long] <seed>;
//     detail <pattern>[; edit <pattern>][; copied]", a bug by its
//     findings' face (see Bug): long mode takes any finding that shows it as
//     known, whatever its seed. The patterns are Go's regular expressions,
//     and may hold # (no comment on a bug's line).
//
// # starts a comment.
func ReadKnown(text string) (*Known, error) {
	k := &Known{Seeds: map[Signature]bool{}}
	for i, line := range strings.Split(text, "\n") {
		if strings.HasPrefix(line, "bug ") {
			b, err := parseBug(line[len("bug "):])
			if err != nil {
				return nil, fmt.Errorf("known.txt:%d: %v", i+1, err)
			}
			for _, o := range k.Bugs {
				if o.Name == b.Name && o.Invariant == b.Invariant && slices.Equal(o.Classes, b.Classes) {
					return nil, fmt.Errorf("known.txt:%d: bug %s twice for %s %s", i+1, b.Name, strings.Join(b.Classes, ","), b.Invariant)
				}
			}
			b.Line = i + 1
			k.Bugs = append(k.Bugs, b)
			continue
		}
		if j := strings.IndexByte(line, '#'); j >= 0 {
			line = line[:j]
		}
		f := strings.Fields(line)
		if len(f) == 0 {
			continue
		}
		long := f[0] == "long"
		if long {
			f = f[1:]
		}
		if len(f) < 3 {
			return nil, fmt.Errorf("known.txt:%d: want [long] <edit class> <invariant> <seed>..., got %q", i+1, line)
		}
		var seeds []uint64
		for _, w := range f[2:] {
			seed, err := strconv.ParseUint(w, 10, 64)
			if err != nil {
				return nil, fmt.Errorf("known.txt:%d: seed %q: %v", i+1, w, err)
			}
			seeds = append(seeds, seed)
		}
		if !knownClass(f[0]) {
			return nil, fmt.Errorf("known.txt:%d: no edit class %q", i+1, f[0])
		}
		if !knownInvariant(f[1]) {
			return nil, fmt.Errorf("known.txt:%d: no invariant %q", i+1, f[1])
		}
		for _, seed := range seeds {
			k.Seeds[Signature{f[0], f[1], seed, long}] = true
		}
	}
	return k, nil
}

// parseBug reads a bug's line after "bug ".
func parseBug(s string) (Bug, error) {
	var b Bug
	name, rest, ok := strings.Cut(s, ":")
	if !ok || strings.ContainsAny(name, " \t") || name == "" {
		return b, fmt.Errorf("want bug <name>: <class>,... <invariant>; repro [long] <seed>; detail <pattern>[; edit <pattern>], got %q", s)
	}
	b.Name = name
	fields := strings.Split(rest, ";")
	head := strings.Fields(fields[0])
	if len(head) != 2 {
		return b, fmt.Errorf("bug %s: want <class>,... <invariant>, got %q", name, fields[0])
	}
	b.Classes = strings.Split(head[0], ",")
	for _, c := range b.Classes {
		if !knownClass(c) {
			return b, fmt.Errorf("bug %s: no edit class %q", name, c)
		}
	}
	if b.Invariant = head[1]; !knownInvariant(b.Invariant) {
		return b, fmt.Errorf("bug %s: no invariant %q", name, b.Invariant)
	}
	for _, f := range fields[1:] {
		key, val, _ := strings.Cut(strings.TrimSpace(f), " ")
		switch key {
		case "repro":
			w := strings.Fields(val)
			long := len(w) == 2 && w[0] == "long"
			if long {
				w = w[1:]
			}
			seed, err := strconv.ParseUint(strings.Join(w, " "), 10, 64)
			if err != nil {
				return b, fmt.Errorf("bug %s: repro %q: want [long] <seed>", name, val)
			}
			b.Repro = Signature{b.Classes[0], b.Invariant, seed, long}
		case "copied":
			if val != "" {
				return b, fmt.Errorf("bug %s: copied takes no value, got %q", name, val)
			}
			b.Copied = true
		case "detail", "edit":
			re, err := regexp.Compile(val)
			if err != nil {
				return b, fmt.Errorf("bug %s: %s: %v", name, key, err)
			}
			if key == "detail" {
				b.Detail = re
			} else {
				b.Edit = re
			}
		default:
			return b, fmt.Errorf("bug %s: no field %q (repro, detail, edit)", name, key)
		}
	}
	if b.Repro.Seed == 0 || b.Detail == nil {
		return b, fmt.Errorf("bug %s: a repro and a detail are needed", name)
	}
	if b.Copied && b.Detail.SubexpIndex("sheet") < 0 {
		return b, fmt.Errorf("bug %s: copied needs a (?P<sheet>...) group in the detail, for the sheet the finding names", name)
	}
	return b, nil
}

// Invariants are the invariants findings name that known.txt may list;
// repeat is an edit that gave different files on the same bytes, panic a
// command that panicked or a panic in judging the files it saved, after which
// nothing can be judged, and missing an edit that left no workbook at its path.
// A hang (an edit that did not finish) is never known: it ends the run.
var Invariants = []string{"xml", "kinds", "reach", "text", "editors", "undo", "repeat", "panic", "missing"}

func knownClass(name string) bool {
	for _, c := range Classes {
		if c.Name == name {
			return true
		}
	}
	return false
}

func knownInvariant(name string) bool {
	for _, v := range Invariants {
		if v == name {
			return true
		}
	}
	return false
}
