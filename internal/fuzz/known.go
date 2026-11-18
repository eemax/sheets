package fuzz

import (
	"fmt"
	"strconv"
	"strings"
)

// ParseKnown reads known.txt: one signature per line, "<edit class>
// <invariant> <seed>", with # starting a comment. A line that starts with
// "long" is a finding of long mode, on the pool it generated (see PoolSeed).
func ParseKnown(text string) (map[Signature]bool, error) {
	known := map[Signature]bool{}
	for i, line := range strings.Split(text, "\n") {
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
		if len(f) != 3 {
			return nil, fmt.Errorf("known.txt:%d: want <edit class> <invariant> <seed>, got %q", i+1, line)
		}
		seed, err := strconv.ParseUint(f[2], 10, 64)
		if err != nil {
			return nil, fmt.Errorf("known.txt:%d: seed %q: %v", i+1, f[2], err)
		}
		if !knownClass(f[0]) {
			return nil, fmt.Errorf("known.txt:%d: no edit class %q", i+1, f[0])
		}
		if !knownInvariant(f[1]) {
			return nil, fmt.Errorf("known.txt:%d: no invariant %q", i+1, f[1])
		}
		known[Signature{f[0], f[1], seed, long}] = true
	}
	return known, nil
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
