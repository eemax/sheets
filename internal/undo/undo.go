// Package undo keeps the previous versions of files sheets overwrites, so
// an agent can take back a bad write. Versions live in the user cache
// directory, not next to the file.
package undo

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/cache"
)

// Keep is how many versions are kept per file.
const Keep = 20

// Entry describes one recorded change.
type Entry struct {
	Seq     int       `json:"seq"`
	Path    string    `json:"path"`
	Command string    `json:"command"`
	Time    time.Time `json:"time"`
	After   string    `json:"after"`   // sha256 of what sheets wrote
	Existed bool      `json:"existed"` // false when the change created the file
}

func digest(b []byte) string {
	h := sha256.Sum256(b)
	return hex.EncodeToString(h[:])
}

// root is the journal directory.
func root() (string, error) { return cache.Dir("undo") }

func dirFor(path string) (string, string, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", "", err
	}
	if real, err := filepath.EvalSymlinks(abs); err == nil {
		abs = real
	}
	r, err := root()
	if err != nil {
		return "", "", err
	}
	return filepath.Join(r, digest([]byte(abs))[:16]), abs, nil
}

// Disabled reports whether SHEETS_UNDO=off turned the journal off.
func Disabled() bool { return strings.EqualFold(os.Getenv("SHEETS_UNDO"), "off") }

// Record stores the state of path before a change. before is nil when the
// file did not exist.
func Record(path string, before, after []byte, command string) error {
	if Disabled() {
		return nil
	}
	dir, abs, err := dirFor(path)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return err
	}
	entries, err := list(dir)
	if err != nil {
		return err
	}
	e := Entry{Seq: 1, Path: abs, Command: command, Time: time.Now().UTC(), After: digest(after), Existed: before != nil}
	if n := len(entries); n > 0 {
		e.Seq = entries[n-1].Seq + 1
	}
	if before != nil {
		if err := os.WriteFile(filepath.Join(dir, name(e.Seq, ".xlsx")), before, 0o600); err != nil {
			return err
		}
	}
	meta, _ := json.Marshal(e)
	if err := os.WriteFile(filepath.Join(dir, name(e.Seq, ".json")), meta, 0o600); err != nil {
		return err
	}
	entries = append(entries, e)
	for len(entries) > Keep {
		drop(dir, entries[0].Seq)
		entries = entries[1:]
	}
	return nil
}

func name(seq int, ext string) string { return fmt.Sprintf("%06d%s", seq, ext) }

func drop(dir string, seq int) {
	os.Remove(filepath.Join(dir, name(seq, ".json")))
	os.Remove(filepath.Join(dir, name(seq, ".xlsx")))
}

func list(dir string) ([]Entry, error) {
	files, err := filepath.Glob(filepath.Join(dir, "*.json"))
	if err != nil {
		return nil, err
	}
	var out []Entry
	for _, f := range files {
		b, err := os.ReadFile(f)
		if err != nil {
			continue
		}
		var e Entry
		if json.Unmarshal(b, &e) == nil {
			out = append(out, e)
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Seq < out[j].Seq })
	return out, nil
}

// List returns the recorded changes to path, oldest first.
func List(path string) ([]Entry, error) {
	dir, _, err := dirFor(path)
	if err != nil {
		return nil, err
	}
	return list(dir)
}

// ErrNothing means no change to take back.
var ErrNothing = errors.New("nothing to undo")

// Undo restores path to its state before the last recorded change. Unless
// force is set it refuses when the file changed since sheets wrote it. A
// change that created the file is undone by moving the file into the
// journal; moved names where.
func Undo(path string, force bool) (e Entry, moved string, err error) {
	dir, _, err := dirFor(path)
	if err != nil {
		return e, "", err
	}
	entries, err := list(dir)
	if err != nil {
		return e, "", err
	}
	if len(entries) == 0 {
		return e, "", ErrNothing
	}
	e = entries[len(entries)-1]
	cur, err := os.ReadFile(path)
	switch {
	case errors.Is(err, os.ErrNotExist):
		if !force {
			return e, "", fmt.Errorf("%s no longer exists; pass --force to restore it anyway", path)
		}
	case err != nil:
		return e, "", err
	case digest(cur) != e.After && !force:
		return e, "", fmt.Errorf("%s changed since sheets wrote it (edited elsewhere?); pass --force to restore the earlier version anyway", path)
	}
	if e.Existed {
		before, err := os.ReadFile(filepath.Join(dir, name(e.Seq, ".xlsx")))
		if err != nil {
			return e, "", err
		}
		if err := replace(path, before); err != nil {
			return e, "", err
		}
	} else if cur != nil {
		moved = filepath.Join(dir, strconv.Itoa(e.Seq)+".undone.xlsx")
		if err := os.WriteFile(moved, cur, 0o600); err != nil {
			return e, "", err
		}
		if err := os.Remove(path); err != nil {
			return e, "", err
		}
	}
	drop(dir, e.Seq)
	return e, moved, nil
}

func replace(path string, data []byte) error {
	tmp, err := os.CreateTemp(filepath.Dir(path), ".sheets-undo-*")
	if err != nil {
		return err
	}
	defer os.Remove(tmp.Name())
	if _, err := tmp.Write(data); err != nil {
		tmp.Close()
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	mode := os.FileMode(0o644)
	if st, err := os.Stat(path); err == nil {
		mode = st.Mode().Perm()
	}
	os.Chmod(tmp.Name(), mode)
	return os.Rename(tmp.Name(), path)
}

// Before returns the file as it was before the last n recorded changes
// (n = 1: before the last one), with the oldest of those changes. It also
// reports whether the file still is what sheets last wrote.
func Before(path string, n int) (data []byte, e Entry, current bool, err error) {
	dir, _, err := dirFor(path)
	if err != nil {
		return nil, e, false, err
	}
	entries, err := list(dir)
	if err != nil {
		return nil, e, false, err
	}
	if n < 1 || n > len(entries) {
		if len(entries) == 0 {
			return nil, e, false, ErrNothing
		}
		return nil, e, false, fmt.Errorf("only %d recorded changes to %s", len(entries), path)
	}
	e = entries[len(entries)-n]
	if cur, err := os.ReadFile(path); err == nil {
		current = digest(cur) == entries[len(entries)-1].After
	}
	if !e.Existed {
		return nil, e, current, nil
	}
	data, err = os.ReadFile(filepath.Join(dir, name(e.Seq, ".xlsx")))
	return data, e, current, err
}
