package cache

import (
	"bytes"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"
)

func TestPutGet(t *testing.T) {
	t.Setenv("SHEETS_CACHE_DIR", t.TempDir())
	key := Key("xlsx", []byte("file content"))
	if key == "" || key == Key("xlsx", []byte("other content")) || key == Key("csv", []byte("file content")) {
		t.Fatalf("keys do not tell entries apart: %q", key)
	}
	if _, ok := Get(key); ok {
		t.Fatal("hit in an empty cache")
	}
	Put(key, []byte("parsed"))
	if b, ok := Get(key); !ok || string(b) != "parsed" {
		t.Fatalf("Get = %q, %v", b, ok)
	}

	// A damaged entry is a miss, and goes.
	p, _ := entryPath(key)
	data, _ := os.ReadFile(p)
	data[len(data)-1] ^= 1
	os.WriteFile(p, data, 0o600)
	if _, ok := Get(key); ok {
		t.Error("damaged entry served")
	}
	if _, err := os.Stat(p); !os.IsNotExist(err) {
		t.Error("damaged entry kept")
	}
}

func TestPrune(t *testing.T) {
	t.Setenv("SHEETS_CACHE_DIR", t.TempDir())
	var keys []string
	for i := range MaxEntries + 3 {
		k := Key("xlsx", []byte(strconv.Itoa(i)))
		keys = append(keys, k)
		Put(k, bytes.Repeat([]byte{1}, 10))
		// Distinct use times, oldest first.
		p, _ := entryPath(k)
		at := time.Now().Add(time.Duration(i-100) * time.Minute)
		os.Chtimes(p, at, at)
	}
	// The three oldest went as the cache filled; using the oldest left
	// saves it from the next prune.
	if _, ok := Get(keys[2]); ok {
		t.Error("entry past the limit kept")
	}
	Get(keys[3])
	Put(Key("xlsx", []byte("last")), []byte("x"))
	dir, _ := Dir(parsed)
	ents, _ := filepath.Glob(filepath.Join(dir, "*.bin"))
	if len(ents) != MaxEntries {
		t.Errorf("%d entries, want %d", len(ents), MaxEntries)
	}
	if _, ok := Get(keys[3]); !ok {
		t.Error("recently used entry pruned")
	}
	if _, ok := Get(keys[4]); ok {
		t.Error("least recently used entry kept")
	}
}

func TestDisabled(t *testing.T) {
	t.Setenv("SHEETS_CACHE", "off")
	if !Disabled() {
		t.Error("SHEETS_CACHE=off not honoured")
	}
}
