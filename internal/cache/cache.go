// Package cache keeps what sheets stores between runs in the user cache
// directory ($SHEETS_CACHE_DIR overrides it): the undo journal, and parsed
// copies of large workbooks so that reading one again skips the parse.
package cache

import (
	"bytes"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"hash/crc32"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

// Dir returns the directory for one kind of cached data: name under
// $SHEETS_CACHE_DIR, or under sheets in the user cache directory.
func Dir(name string) (string, error) {
	if d := os.Getenv("SHEETS_CACHE_DIR"); d != "" {
		return filepath.Join(d, name), nil
	}
	d, err := os.UserCacheDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(d, "sheets", name), nil
}

// Limits of the parsed-workbook cache. Files smaller than MinSize parse in
// well under a tenth of a second and are not cached; past MaxEntries or
// MaxBytes the least recently used entries go.
const (
	MinSize    = 1 << 20
	MaxEntries = 32
	MaxBytes   = 2 << 30
)

// Disabled reports whether SHEETS_CACHE=off turned the parsed-workbook
// cache off.
func Disabled() bool { return strings.EqualFold(os.Getenv("SHEETS_CACHE"), "off") }

var (
	buildOnce sync.Once
	build     string
)

// buildID identifies the running sheets binary, so a new build never reads
// what an older one cached (its reader or model may differ). Empty when
// unknown, which turns the cache off.
func buildID() string {
	buildOnce.Do(func() {
		exe, err := os.Executable()
		if err != nil {
			return
		}
		fi, err := os.Stat(exe)
		if err != nil {
			return
		}
		build = strconv.FormatInt(fi.Size(), 10) + "-" + strconv.FormatInt(fi.ModTime().UnixNano(), 10)
	})
	return build
}

// Key names the entry for data read as kind by this build of sheets. It is
// empty when entries cannot be told apart safely.
func Key(kind string, data []byte) string {
	id := buildID()
	if id == "" {
		return ""
	}
	h := sha256.New()
	h.Write([]byte(kind + "\x00" + id + "\x00"))
	h.Write(data)
	return hex.EncodeToString(h.Sum(nil))
}

// An entry is magic, a CRC-32C of the content, then the content.
var magic = []byte("SHC1")

var crcTable = crc32.MakeTable(crc32.Castagnoli)

const parsed = "parsed"

func entryPath(key string) (string, error) {
	d, err := Dir(parsed)
	if err != nil {
		return "", err
	}
	return filepath.Join(d, key+".bin"), nil
}

// Get returns the content stored under key, if it is there and intact.
func Get(key string) ([]byte, bool) {
	if key == "" {
		return nil, false
	}
	p, err := entryPath(key)
	if err != nil {
		return nil, false
	}
	b, err := os.ReadFile(p)
	if err != nil {
		return nil, false
	}
	if len(b) < 8 || !bytes.Equal(b[:4], magic) || binary.LittleEndian.Uint32(b[4:8]) != crc32.Checksum(b[8:], crcTable) {
		os.Remove(p)
		return nil, false
	}
	now := time.Now()
	os.Chtimes(p, now, now) // recently used
	return b[8:], true
}

// Remove drops the entry under key, for content that turned out unusable.
func Remove(key string) {
	if p, err := entryPath(key); err == nil && key != "" {
		os.Remove(p)
	}
}

// Put stores content under key, then trims the cache to its limits. It is
// best effort: a cache that cannot be written is no cache.
func Put(key string, content []byte) {
	if key == "" {
		return
	}
	p, err := entryPath(key)
	if err != nil {
		return
	}
	dir := filepath.Dir(p)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return
	}
	f, err := os.CreateTemp(dir, ".tmp-*")
	if err != nil {
		return
	}
	head := make([]byte, 8)
	copy(head, magic)
	binary.LittleEndian.PutUint32(head[4:], crc32.Checksum(content, crcTable))
	_, err = f.Write(head)
	if err == nil {
		_, err = f.Write(content)
	}
	if cerr := f.Close(); err == nil {
		err = cerr
	}
	if err == nil {
		err = os.Rename(f.Name(), p)
	}
	if err != nil {
		os.Remove(f.Name())
		return
	}
	prune(dir)
}

// prune removes the least recently used entries past the limits, and
// temporary files a crashed run left behind.
func prune(dir string) {
	ents, err := os.ReadDir(dir)
	if err != nil {
		return
	}
	type entry struct {
		path string
		size int64
		used time.Time
	}
	var all []entry
	for _, e := range ents {
		fi, err := e.Info()
		if err != nil || !fi.Mode().IsRegular() {
			continue
		}
		p := filepath.Join(dir, e.Name())
		if strings.HasPrefix(e.Name(), ".tmp-") {
			if time.Since(fi.ModTime()) > time.Hour {
				os.Remove(p)
			}
			continue
		}
		if strings.HasSuffix(e.Name(), ".bin") {
			all = append(all, entry{p, fi.Size(), fi.ModTime()})
		}
	}
	sort.Slice(all, func(i, j int) bool { return all[i].used.After(all[j].used) })
	var total int64
	for i, e := range all {
		total += e.size
		if i >= MaxEntries || total > MaxBytes {
			os.Remove(e.path)
		}
	}
}
