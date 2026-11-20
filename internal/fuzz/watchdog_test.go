package fuzz

import (
	"io"
	"os"
	"strings"
	"testing"
	"time"
)

// TestHangIsAFinding: an edit that does not finish within editTimeout is a
// finding of the invariant hang, with the seed, the workbook and the edits
// that led to it, and the run stops there. (Before, a command that hung ran
// until go test's own timeout, which printed a dump of goroutines and no
// workbook, no edits and no seed.)
func TestHangIsAFinding(t *testing.T) {
	oldTimeout, oldMain := editTimeout, cliMain
	release, entered := make(chan struct{}), make(chan struct{}, 1)
	t.Cleanup(func() {
		select {
		case <-entered: // the hung command has read cliMain: it is not read again
		default: // no command hung
		}
		close(release) // and ends
		editTimeout, cliMain = oldTimeout, oldMain
	})
	// Longer than an edit takes under -race and a busy machine, short enough
	// for a test.
	editTimeout = 1500 * time.Millisecond
	// A clear never returns, as a command that loops would not.
	cliMain = func(args []string, stdin io.Reader, out, errw io.Writer) int {
		if len(args) > 0 && args[0] == "clear" {
			entered <- struct{}{}
			<-release
			return 0
		}
		return oldMain(args, stdin, out, errw)
	}
	pool := loadPool(t)
	start := time.Now()
	res, err := Run(pool, Options{Seed: 1, Edits: 2000, Dir: t.TempDir()})
	if err != nil {
		t.Fatal(err)
	}
	if !res.Hung || len(res.Findings) == 0 {
		t.Fatalf("the run did not stop at a hang: %d edits, %d findings", res.Edits, len(res.Findings))
	}
	if res.Edits >= 2000 || time.Since(start) > 30*time.Second {
		t.Errorf("the run went on after the hang: %d edits in %s", res.Edits, time.Since(start))
	}
	f := res.Findings[len(res.Findings)-1]
	if f.Invariant != "hang" || f.Class != "clear" {
		t.Fatalf("the finding is %s, want a hang of a clear", f.Signature)
	}
	if last := f.Edits[len(f.Edits)-1]; last.Args[0] != "clear" {
		t.Errorf("the sequence should end at the edit that hung, a clear: %s", last)
	}
	report := f.Report()
	for _, want := range []string{"finding clear hang " + strings.Fields(f.Signature.String())[2], "workbook: testdata/pool/" + f.Workbook, "sheets clear book.xlsx", "did not finish in 1.5s"} {
		if !strings.Contains(report, want) {
			t.Errorf("the report has no %q:\n%s", want, report)
		}
	}
	// A hang is not shrunk (every replay would wait out the timeout).
	g, err := Shrink(pool, f, t.TempDir())
	if err != nil || g.Report() != report {
		t.Errorf("a hang was shrunk: %v\n%s", err, g.Report())
	}
	// And known.txt cannot list one: it would end the run each time.
	if _, err := ParseKnown(f.Signature.String() + "\n"); err == nil {
		t.Error("known.txt accepted a hang")
	}
	t.Log(report)
}

// TestSlowEditIsNoHang: an edit that takes less than editTimeout is judged as
// before.
func TestSlowEditIsNoHang(t *testing.T) {
	oldTimeout, oldMain := editTimeout, cliMain
	t.Cleanup(func() { editTimeout, cliMain = oldTimeout, oldMain })
	editTimeout = 5 * time.Second
	cliMain = func(args []string, stdin io.Reader, out, errw io.Writer) int {
		time.Sleep(100 * time.Millisecond)
		return oldMain(args, stdin, out, errw)
	}
	pool := loadPool(t)
	s, err := newSession(t.TempDir(), pool[0].Data)
	if err != nil {
		t.Fatal(err)
	}
	defer s.close()
	o, err := s.step(Edit{Class: "set", Args: []string{"set", FileArg, "Data!B2", "5"}}, 0)
	if err != nil || len(o.broken) > 0 || !o.applied {
		t.Fatalf("a slow edit: %+v, %v", o, err)
	}
	if _, err := os.Stat(s.paths[0]); err != nil {
		t.Fatal(err)
	}
}
