package fuzz

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// TestMutantsApply: each mutant (mutants/*.patch: a wave-1 bug put back into
// sheets, for scripts/fuzz-mutants.sh) still applies to the tree as it is, so
// a change to the code it patches, which would leave the mutant stale and the
// fuzzer's check of itself silently weaker, fails the suite. A stale mutant
// is written again against the new code, putting back the same bug (README,
// How the fuzzer was checked). Each says first what it puts back, and
// patches sheets' code, never a test or the fuzzer.
func TestMutantsApply(t *testing.T) {
	git, err := exec.LookPath("git")
	if err != nil {
		t.Fatal("git applies the mutants: not found")
	}
	patches, err := filepath.Glob(filepath.Join("mutants", "*.patch"))
	if err != nil {
		t.Fatal(err)
	}
	if len(patches) != 7 {
		t.Errorf("%d mutants, want the 7 of wave 1's audit: %v", len(patches), patches)
	}
	for _, p := range patches {
		name := strings.TrimSuffix(filepath.Base(p), ".patch")
		t.Run(name, func(t *testing.T) {
			text, err := os.ReadFile(p)
			if err != nil {
				t.Fatal(err)
			}
			head, body, ok := strings.Cut(string(text), "\n--- a/")
			if !ok || strings.TrimSpace(head) == "" {
				t.Fatalf("no description before the patch, or no patch")
			}
			for _, line := range strings.Split("--- a/"+body, "\n") {
				if f, ok := strings.CutPrefix(line, "+++ b/"); ok {
					if !strings.HasPrefix(f, "internal/") || strings.HasSuffix(f, "_test.go") || strings.HasPrefix(f, "internal/fuzz/") {
						t.Errorf("patches %s: a mutant patches sheets' code, not a test or the fuzzer", f)
					}
				}
			}
			abs, err := filepath.Abs(p)
			if err != nil {
				t.Fatal(err)
			}
			cmd := exec.Command(git, "apply", "--check", abs)
			cmd.Dir = repoRoot
			if out, err := cmd.CombinedOutput(); err != nil {
				t.Errorf("does not apply to the tree (stale: write it again against the new code, putting back the same bug): %v\n%s", err, out)
			}
		})
	}
}

// TestMutantsApplyRefuses: the check refuses a mutant whose code has changed
// under it (here, the code as the mutant leaves it: a patch applied twice),
// and passes it on the code it was written against.
func TestMutantsApplyRefuses(t *testing.T) {
	git, err := exec.LookPath("git")
	if err != nil {
		t.Fatal("git applies the mutants: not found")
	}
	p, err := filepath.Abs(filepath.Join("mutants", "rewrite-loses-kind.patch"))
	if err != nil {
		t.Fatal(err)
	}
	const file = "internal/xlsx/edit.go"
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, filepath.Dir(file)), 0o755); err != nil {
		t.Fatal(err)
	}
	src, err := os.ReadFile(filepath.Join(repoRoot, file))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, file), src, 0o644); err != nil {
		t.Fatal(err)
	}
	apply := func(args ...string) error {
		cmd := exec.Command(git, append([]string{"apply"}, append(args, p)...)...)
		cmd.Dir = dir
		return cmd.Run()
	}
	if err := apply("--check"); err != nil {
		t.Fatalf("the mutant does not apply to the code it was written against: %v", err)
	}
	if err := apply(); err != nil {
		t.Fatal(err)
	}
	if err := apply("--check"); err == nil {
		t.Error("a mutant applied to code that changed under it (stale) passed the check")
	}
}

