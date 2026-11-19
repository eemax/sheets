package cli

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/xlsx"
)

// The features a command's own report finds lost: none that are still
// there, after an undo of what renamed or moved them (change-feedback f4.1)
// or a move of their cells to another sheet (f4.2).

// feedbackBook writes Data!A1:B3 (Name, Qty) with a validation on B2:B3, a
// conditional format there and a chart, and a sheet Other.
func feedbackBook(t *testing.T) (dir, path string) {
	t.Helper()
	dir = t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	path = filepath.Join(dir, "r.xlsx")
	for _, c := range [][]string{
		{"write", path, "Data!A1"},
		{"validate", path, "Data!B2:B3", "whole", ">= 0"},
		{"cond", path, "Data!B2:B3", "> 1"},
		{"chart", "add", path, "Data!A1:B3", "--at", "Data!D1"},
		{"sheet", "add", path, "Other"},
	} {
		stdin := ""
		if c[0] == "write" {
			stdin = "Name\tQty\nA\t1\nB\t3\n"
		}
		if out, code := runCLI(stdin, c...); code != 0 {
			t.Fatalf("%v: %s", c, out)
		}
	}
	return dir, path
}

// head is a command's change report's head line.
func head(t *testing.T, out string) string {
	t.Helper()
	for _, l := range strings.Split(out, "\n") {
		if strings.HasPrefix(l, "change: ") {
			return l
		}
	}
	t.Fatalf("no change report in:\n%s", out)
	return ""
}

// Undoing a sheet rename, a column inserted, rows inserted: the features
// are back where they were, none lost.
func TestFeedbackFeaturesUndo(t *testing.T) {
	for _, edit := range [][]string{
		{"sheet", "rename", "@", "Data", "Facts"},
		{"cols", "insert", "@", "Data!A"},
		{"rows", "insert", "@", "Data!2", "-n", "3"},
		{"sheet", "move", "@", "Other", "--before", "Data"},
	} {
		_, path := feedbackBook(t)
		args := append([]string{}, edit...)
		for i, a := range args {
			if a == "@" {
				args[i] = path
			}
		}
		if out, code := runCLI("", args...); code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
			t.Fatalf("%v: exit %d:\n%s", edit, code, out)
		}
		out, code := runCLI("", "undo", path)
		if code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
			t.Errorf("undo of %v: exit %d:\n%s", edit, code, out)
		}
	}
	// An undo that does take one away says so.
	_, path := feedbackBook(t)
	if out, code := runCLI("", "validate", path, "Data!A2:A3", "list", "A,B"); code != 0 {
		t.Fatal(out)
	}
	if out, _ := runCLI("", "undo", path); !strings.Contains(head(t, out), ", 1 feature lost") {
		t.Errorf("undo of a validation added:\n%s", out)
	}
}

// Cells moved to another sheet, and back, in both editors: their
// validation and conditional format went with them, none lost.
func TestFeedbackFeaturesMoveAcross(t *testing.T) {
	for _, stream := range []string{"never", "always"} {
		t.Setenv("SHEETS_STREAM", stream)
		_, path := feedbackBook(t)
		for _, mv := range [][2]string{{"Data!A1:B3", "Other!C3"}, {"Other!C3:D5", "Other!F10"}, {"Other!F10:G12", "Data!A1"}} {
			out, code := runCLI("", "move", path, mv[0], mv[1])
			if code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
				t.Errorf("%s: move %s %s: exit %d:\n%s", stream, mv[0], mv[1], code, out)
			}
		}
		out, _ := runCLI("", "validate", path, "Data", "--list")
		if !strings.Contains(out, "Data!B2:B3") {
			t.Errorf("%s: the validation is not back: %s", stream, out)
		}
		// In a batch, the move and a set after it: one report, nothing lost.
		out, code := runCLI("move Data!A1:B3 Other!A1\nset Other!B2 5\n", "do", path)
		if code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
			t.Errorf("%s: batch: exit %d:\n%s", stream, code, out)
		}
	}
}

// A move of a table to a new sheet, then its undo: the chart that stayed on
// its sheet, still there, is not lost (f0.4). The pairing of the versions by
// what the sheets hold takes the cells cut and pasted across sheets for a
// sheet renamed, which moves the chart that stayed; the session's own way
// finds it where it was. Undoing a rename or an insert still loses nothing.
func TestFeedbackFeaturesUndoMoveAcross(t *testing.T) {
	for _, stream := range []string{"never", "always"} {
		t.Setenv("SHEETS_STREAM", stream)
		_, path := causesFixture(t)
		if out, code := runCLI("", "move", path, "Data!A1:D7", "Other!A1"); code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
			t.Fatalf("%s: move: exit %d:\n%s", stream, code, out)
		}
		out, code := runCLI("", "undo", path)
		if code != 0 || strings.Contains(out, "chart") {
			t.Errorf("%s: undo of the move: exit %d:\n%s", stream, code, out)
		}
		if info, _ := runCLI("", "info", path); !strings.Contains(info, "F2  column") {
			t.Errorf("%s: the chart is not back:\n%s", stream, info)
		}
	}
}

// Undoing a move of a table to another sheet (f2.6): the validations that
// went with it are back on the cells they were on, so none is lost, in both
// editors. The undo does not know what it undid, so a rule not found where it
// was is found by what it is.
func TestFeedbackFeaturesUndoMoveValidations(t *testing.T) {
	for _, stream := range []string{"never", "always"} {
		t.Setenv("SHEETS_STREAM", stream)
		dir, path := causesFixture(t)
		orig := filepath.Join(dir, "orig.xlsx")
		if err := os.WriteFile(orig, mustRead(t, path), 0o644); err != nil {
			t.Fatal(err)
		}
		if out, code := runCLI("", "move", path, "Data!A1:D7", "Other!A1"); code != 0 || !strings.Contains(head(t, out), ", 0 features lost") {
			t.Fatalf("%s: move: exit %d:\n%s", stream, code, out)
		}
		out, code := runCLI("", "undo", path)
		if code != 0 || !strings.Contains(head(t, out), ", 0 features lost") || strings.Contains(out, "\nlost:") {
			t.Errorf("%s: undo of the move: exit %d:\n%s", stream, code, out)
		}
		if got, _ := runCLI("", "validate", path, "Data", "--list"); !strings.Contains(got, "Data!B2:B6") || !strings.Contains(got, "Data!C2:C6") {
			t.Errorf("%s: the validations are not back: %s", stream, got)
		}
		if res, code := against(t, path, orig); code != 0 || res.Summary.FeaturesLost != 0 {
			t.Errorf("%s: the file against the original: exit %d, %+v", stream, code, res.Summary)
		}
	}
}

// What the undo finds elsewhere it finds one for one: a rule the file lost and
// another of the same kind, type and rule on another sheet it did not have
// before are the one moved; two rules lost, one found, is one lost.
func TestFeedbackFeaturesAcross(t *testing.T) {
	rule := func(sheet, where, what, typ, rule string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureValidation, Sheet: sheet, Where: where, What: what, Type: typ, Rule: rule}
	}
	before := []xlsx.Feature{
		rule("Other", "B2:B6", "whole validation", "whole", "between|0|100"),
		rule("Other", "C2:C6", "decimal validation", "decimal", "greaterThanOrEqual|0"),
	}
	after := []xlsx.Feature{rule("Data", "B2:B6", "whole validation", "whole", "between|0|100")}
	if lost, _ := compareFeaturesAcross(before, after, nil, nil, nil, false); len(lost) != 2 {
		t.Errorf("not looked for across sheets: %+v", lost)
	}
	lost, _ := compareFeaturesAcross(before, after, nil, nil, nil, true)
	if len(lost) != 1 || lost[0].Where != "Other!C2:C6" {
		t.Errorf("looked for across sheets: %+v", lost)
	}
	// One found on the sheet it was on stays that one's: the other has no match.
	both := []xlsx.Feature{
		rule("Other", "B2:B6", "whole validation", "whole", "between|0|100"),
		rule("Data", "B2:B6", "whole validation", "whole", "between|0|100"),
	}
	if lost, _ := compareFeaturesAcross(both, []xlsx.Feature{both[1]}, nil, nil, nil, true); len(lost) != 1 {
		t.Errorf("two rules, one left: %+v", lost)
	}
}
