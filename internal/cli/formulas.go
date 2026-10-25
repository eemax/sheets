package cli

import (
	"fmt"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// formulaReport tells an agent what the formulas it just wrote evaluate to.
type formulaReport struct {
	Values      []formulaValue `json:"values,omitempty"`      // a few results, to sanity-check
	Errors      []formulaValue `json:"errors,omitempty"`      // #DIV/0!, #REF!, ...
	Unevaluated []formulaValue `json:"unevaluated,omitempty"` // the calculator cannot compute these; Excel will
	More        int            `json:"more,omitempty"`        // results evaluated but not listed
	Skipped     int            `json:"skipped,omitempty"`     // not evaluated, over the limit
}

type formulaValue struct {
	Ref   string `json:"ref"`
	Value any    `json:"value,omitempty"`
	Note  string `json:"note,omitempty"`
}

type cellRef struct {
	sheet    string
	row, col int
	spills   bool // a dynamic array: the value shown is its first element
}

// spills tells whether a formula returns an array, which spills: a
// function returning one (=SORT(A2:A9)), or arithmetic over a range
// (=A2:A9*2).
func spills(f string) bool {
	_, arr := engine.ArrayUse(f)
	return arr || formula.Spills(f)
}

const (
	maxEvaluate    = 2000
	maxListed      = 5
	maxErrorListed = 10
)

// evaluator gives what a cell holds after a change: Editor.Evaluate.
type evaluator interface {
	Evaluate(sheet string, row, col int) (v grid.Cell, note string, ok bool)
}

// evaluateFormulas evaluates written formulas. qualify prefixes refs with
// their sheet, for commands that may touch several sheets.
func evaluateFormulas(ed evaluator, cells []cellRef, qualify bool) *formulaReport {
	if len(cells) == 0 {
		return nil
	}
	rep := &formulaReport{}
	for i, c := range cells {
		if i == maxEvaluate {
			rep.Skipped = len(cells) - i
			break
		}
		ref := grid.CellName(c.row, c.col)
		if qualify {
			ref = grid.FormatRef(c.sheet, grid.Range{R1: c.row, C1: c.col, R2: c.row, C2: c.col})
		}
		v, note, ok := ed.Evaluate(c.sheet, c.row, c.col)
		switch {
		case !ok:
			rep.Unevaluated = append(rep.Unevaluated, formulaValue{Ref: ref, Note: note})
		case v.Kind == grid.Error:
			rep.Errors = append(rep.Errors, formulaValue{Ref: ref, Value: v.Text, Note: note})
		case len(rep.Values) < maxListed:
			fv := formulaValue{Ref: ref, Value: plainValue(v)}
			if c.spills {
				fv.Note = "spills"
			}
			rep.Values = append(rep.Values, fv)
		default:
			rep.More++
		}
	}
	return rep
}

func (a *app) printFormulaReport(r *formulaReport) {
	if r == nil {
		return
	}
	if len(r.Values) > 0 {
		var parts []string
		for _, f := range r.Values {
			p := f.Ref + " = " + showValue(f.Value)
			if f.Note != "" {
				p += " (" + f.Note + ")"
			}
			parts = append(parts, p)
		}
		line := "formulas: " + strings.Join(parts, ", ")
		if r.More > 0 {
			line += fmt.Sprintf(", … %d more", r.More)
		}
		a.printf("%s\n", line)
	}
	if len(r.Errors) > 0 {
		var parts []string
		for _, f := range r.Errors[:min(len(r.Errors), maxErrorListed)] {
			p := f.Ref + " " + fmt.Sprint(f.Value)
			if f.Note != "" {
				p += " (" + f.Note + ")"
			}
			parts = append(parts, p)
		}
		line := fmt.Sprintf("%s: %s", plural(len(r.Errors), "formula error"), strings.Join(parts, ", "))
		if len(r.Errors) > maxErrorListed {
			line += ", …"
		}
		a.printf("%s\n", line)
	}
	if len(r.Unevaluated) > 0 {
		// Group by reason: usually one function explains them all.
		byNote := map[string][]string{}
		for _, f := range r.Unevaluated {
			byNote[f.Note] = append(byNote[f.Note], f.Ref)
		}
		notes := make([]string, 0, len(byNote))
		for n := range byNote {
			notes = append(notes, n)
		}
		sort.Strings(notes)
		for _, n := range notes {
			refs := byNote[n]
			list := strings.Join(refs[:min(len(refs), maxListed)], ", ")
			if len(refs) > maxListed {
				list += fmt.Sprintf(", … %d more", len(refs)-maxListed)
			}
			a.printf("not evaluated: %s: %s\n", list, n)
		}
	}
	if r.Skipped > 0 {
		a.printf("%s not evaluated (limit %d per command)\n", plural(r.Skipped, "formula"), maxEvaluate)
	}
}

func (a *app) printWarnings(ws []string) {
	for _, w := range ws {
		a.printf("warning: %s\n", w)
	}
}
