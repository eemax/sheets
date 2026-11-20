package cli

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/render"
	"github.com/eemax/sheets/internal/xlsx"
)

// check --against --receipt: the report of check --against, written for a
// person as Markdown, with before and after pictures of the ranges that
// changed (sheets render), beside the receipt. It is built from the
// againstResult alone: nothing is compared again.

// receiptInfo is what --json says of the receipt written.
type receiptInfo struct {
	File     string           `json:"file"`
	Previews []receiptPreview `json:"previews"`
	// Unshown are the changes outside the previews (there are at most
	// receiptPreviews of them).
	Unshown int `json:"not_previewed,omitempty"`
}

type receiptPreview struct {
	Before      string `json:"before"` // Sheet!A1:D7 in the earlier version, "" when not there
	After       string `json:"after"`  // and in the file checked
	BeforeImage string `json:"before_image,omitempty"`
	AfterImage  string `json:"after_image,omitempty"`
}

// receiptPreviews bounds the ranges pictured; receiptRows and receiptCols
// what one range spans (render clips what is larger, and says so);
// receiptListed the entries each section lists.
const (
	receiptPreviews = 3
	receiptRows     = 30
	receiptCols     = 12
	receiptListed   = 25
)

// spot is a cell a change is at, in both versions.
type spot struct {
	after, before place
	kind          string // what it is, for the preview's caption
	unknown       bool
}

type place struct {
	sheet    string
	row, col int
}

func (p place) ok() bool { return p.sheet != "" }

// region is one preview: the cells it shows in each version.
type region struct {
	sheet, beforeSheet string
	after, before      grid.Range
	kinds              map[string]*kindCells
	unknown            []string
}

// kindCells are a region's changes of one kind: their count, and the first
// few cells.
type kindCells struct {
	n     int
	cells []string
}

// at is the one cell a place is, or nothing when the version has no such
// sheet.
func (p place) at() grid.Range {
	if !p.ok() {
		return grid.Range{R1: 1, C1: 1}
	}
	return grid.Range{R1: p.row, C1: p.col, R2: p.row, C2: p.col}
}

func (g *region) fits(s spot) bool {
	if !strings.EqualFold(g.sheet, s.after.sheet) || !strings.EqualFold(g.beforeSheet, s.before.sheet) {
		return false
	}
	a, b := g.after.Union(s.after.at()), g.before.Union(s.before.at())
	return a.Rows() <= receiptRows && a.Cols() <= receiptCols && b.Rows() <= receiptRows && b.Cols() <= receiptCols
}

func (g *region) add(s spot) {
	g.after, g.before = g.after.Union(s.after.at()), g.before.Union(s.before.at())
	k := g.kinds[s.kind]
	if k == nil {
		k = &kindCells{}
		g.kinds[s.kind] = k
	}
	k.n++
	if len(k.cells) < causeListed {
		if s.after.ok() {
			k.cells = append(k.cells, grid.CellName(s.after.row, s.after.col))
		} else {
			k.cells = append(k.cells, grid.CellName(s.before.row, s.before.col)+" before")
		}
	}
	if s.unknown && len(g.unknown) < causeListed {
		g.unknown = append(g.unknown, grid.CellName(s.after.row, s.after.col))
	}
}

// framed widens a range by a row and a column around it, and to row 1 or
// column A when they are near (the labels a table starts with), within
// what one preview spans.
func framed(rg grid.Range) grid.Range {
	if rg.Empty() {
		return rg
	}
	rg = grid.Range{R1: max(rg.R1-1, 1), C1: max(rg.C1-1, 1), R2: rg.R2 + 1, C2: rg.C2 + 1}
	if rg.R1 <= 4 && rg.R2 < receiptRows+2 {
		rg.R1 = 1
	}
	if rg.C1 <= 3 && rg.C2 < receiptCols+2 {
		rg.C1 = 1
	}
	return rg
}

// receiptSpots are the places of what changed, the most telling first:
// values unexplained, charts lost, flagged formulas, inputs and formulas
// changed, values moved, values unknown. A feature that no picture shows
// goes in unshown instead.
func (r *againstResult) receiptSpots() (spots []spot, unshown []string) {
	A, B := r.sides[0], r.sides[1]
	isA := func(n string) bool { return A.wb.Sheet(n) != nil }
	isB := func(n string) bool { return B.wb.Sheet(n) != nil }
	at := func(ref string, isSheet func(string) bool, sheet string) place {
		p, err := grid.ParseRef(ref, isSheet)
		if err != nil || p.Range == nil {
			return place{}
		}
		if p.Sheet == "" {
			p.Sheet = sheet
		}
		if !isSheet(p.Sheet) {
			return place{}
		}
		return place{sheet: p.Sheet, row: p.Range.R1, col: p.Range.C1}
	}
	// oldSheet is a sheet's name before; newSheet after.
	oldSheet := func(name string) string {
		for old, now := range r.renamed {
			if strings.EqualFold(now, name) {
				for _, s := range A.wb.Sheets {
					if strings.EqualFold(s.Name, old) {
						return s.Name
					}
				}
			}
		}
		if isA(name) {
			return name
		}
		return ""
	}
	newSheet := func(name string) string {
		if now, ok := r.renamed[strings.ToLower(name)]; ok {
			return now
		}
		if isB(name) {
			return name
		}
		return ""
	}
	cell := func(c, was, kind string, unknown bool) {
		s := spot{kind: kind, unknown: unknown}
		s.after = at(c, isB, "")
		if !s.after.ok() {
			return
		}
		if was == "" {
			if old := oldSheet(s.after.sheet); old != "" {
				s.before = place{sheet: old, row: s.after.row, col: s.after.col}
			}
		} else {
			s.before = at(was, isA, oldSheet(s.after.sheet))
		}
		spots = append(spots, s)
	}
	for _, m := range r.Unexplained {
		cell(m.Cell, m.Was, "value unexplained", false)
	}
	for _, f := range r.FeaturesLost {
		where := f.Kind + " " + f.Where
		if f.What != "" {
			where += " (" + f.What + ")"
		}
		switch {
		case f.Kind == xlsx.FeatureChart && strings.Contains(f.Where, "!"):
			s := spot{kind: "chart lost"}
			first, _, _ := strings.Cut(f.Where, ":")
			if s.before = at(first, isA, ""); !s.before.ok() {
				unshown = append(unshown, where+": not found to draw")
				continue
			}
			// Where the chart was, in the sheet it was on; gone with its
			// sheet, the picture before is all there is.
			if now := newSheet(s.before.sheet); now != "" {
				s.after = place{sheet: now, row: s.before.row, col: s.before.col}
			}
			spots = append(spots, s)
		case f.Kind == xlsx.FeatureChart:
			unshown = append(unshown, where+": a chart sheet is not drawn")
		case f.Kind == xlsx.FeatureShape || f.Kind == xlsx.FeaturePicture:
			unshown = append(unshown, where+": sheets render draws no shapes or pictures")
		case f.Kind == xlsx.FeatureValidation:
			unshown = append(unshown, where+": a validation shows in no picture")
		case f.Kind == xlsx.FeatureCond:
			unshown = append(unshown, where+": sheets render draws no conditional formats but banding, so a picture cannot show it gone")
		case f.Kind == xlsx.FeatureTable:
			unshown = append(unshown, where+": sheets render draws no table styles")
		default:
			unshown = append(unshown, where+": a name shows in no picture")
		}
	}
	for _, c := range r.Changed {
		if c.Flag != "" {
			cell(c.Cell, c.Was, "formula flagged", false)
		}
	}
	for _, c := range r.Changed {
		if c.Flag == "" {
			cell(c.Cell, c.Was, c.Kind+" changed", false)
		}
	}
	for _, m := range r.Moved {
		cell(m.Cell, m.Was, "value moved", false)
	}
	for _, u := range r.Unknown {
		cell(u.Cell, u.Was, "value unknown", true)
	}
	return spots, unshown
}

// regions gathers spots into at most receiptPreviews ranges; left counts
// the spots outside them, and where names a few. A sheet gets a range of
// its own before another gets a second, so a total on a sheet of its own
// is pictured beside the rows that moved it.
func regions(spots []spot) (out []*region, left int, where []string) {
	placed := make([]bool, len(spots))
	for pass := 0; pass < 2; pass++ {
		for i, s := range spots {
			if placed[i] {
				continue
			}
			for _, g := range out {
				if g.fits(s) {
					g.add(s)
					placed[i] = true
					break
				}
			}
			if placed[i] || len(out) == receiptPreviews {
				continue
			}
			if pass == 0 && slices.ContainsFunc(out, func(g *region) bool { return strings.EqualFold(g.sheet, s.after.sheet) }) {
				continue
			}
			g := &region{sheet: s.after.sheet, beforeSheet: s.before.sheet, after: grid.Range{R1: 1, C1: 1}, before: grid.Range{R1: 1, C1: 1}, kinds: map[string]*kindCells{}}
			g.add(s)
			out = append(out, g)
			placed[i] = true
		}
	}
	for i, s := range spots {
		if placed[i] {
			continue
		}
		left++
		if len(where) < causeListed {
			if s.after.ok() {
				where = append(where, cellAt(s.after.sheet, s.after.row, s.after.col))
			} else {
				where = append(where, cellAt(s.before.sheet, s.before.row, s.before.col)+" in the earlier version")
			}
		}
	}
	for _, g := range out {
		g.after, g.before = framed(g.after), framed(g.before)
	}
	return out, left, where
}

// writeReceipt writes the receipt to path, and its pictures beside it.
func (a *app) writeReceipt(path string, r *againstResult) (*receiptInfo, error) {
	if fi, err := os.Stat(path); err == nil {
		for _, f := range []string{r.File, r.Against} {
			if g, err := os.Stat(f); err == nil && os.SameFile(fi, g) {
				return nil, fmt.Errorf("--receipt %s is %s, which it would overwrite: name a .md file", path, f)
			}
		}
	}
	info := &receiptInfo{File: path, Previews: []receiptPreview{}}
	base := strings.TrimSuffix(path, filepath.Ext(path))
	var b bytes.Buffer
	w := func(format string, args ...any) { fmt.Fprintf(&b, format, args...) }

	w("# Change receipt: %s\n\n", mdCode(r.File))
	w("What changed in %s since %s, and why, as `sheets check --against` found it.\n\n", mdCode(r.File), mdCode(r.Against))
	w("**%s**\n\n", receiptVerdict(r))
	s := r.Summary
	w("| What | Count |\n|---|---:|\n")
	if r.ValuesOnly {
		w("| Values changed | %d |\n", s.Inputs)
	} else {
		w("| Inputs changed | %d |\n| Formulas changed | %d |\n| Formula changes flagged as likely mistakes | %d |\n", s.Inputs, s.Formulas, s.Flagged)
		w("| Computed values moved, each with its cause | %d |\n| Values moved with nothing upstream changed (unexplained) | %d |\n", s.Moved, s.Unexplained)
		w("| Features lost | %d |\n", s.FeaturesLost)
		if s.FeaturesChanged > 0 {
			w("| Features changed (rules edited on their cells; not lost) | %d |\n", s.FeaturesChanged)
		}
		w("| Values sheets cannot compare (unknown) | %d |\n", s.Unknown)
	}
	w("| Rows moved (as by a sort; not changes) | %d |\n", s.MovedRows)
	w("| Check of the file: errors, warnings, notes | %d, %d, %d |\n\n", r.Errors, r.Warnings, r.Notes)
	if r.ValuesOnly {
		w("A CSV file holds values only: formulas, causes and features are not compared.\n\n")
	}
	all := fmt.Sprintf("`sheets --json check %s --against %s` lists them all", shellQuote(r.File), shellQuote(r.Against))
	more := func(n int) {
		if n > receiptListed {
			w("\n… and %s more: %s.\n", grid.FormatNumber(float64(n-receiptListed)), all)
		}
		w("\n")
	}
	if n := len(r.Unexplained); n > 0 {
		w("## Unexplained: values that moved with nothing upstream changed\n\n| Cell | Before | After | Note |\n|---|---|---|---|\n")
		for _, m := range r.Unexplained[:min(n, receiptListed)] {
			w("| %s | %s | %s | %s |\n", mdCode(withWas(m.Cell, m.Was)), mdCode(clip(m.Before)), mdCode(clip(m.After)), mdText(clipNote(unexplainedNote(m))))
		}
		more(n)
	}
	if n := len(r.FeaturesLost); n > 0 {
		w("## Features lost\n\n| Kind | Where | What |\n|---|---|---|\n")
		for _, f := range r.FeaturesLost[:min(n, receiptListed)] {
			w("| %s | %s | %s |\n", mdText(strings.ReplaceAll(f.Kind, "_", " ")), mdCode(clip(f.Where)), mdText(clipNote(f.What)))
		}
		more(n)
	}
	if n := len(r.FeaturesChanged); n > 0 {
		w("## Features changed: rules edited on the cells they cover (not lost)\n\n| Kind | Where | Before | After |\n|---|---|---|---|\n")
		for _, f := range r.FeaturesChanged[:min(n, receiptListed)] {
			w("| %s | %s | %s | %s |\n", mdText(strings.ReplaceAll(f.Kind, "_", " ")), mdCode(clip(f.Where)), mdCode(clip(f.Before)), mdCode(clip(f.After)))
		}
		more(n)
	}
	var flagged, plain []changedCell
	for _, c := range r.Changed {
		if c.Flag != "" {
			flagged = append(flagged, c)
		} else {
			plain = append(plain, c)
		}
	}
	if n := len(flagged); n > 0 {
		w("## Flagged: formula changes that look like mistakes\n\n| Cell | Before | After | Why |\n|---|---|---|---|\n")
		for _, c := range flagged[:min(n, receiptListed)] {
			w("| %s | %s | %s | %s |\n", mdCode(withWas(c.Cell, c.Was)), mdCode(clip(c.Before)), mdCode(clip(c.After)), mdText(clipNote(c.Flag)))
		}
		more(n)
	}
	if n := len(plain); n > 0 {
		w("## Changed\n\n| Cell | Kind | Before | After |\n|---|---|---|---|\n")
		for _, c := range plain[:min(n, receiptListed)] {
			w("| %s | %s | %s | %s |\n", mdCode(withWas(c.Cell, c.Was)), c.Kind, mdCode(clip(c.Before)), mdCode(clip(c.After)))
		}
		more(n)
	}
	if n := len(r.Moved); n > 0 {
		w("## Values that moved, and why\n\n| Cell | Before | After | Cause |\n|---|---|---|---|\n")
		for _, m := range r.Moved[:min(n, receiptListed)] {
			cause := mdText(clipNote(m.Cause))
			if m.Note != "" {
				cause += " (" + mdText(clipNote(m.Note)) + ")"
			}
			w("| %s | %s | %s | %s |\n", mdCode(withWas(m.Cell, m.Was)), mdCode(clip(m.Before)), mdCode(clip(m.After)), cause)
		}
		more(n)
	}
	if n := len(r.Unknown); n > 0 {
		w("## Unknown: values sheets cannot compare\n\n| Cell | Before | After | Why |\n|---|---|---|---|\n")
		for _, u := range r.Unknown[:min(n, receiptListed)] {
			w("| %s | %s | %s | %s |\n", mdCode(withWas(u.Cell, u.Was)), mdCode(clip(u.Before)), mdCode(clip(u.After)), mdText(clipNote(u.Why)))
		}
		more(n)
	}
	if n := len(r.MovedRows); n > 0 {
		w("## Rows moved (as by a sort; not changes)\n\n")
		for _, m := range r.MovedRows[:min(n, receiptListed)] {
			w("- %s, were %s%s\n", mdCode(movedRange(m.Sheet, m.To, m.To+m.Rows-1, m.c1, m.c2)),
				mdCode(strings.TrimPrefix(movedRange(m.Sheet, m.From, m.From+m.Rows-1, m.wasC1, m.wasC2), grid.QuoteSheet(m.Sheet)+"!")), mdText(heldNumber(m)))
		}
		more(n)
	}
	if n := len(r.Structure); n > 0 {
		w("## Structure\n\n")
		for _, e := range r.Structure[:min(n, receiptListed)] {
			w("- %s: %s\n", e.Change, mdCode(clipNote(e.Where)))
		}
		more(n)
	}
	w("## Check of %s\n\n", mdCode(r.File))
	if n := len(r.Findings); n == 0 {
		w("No problems found.\n\n")
	} else {
		w("%s, %s, %s.\n\n| Severity | Cells | Finding |\n|---|---|---|\n", plural(r.Errors, "error"), plural(r.Warnings, "warning"), plural(r.Notes, "note"))
		fs := append([]checkFinding{}, r.Findings...)
		sort.SliceStable(fs, func(i, j int) bool { return fs[i].Severity > fs[j].Severity })
		for _, f := range fs[:min(n, receiptListed)] {
			w("| %s | %s | %s |\n", f.Severity, mdCode(clip(findingCells(f))), mdText(grid.Truncate(f.Message, againstFinding)))
		}
		more(n)
	}

	w("## Before and after\n\n")
	if err := a.receiptPreviews(&b, base, r, info); err != nil {
		return nil, err
	}
	if err := writeFileAtomic(path, b.Bytes()); err != nil {
		return nil, err
	}
	return info, nil
}

// receiptVerdict is the receipt's first line: what a person must look at.
func receiptVerdict(r *againstResult) string {
	var look []string
	s := r.Summary
	if s.Unexplained > 0 {
		look = append(look, countOf(s.Unexplained, "value")+" moved with nothing upstream changed")
	}
	if s.FeaturesLost > 0 {
		look = append(look, countOf(s.FeaturesLost, "feature")+" lost")
	}
	if s.Flagged > 0 {
		look = append(look, countOf(s.Flagged, "formula change")+" flagged")
	}
	if r.Errors > 0 {
		look = append(look, plural(r.Errors, "error")+" in the check of the file")
	}
	if len(look) > 0 {
		return "To look at: " + strings.Join(look, "; ") + "."
	}
	if s.Unknown > 0 {
		return fmt.Sprintf("Every value sheets can compare that moved has its cause; %s cannot be compared.", countOf(s.Unknown, "value"))
	}
	if s.Inputs+s.Formulas+s.Moved+s.MovedRows+s.FeaturesChanged+len(r.Structure) == 0 {
		return "No changes."
	}
	return "Every value that moved has its cause; no feature was lost."
}

// receiptPreviews writes the before and after section and its pictures.
func (a *app) receiptPreviews(b *bytes.Buffer, base string, r *againstResult, info *receiptInfo) error {
	w := func(format string, args ...any) { fmt.Fprintf(b, format, args...) }
	// Pictures an earlier receipt left under these names would be read as
	// this one's.
	for i := 1; i <= receiptPreviews; i++ {
		for _, side := range []string{"before", "after"} {
			os.Remove(fmt.Sprintf("%s-%d-%s.png", base, i, side))
		}
	}
	if r.ValuesOnly {
		w("A CSV file has no formatting to draw: no pictures.\n")
		return nil
	}
	spots, unshown := r.receiptSpots()
	regs, left, where := regions(spots)
	if len(regs) == 0 {
		w("Nothing changed that a picture shows.\n\n")
	}
	if len(regs) > 0 {
		w("The pictures are drawn by `sheets render`, from each version with the results sheets computed where the file stores none, as the tables above give them.\n\n")
	}
	dir, err := os.MkdirTemp("", "sheets-receipt-")
	if err != nil {
		return err
	}
	defer os.RemoveAll(dir)
	files := [2]string{}
	for i, side := range r.sides {
		files[i] = side.path
		if side.fixed != nil {
			files[i] = filepath.Join(dir, fmt.Sprintf("%d.xlsx", i))
			if err := os.WriteFile(files[i], side.fixed, 0o600); err != nil {
				return err
			}
		}
	}
	for i, g := range regs {
		n := i + 1
		p := receiptPreview{}
		var notes []string
		var cells [2]string
		for k, side := range []struct {
			sheet, label string
			rg           grid.Range
			file, shown  string
			image        *string
			ref          *string
		}{
			{g.beforeSheet, "before", g.before, files[0], r.Against, &p.BeforeImage, &p.Before},
			{g.sheet, "after", g.after, files[1], r.File, &p.AfterImage, &p.After},
		} {
			if side.sheet == "" {
				cells[k] = fmt.Sprintf("*not in %s*", mdCode(side.shown))
				continue
			}
			*side.ref = grid.FormatRef(side.sheet, side.rg)
			img := fmt.Sprintf("%s-%d-%s.png", base, n, side.label)
			res, err := renderTo(side.file, img, render.Options{Sheet: side.sheet, Range: &side.rg, Scale: 2, Headers: true, MaxRows: receiptRows + 4, MaxCols: receiptCols + 4})
			if err != nil {
				cells[k] = "*could not be drawn: " + mdText(err.Error()) + "*"
				continue
			}
			*side.image = img
			cells[k] = fmt.Sprintf("![%s %s](%s)", mdText(*side.ref), side.label, mdLink(filepath.Base(img)))
			if res.Of != "" {
				notes = append(notes, fmt.Sprintf("%s: only %s of %s is drawn", side.label, res.Range, res.Of))
			}
			if res.Missing != "" {
				notes = append(notes, fmt.Sprintf("%s: not drawn: %s", side.label, res.Missing))
			}
		}
		title := p.After
		if title == "" {
			title = p.Before
		} else if p.Before != "" && p.Before != p.After {
			title = p.After + " (was " + p.Before + ")"
		}
		w("### %d. %s\n\n", n, mdCode(title))
		var kinds []string
		for _, k := range []string{"value unexplained", "chart lost", "formula flagged", "input changed", "formula changed", "value changed", "value moved", "value unknown"} {
			if c := g.kinds[k]; c != nil {
				word, rest, _ := strings.Cut(k, " ")
				at := strings.Join(c.cells, ", ")
				if c.n > len(c.cells) {
					at += " and more"
				}
				kinds = append(kinds, countOf(c.n, word)+" "+rest+" ("+at+")")
			}
		}
		w("Here: %s.\n\n", strings.Join(kinds, ", "))
		w("| Before (%s) | After (%s) |\n|---|---|\n| %s | %s |\n\n", mdCode(r.Against), mdCode(r.File), cells[0], cells[1])
		if len(g.unknown) > 0 {
			notes = append(notes, fmt.Sprintf("values sheets cannot know (%s) are drawn as the file holds them, which sheets cannot vouch for", strings.Join(g.unknown, ", ")))
		}
		for _, x := range notes {
			w("- %s\n", mdText(x))
		}
		if len(notes) > 0 {
			w("\n")
		}
		info.Previews = append(info.Previews, p)
	}
	if left > 0 {
		info.Unshown = left
		w("Not pictured: %s outside these %s, at %s", countOf(left, "change"), countOf(len(regs), "range"), mdCode(strings.Join(where, ", ")))
		if left > len(where) {
			w(" and more")
		}
		w(".\n\n")
	}
	if len(unshown) > 0 {
		w("No picture can show these:\n\n")
		for _, u := range unshown[:min(len(unshown), receiptListed)] {
			w("- %s\n", mdText(u))
		}
		if len(unshown) > receiptListed {
			w("- … and %s more\n", grid.FormatNumber(float64(len(unshown)-receiptListed)))
		}
		w("\n")
	}
	return nil
}

// renderTo draws a range of a workbook to a PNG file.
func renderTo(file, out string, opts render.Options) (*render.Result, error) {
	tmp := out + ".tmp"
	fh, err := os.Create(tmp)
	if err != nil {
		return nil, err
	}
	res, err := render.PNG(file, fh, opts)
	if cerr := fh.Close(); err == nil {
		err = cerr
	}
	if err != nil {
		os.Remove(tmp)
		return nil, err
	}
	return res, os.Rename(tmp, out)
}

func writeFileAtomic(path string, data []byte) error {
	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, data, 0o644); err != nil {
		return err
	}
	return os.Rename(tmp, path)
}

// mdCode is a value as a code span in a Markdown table: as it is, pipes
// escaped; "empty" in italics for nothing.
func mdCode(s string) string {
	if s == "" {
		return "*empty*"
	}
	s = strings.NewReplacer("\r\n", " ", "\n", " ", "\r", " ").Replace(s)
	fence := "`"
	for strings.Contains(s, fence) {
		fence += "`"
	}
	pad := ""
	if strings.HasPrefix(s, "`") || strings.HasSuffix(s, "`") {
		pad = " "
	}
	return fence + pad + strings.ReplaceAll(s, "|", `\|`) + pad + fence
}

var mdEscaper = strings.NewReplacer(`\`, `\\`, "`", "\\`", "*", `\*`, "_", `\_`, "[", `\[`, "]", `\]`, "<", `\<`, ">", `\>`, "|", `\|`, "#", `\#`,
	"\r\n", " ", "\n", " ", "\r", " ")

// mdText is prose with Markdown's marks escaped.
func mdText(s string) string { return mdEscaper.Replace(s) }

// mdLink is a file name as a link target.
func mdLink(name string) string {
	if strings.ContainsAny(name, " ()<>") {
		return "<" + strings.NewReplacer("<", "%3C", ">", "%3E").Replace(name) + ">"
	}
	return name
}

func (a *app) printReceipt(info *receiptInfo) {
	n := 0
	for _, p := range info.Previews {
		if p.BeforeImage != "" {
			n++
		}
		if p.AfterImage != "" {
			n++
		}
	}
	a.printf("receipt: %s (%s of %s)\n", info.File, plural(n, "picture"), plural(len(info.Previews), "range"))
}
