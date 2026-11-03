package cli

import (
	"errors"
	"fmt"
	"sort"
	"strings"
	"text/tabwriter"
	"time"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/lint"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/spf13/cobra"
)

// checkResult is what check reports.
type checkResult struct {
	File string `json:"file"`
	// Fixed: --fix saved a change (false on a dry run, or with nothing
	// to change).
	Fixed    bool           `json:"fixed"`
	Formulas *checkFormulas `json:"formulas,omitempty"`
	Compared *checkCompared `json:"compared,omitempty"`
	// Changes are what --fix changed (or, on a dry run, would change).
	Changes  []checkChange  `json:"changes,omitempty"`
	Errors   int            `json:"errors"`
	Warnings int            `json:"warnings"`
	Notes    int            `json:"notes"`
	Findings []checkFinding `json:"findings"`
}

// checkFormulas counts formulas by the state of their stored results.
type checkFormulas struct {
	Total      int `json:"total"`
	Vouched    int `json:"vouched"` // stored by Excel (or vouched for by a sheets stamp)
	Missing    int `json:"missing"`
	Stale      int `json:"stale"`
	Unprefixed int `json:"unprefixed"`
	// Of the missing and stale: computed by sheets (stored by --fix),
	// computed to what the file already holds, and unknown.
	Computable int `json:"computable"`
	Confirmed  int `json:"confirmed"`
	Unknown    int `json:"unknown"`
}

// checkCompared is the comparison of Excel's results with the sheets
// calculator's.
type checkCompared struct {
	Compared     int `json:"compared"`
	Agree        int `json:"agree"`
	Differ       int `json:"differ"`
	Uncomputable int `json:"uncomputable"`
	// ReadsUnstored: not compared, as they read a result Excel did not
	// store.
	ReadsUnstored int `json:"reads_unstored,omitempty"`
	OutOfTime     int `json:"out_of_time,omitempty"`
}

type checkChange struct {
	Change  string   `json:"change"`
	Sheet   string   `json:"sheet"`
	Cells   []string `json:"cells"`
	Count   int      `json:"count"`
	Message string   `json:"message"`
}

type checkFinding struct {
	Rule     string        `json:"rule"`
	Severity lint.Severity `json:"severity"`
	Sheet    string        `json:"sheet,omitempty"`
	Cells    []string      `json:"cells,omitempty"`
	Count    int           `json:"count,omitempty"`
	Message  string        `json:"message"`
}

// maxCellsJSON bounds the cells a JSON finding or change lists; count has
// them all.
const maxCellsJSON = 20

// checkTime bounds computing results, and comparing Excel's with the
// calculator's.
var checkTime = xlsx.CheckOptions{Budget: 2 * time.Minute, Compare: 10 * time.Second}

func (a *app) checkCmd() *cobra.Command {
	var (
		fix    bool
		failOn string
	)
	cmd := &cobra.Command{
		Use:   "check <file>",
		Short: "Check formula results and problems; --fix stores every result sheets can compute",
		Long: `Check a workbook, whatever wrote it, before handing it over: formulas with
no stored result (openpyxl stores none, so pandas and previews read blanks)
or a stale one, results sheets cannot know and why, formula errors,
functions stored without the _xlfn. prefix Excel needs (#NAME? in Excel),
results Excel stored that the sheets calculator does not agree with, and
everything lint finds. It writes nothing.

--fix stores every result sheets can compute and adds missing prefixes, in
one change you can undo. A result Excel stored is never replaced, and
nothing else changes. Running it again changes nothing.

Exits 1 on errors, as lint does (--fail-on warning|note|never). On a CSV
file it runs lint.`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			threshold := lint.Severity(99)
			if failOn != "never" {
				s, ok := lint.ParseSeverity(failOn)
				if !ok {
					return fmt.Errorf("--fail-on %q: use error, warning, note or never", failOn)
				}
				threshold = s
			}
			res, err := a.check(args[0], fix)
			if err != nil {
				return err
			}
			fail := false
			for _, f := range res.Findings {
				switch f.Severity {
				case lint.Error:
					res.Errors++
				case lint.Warning:
					res.Warnings++
				default:
					res.Notes++
				}
				fail = fail || f.Severity >= threshold
			}
			if a.json {
				if err := a.printJSON(res); err != nil {
					return err
				}
			} else {
				a.printCheck(res, fix)
			}
			if fail {
				return errSilent
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.BoolVar(&fix, "fix", false, "store every result sheets can compute and add missing prefixes")
	f.StringVar(&failOn, "fail-on", "error", "exit 1 on findings this severe: error, warning, note or never")
	return cmd
}

// check checks a file, fixing it with fix.
func (a *app) check(path string, fix bool) (*checkResult, error) {
	res := &checkResult{File: path}
	kind, err := book.FormatOf(path)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		// A CSV file holds values: there are no results to compute.
		wb, err := open(path)
		if err != nil {
			return nil, err
		}
		res.Findings = lintFindings(lint.Check(wb, lint.Options{}))
		return res, nil
	}
	data, err := xlsx.ReadFile(path)
	if err != nil {
		return nil, err
	}
	cr, err := xlsx.CheckResults(path, data, checkTime)
	if err != nil {
		return nil, err
	}
	if fix && cr.FixError != "" {
		return nil, errors.New(cr.FixError)
	}
	fc := &checkFormulas{Total: cr.Formulas, Vouched: cr.Vouched, Missing: cr.Missing, Stale: cr.Stale, Unprefixed: cr.Unprefixed,
		Confirmed: cr.Confirmed, Unknown: len(cr.Unknown)}
	for _, c := range cr.Changes {
		if c.Kind == "stored" || c.Kind == "replaced" {
			fc.Computable++
		}
	}
	res.Formulas = fc
	if cr.Compared+cr.Uncomputable+cr.NotCompared+cr.ReadsUnstored > 0 {
		res.Compared = &checkCompared{Compared: cr.Compared, Agree: cr.Agreed, Differ: len(cr.Differ), Uncomputable: cr.Uncomputable, OutOfTime: cr.NotCompared, ReadsUnstored: cr.ReadsUnstored}
	}
	res.Changes = changeGroups(cr.Changes)
	wb := cr.Before
	own := []checkFinding{}
	if fix {
		// What is left once fixed.
		if cr.Data != nil && !a.dryRun {
			if err := xlsx.SaveData(path, cr.Data); err != nil {
				return nil, err
			}
			res.Fixed = true
		}
		wb = cr.After
	} else {
		own = append(own, fixable(cr.Changes)...)
	}
	if cr.FixError != "" {
		own = append(own, checkFinding{Rule: "fix-failed", Severity: lint.Error,
			Message: "sheets cannot work out what --fix would store, so the results above are not checked: " + cr.FixError})
	}
	own = append(own, unknownFindings(cr.Unknown, !fix)...)
	own = append(own, skippedFindings(cr)...)
	own = append(own, differFindings(cr.Differ)...)
	res.Findings = append(own, lintFindings(lint.Check(wb, lint.Options{}))...)
	sort.SliceStable(res.Findings, func(i, j int) bool { return res.Findings[i].Severity > res.Findings[j].Severity })
	return res, nil
}

func lintFindings(fs []lint.Finding) []checkFinding {
	out := []checkFinding{}
	for _, f := range fs {
		out = append(out, checkFinding{Rule: f.Rule, Severity: f.Severity, Sheet: f.Sheet, Cells: capCells(f.Cells), Count: len(f.Cells), Message: f.Message})
	}
	return out
}

func capCells(cells []string) []string {
	if len(cells) > maxCellsJSON {
		return cells[:maxCellsJSON]
	}
	return cells
}

// cellGroup gathers cells under a key, in the order first seen.
type cellGroup struct {
	key, sheet string
	cells      []string
	first      xlsx.Change
	fns        []string
}

type groups struct {
	list []*cellGroup
	at   map[string]*cellGroup
}

func (g *groups) add(key, sheet string, row, col int) *cellGroup {
	k := sheet + "\x00" + key
	if g.at == nil {
		g.at = map[string]*cellGroup{}
	}
	x := g.at[k]
	if x == nil {
		x = &cellGroup{key: key, sheet: sheet}
		g.at[k] = x
		g.list = append(g.list, x)
	}
	x.cells = append(x.cells, grid.CellName(row, col))
	return x
}

// changeGroups groups what fixing changes by kind and sheet (and why, for
// results dropped).
func changeGroups(cs []xlsx.Change) []checkChange {
	var g groups
	for _, c := range cs {
		key := c.Kind
		if c.Kind == "dropped" {
			root, _ := rootReason(c.Why)
			key += "\x00" + root
		}
		x := g.add(key, c.Sheet, c.Row, c.Col)
		if len(x.cells) == 1 {
			x.first = c
		}
		for _, fn := range c.Functions {
			if !contains(x.fns, fn) {
				x.fns = append(x.fns, fn)
			}
		}
	}
	order := map[string]int{"prefixed": 0, "stored": 1, "replaced": 2, "filled": 3, "cleared": 4, "dropped": 5}
	sort.SliceStable(g.list, func(i, j int) bool { return order[g.list[i].first.Kind] < order[g.list[j].first.Kind] })
	var out []checkChange
	for _, x := range g.list {
		n, c := len(x.cells), x.first
		var msg string
		switch c.Kind {
		case "prefixed":
			msg = fmt.Sprintf("added the _xlfn. prefix Excel needs to %s: %s", countOf(n, "formula"), strings.Join(x.fns, ", "))
		case "stored":
			msg = fmt.Sprintf("stored %s where there %s none (%s = %s)", countOf(n, "computed result"), map[bool]string{true: "was", false: "were"}[n == 1], cellName(c), showCell(c.After))
		case "replaced":
			msg = fmt.Sprintf("replaced %s marked out of date with the computed %s (%s: %s → %s)", countOf(n, "result"), map[bool]string{true: "one", false: "ones"}[n == 1], cellName(c), showCell(c.Before), showCell(c.After))
		case "dropped":
			root, _ := rootReason(c.Why)
			msg = fmt.Sprintf("removed %s marked out of date that sheets cannot compute (%s)", countOf(n, "result"), root) + whenExcel(root, n)
		case "filled":
			msg = fmt.Sprintf("filled %s of an array formula's result", countOf(n, "cell"))
		case "cleared":
			msg = fmt.Sprintf("cleared %s of an array formula's old result", countOf(n, "cell"))
		}
		out = append(out, checkChange{Change: c.Kind, Sheet: x.sheet, Cells: capCells(x.cells), Count: n, Message: msg})
	}
	return out
}

func cellName(c xlsx.Change) string { return grid.CellName(c.Row, c.Col) }

// fixable are the findings a report makes of what --fix would change.
func fixable(cs []xlsx.Change) []checkFinding {
	var g groups
	for _, c := range cs {
		if c.Kind != "stored" && c.Kind != "replaced" {
			continue // prefixes: lint's; results dropped: unknown
		}
		x := g.add(c.Kind, c.Sheet, c.Row, c.Col)
		if len(x.cells) == 1 {
			x.first = c
		}
	}
	var out []checkFinding
	for _, x := range g.list {
		n, c := len(x.cells), x.first
		var f checkFinding
		if c.Kind == "stored" {
			f = checkFinding{Rule: "missing-result", Severity: lint.Error,
				Message: fmt.Sprintf("%s no stored result, so pandas, openpyxl and previews read %s; sheets computes %s (%s = %s): sheets check --fix stores %s",
					countOf(n, "formula")+map[bool]string{true: " has", false: " have"}[n == 1], map[bool]string{true: "a blank", false: "blanks"}[n == 1],
					map[bool]string{true: "it", false: "them"}[n == 1], cellName(c), showCell(c.After), map[bool]string{true: "it", false: "them"}[n == 1])}
		} else {
			f = checkFinding{Rule: "stale-result", Severity: lint.Error,
				Message: fmt.Sprintf("%s a result the file marks as out of date that differs from what sheets computes (%s holds %s, sheets gets %s): sheets check --fix replaces %s",
					countOf(n, "formula")+map[bool]string{true: " holds", false: " hold"}[n == 1], cellName(c), showCell(c.Before), showCell(c.After),
					map[bool]string{true: "it", false: "them"}[n == 1])}
		}
		f.Sheet, f.Cells, f.Count = x.sheet, capCells(x.cells), n
		out = append(out, f)
	}
	return out
}

// unknownFindings group the formulas whose results sheets cannot know by
// why; report marks stale results, which --fix removes.
func unknownFindings(us []xlsx.Unknown, report bool) []checkFinding {
	var g groups
	through := map[*cellGroup]int{}
	for _, u := range us {
		root, via := rootReason(u.Why)
		key := root
		if report && u.Stale {
			key = "stale\x00" + root
		}
		x := g.add(key, u.Sheet, u.Row, u.Col)
		if via {
			through[x]++
		}
	}
	var out []checkFinding
	for _, x := range g.list {
		n := len(x.cells)
		root := x.key
		stale := strings.HasPrefix(root, "stale\x00")
		root = strings.TrimPrefix(root, "stale\x00")
		msg := fmt.Sprintf("%s no result sheets can know: %s", countOf(n, "formula")+map[bool]string{true: " has", false: " have"}[n == 1], root)
		if k := through[x]; k > 0 && k < n {
			msg += fmt.Sprintf(" (for %d of them, in a formula they read)", k)
		} else if k > 0 {
			msg += " (in a formula " + map[bool]string{true: "it reads", false: "they read"}[n == 1] + ")"
		}
		if stale {
			msg += fmt.Sprintf("; %s a result marked as out of date, which sheets check --fix removes", map[bool]string{true: "it holds", false: "they hold"}[n == 1])
		}
		msg += whenExcel(root, n)
		out = append(out, checkFinding{Rule: "unknown-result", Severity: lint.Warning, Sheet: x.sheet, Cells: capCells(x.cells), Count: n, Message: msg})
	}
	return out
}

// skippedFindings say which prefixes --fix leaves out, and why.
func skippedFindings(cr *xlsx.CheckResult) []checkFinding {
	var out []checkFinding
	if cr.Withheld != "" {
		out = append(out, checkFinding{Rule: "prefix-not-added", Severity: lint.Warning,
			Message: "sheets check --fix adds no _xlfn. prefix to this workbook's formulas: " + cr.Withheld + "; enter them again with sheets set"})
	}
	var g groups
	for _, s := range cr.Skipped {
		x := g.add(s.Why, s.Sheet, s.Row, s.Col)
		_ = x
	}
	for _, x := range g.list {
		n := len(x.cells)
		out = append(out, checkFinding{Rule: "prefix-not-added", Severity: lint.Warning, Sheet: x.sheet, Cells: capCells(x.cells), Count: n,
			Message: fmt.Sprintf("sheets check --fix adds no _xlfn. prefix to %s: %s", countOf(n, "formula"), x.key)})
	}
	return out
}

// differFindings are results Excel stored that the calculator does not
// agree with: notes, as Excel's result is kept.
func differFindings(ds []xlsx.Difference) []checkFinding {
	bySheet := map[string][]xlsx.Difference{}
	var order []string
	for _, d := range ds {
		if bySheet[d.Sheet] == nil {
			order = append(order, d.Sheet)
		}
		bySheet[d.Sheet] = append(bySheet[d.Sheet], d)
	}
	var out []checkFinding
	for _, sheet := range order {
		list := bySheet[sheet]
		var cells, ex []string
		for i, d := range list {
			cells = append(cells, grid.CellName(d.Row, d.Col))
			if i < 2 {
				ex = append(ex, fmt.Sprintf("%s =%s: Excel %s, sheets %s", grid.CellName(d.Row, d.Col), d.Formula, showCell(d.Excel), showCell(d.Calc)))
			}
		}
		n := len(list)
		out = append(out, checkFinding{Rule: "calc-differs", Severity: lint.Note, Sheet: sheet, Cells: capCells(cells), Count: n,
			Message: fmt.Sprintf("the sheets calculator gets another result than Excel stored for %s (%s); Excel's %s kept", countOf(n, "formula"), strings.Join(ex, "; "),
				map[bool]string{true: "result is", false: "results are"}[n == 1])})
	}
	return out
}

// rootReason is why a result is unknown at its source: "it reads X,
// whose result is unknown here (why)" gives why, and via.
func rootReason(why string) (root string, via bool) {
	const mark = ", whose result is unknown here ("
	for strings.HasPrefix(why, "it reads ") {
		i := strings.Index(why, mark)
		if i < 0 {
			break
		}
		// The parenthesis that closes the one the mark opens.
		inner, depth, end := why[i+len(mark):], 1, -1
		for k, r := range inner {
			switch r {
			case '(':
				depth++
			case ')':
				depth--
			}
			if depth == 0 {
				end = k
				break
			}
		}
		if end < 0 {
			break
		}
		why, via = inner[:end], true
	}
	return why, via
}

// whenExcel says when Excel gives formulas unknown for a reason their
// results: on open, but not for #NAME?, a circular reference, or a link
// to another workbook.
func whenExcel(root string, n int) string {
	switch {
	case strings.Contains(root, "#NAME?") || strings.HasPrefix(root, "circular"):
		return ""
	case strings.Contains(root, "another workbook"):
		return "; Excel updates " + map[bool]string{true: "it", false: "them"}[n == 1] + " when that workbook is at hand"
	}
	return "; Excel computes " + map[bool]string{true: "it", false: "them"}[n == 1] + " on open"
}

func showCell(c grid.Cell) string {
	switch c.Kind {
	case grid.Text:
		s := c.Text
		if len([]rune(s)) > 40 {
			s = string([]rune(s)[:40]) + "…"
		}
		return fmt.Sprintf("%q", s)
	case grid.Error:
		return c.Text
	case grid.Bool:
		return strings.ToUpper(fmt.Sprint(c.Bool))
	case grid.Number:
		return grid.FormatNumber(c.Num)
	}
	return "nothing"
}

func countOf(n int, word string) string {
	if n == 1 {
		return "1 " + word
	}
	return grid.FormatNumber(float64(n)) + " " + word + "s"
}

func (a *app) printCheck(res *checkResult, fix bool) {
	if fix {
		a.printDryRun()
	}
	f := res.Formulas
	switch {
	case f == nil:
		a.printf("%s: a CSV file holds values, so only lint applies\n", res.File)
	case f.Total == 0:
		a.printf("%s: no formulas\n", res.File)
	default:
		line := fmt.Sprintf("%s: %s; results: %d stored and vouched for, %d missing, %d stale", res.File, countOf(f.Total, "formula"), f.Vouched, f.Missing, f.Stale)
		if f.Missing+f.Stale > 0 {
			line += fmt.Sprintf(" (sheets computes %d", f.Computable+f.Confirmed)
			if f.Unknown > 0 {
				line += fmt.Sprintf(", %d unknown", f.Unknown)
			}
			line += ")"
		}
		a.printf("%s\n", line)
	}
	if c := res.Compared; c != nil {
		line := fmt.Sprintf("stored results against the sheets calculator: %d agree, %d differ", c.Agree, c.Differ)
		if c.Uncomputable > 0 {
			line += fmt.Sprintf(", %d it cannot compute", c.Uncomputable)
		}
		if c.ReadsUnstored > 0 {
			line += fmt.Sprintf(", %d not compared as they read results Excel did not store", c.ReadsUnstored)
		}
		if c.OutOfTime > 0 {
			line += fmt.Sprintf(", %d left out of time", c.OutOfTime)
		}
		a.printf("%s\n", line)
	}
	if fix {
		switch {
		case len(res.Changes) == 0:
			a.printf("nothing to fix\n")
		case res.Fixed:
			a.printf("fixed (one change; sheets undo %s takes it back):\n", shellQuote(res.File))
		default:
			a.printf("--fix would change:\n")
		}
		tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
		for _, c := range res.Changes {
			fmt.Fprintf(tw, "%s\t%s\t%s\n", c.Change, changeCells(c), c.Message)
		}
		tw.Flush()
	}
	if len(res.Findings) == 0 {
		a.printf("no problems found\n")
	} else {
		a.printf("%s, %s, %s\n", plural(res.Errors, "error"), plural(res.Warnings, "warning"), plural(res.Notes, "note"))
		tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
		for _, f := range res.Findings {
			fmt.Fprintf(tw, "%s\t%s\t%s\n", f.Severity, findingCells(f), f.Message)
		}
		tw.Flush()
	}
	switch {
	case fix && res.Fixed:
		a.hint("sheets diff %s", shellQuote(res.File))
	case !fix && f != nil && (len(res.Changes) > 0):
		a.hint("sheets check --fix %s", shellQuote(res.File))
	}
}

func changeCells(c checkChange) string {
	return cellCount(c.Sheet, c.Cells, c.Count)
}

func findingCells(f checkFinding) string {
	return cellCount(f.Sheet, f.Cells, f.Count)
}

// cellCount renders Sheet!A1, B2 (+5 more), from the cells listed and
// their count.
func cellCount(sheet string, cells []string, count int) string {
	if len(cells) == 0 {
		return "(workbook)"
	}
	shown := cells[:min(len(cells), maxCellsListed)]
	s := grid.QuoteSheet(sheet) + "!" + strings.Join(shown, ", ")
	if count > len(shown) {
		s += fmt.Sprintf(" (+%s more)", grid.FormatNumber(float64(count-len(shown))))
	}
	return s
}
