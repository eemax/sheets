package xlsx

import (
	"encoding/xml"
	"path"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Notes and hyperlinks. A sheet's notes are in a comments part (Excel's
// notes, called comments before Excel 365); Excel 365's comment threads
// are in a threadedComments part, with a note standing in for each thread
// in the comments part for older Excel. Hyperlinks are in the sheet, their
// web addresses in its relationships.

// rawLink is a hyperlink as the sheet holds it.
type rawLink struct {
	ref, relID, location, tooltip string
}

// sheetParts are the parts attached to a worksheet that the reader uses.
type sheetParts struct {
	tables   []tablePart
	drawing  string
	comments string            // the notes part
	threads  string            // the threadedComments part
	links    map[string]string // hyperlink relationship id → address
}

// notesOf reads a sheet's notes: threads, and notes on cells without one.
func (p *pkg) notesOf(parts sheetParts, people map[string]string) []grid.Note {
	var notes []grid.Note
	threaded := map[[2]int]bool{}
	if parts.threads != "" {
		for _, n := range p.threads(parts.threads, people) {
			threaded[[2]int{n.Row, n.Col}] = true
			notes = append(notes, n)
		}
	}
	if parts.comments != "" {
		for _, n := range p.comments(parts.comments) {
			if !threaded[[2]int{n.Row, n.Col}] {
				notes = append(notes, n)
			}
		}
	}
	sort.Slice(notes, func(i, j int) bool {
		if notes[i].Row != notes[j].Row {
			return notes[i].Row < notes[j].Row
		}
		return notes[i].Col < notes[j].Col
	})
	return notes
}

// xRichText is the text of a note: plain, or in runs.
type xRichText struct {
	T    *string `xml:"t"`
	Runs []struct {
		T string `xml:"t"`
	} `xml:"r"`
}

func (x xRichText) String() string {
	var b strings.Builder
	if x.T != nil {
		b.WriteString(*x.T)
	}
	for _, r := range x.Runs {
		b.WriteString(r.T)
	}
	return unescapeOOXML(b.String())
}

// comments reads a notes part. Excel starts a note's text with its
// author's name ("Ann:" in bold); that line is left out of the text.
func (p *pkg) comments(name string) []grid.Note {
	var x struct {
		Authors  []string `xml:"authors>author"`
		Comments []struct {
			Ref      string    `xml:"ref,attr"`
			AuthorID int       `xml:"authorId,attr"`
			Text     xRichText `xml:"text"`
		} `xml:"commentList>comment"`
	}
	if p.decode(name, &x) != nil {
		return nil
	}
	var out []grid.Note
	for _, c := range x.Comments {
		r, col, ok := grid.ParseCell(c.Ref)
		if !ok {
			continue
		}
		n := grid.Note{Row: r, Col: col, Text: c.Text.String()}
		if c.AuthorID >= 0 && c.AuthorID < len(x.Authors) {
			n.Author = x.Authors[c.AuthorID]
		}
		if strings.HasPrefix(n.Author, "tc=") {
			n.Author = "" // the stand-in for a thread, named by its id
		}
		if n.Author != "" {
			if rest, ok := strings.CutPrefix(n.Text, n.Author+":"); ok {
				n.Text = strings.TrimLeft(rest, "\r\n")
			}
		}
		out = append(out, n)
	}
	return out
}

// threads reads a threadedComments part: each thread becomes one note, its
// replies on lines of their own after the first comment.
func (p *pkg) threads(name string, people map[string]string) []grid.Note {
	var x struct {
		Comments []struct {
			Ref      string `xml:"ref,attr"`
			ID       string `xml:"id,attr"`
			ParentID string `xml:"parentId,attr"`
			PersonID string `xml:"personId,attr"`
			Text     string `xml:"text"`
		} `xml:"threadedComment"`
	}
	if p.decode(name, &x) != nil {
		return nil
	}
	byID := map[string]int{}
	var out []grid.Note
	for _, c := range x.Comments {
		text := unescapeOOXML(c.Text)
		if c.ParentID != "" {
			if i, ok := byID[c.ParentID]; ok {
				who := people[c.PersonID]
				if who == "" {
					who = "reply"
				}
				out[i].Text += "\n" + who + ": " + text
				byID[c.ID] = i
			}
			continue
		}
		r, col, ok := grid.ParseCell(c.Ref)
		if !ok {
			continue
		}
		byID[c.ID] = len(out)
		out = append(out, grid.Note{Row: r, Col: col, Author: people[c.PersonID], Text: text, Thread: true})
	}
	return out
}

// people reads the workbook's persons part: who wrote comment threads.
func (p *pkg) people(wbRels []xRel) map[string]string {
	out := map[string]string{}
	for _, r := range wbRels {
		if !strings.HasSuffix(r.Type, "/person") {
			continue
		}
		var x struct {
			People []struct {
				ID   string `xml:"id,attr"`
				Name string `xml:"displayName,attr"`
			} `xml:"person"`
		}
		if p.decode(r.Target, &x) == nil {
			for _, person := range x.People {
				out[person.ID] = person.Name
			}
		}
	}
	return out
}

// linksOf resolves a sheet's hyperlinks. A link over a range is reported
// on its first cell.
func linksOf(raw []rawLink, addresses map[string]string) []grid.Link {
	var out []grid.Link
	for _, l := range raw {
		ref, _, _ := strings.Cut(l.ref, ":")
		r, c, ok := grid.ParseCell(ref)
		if !ok {
			continue
		}
		link := grid.Link{Row: r, Col: c, Location: l.location, Tooltip: l.tooltip}
		if l.relID != "" {
			link.URL = addresses[l.relID]
		}
		if link.URL == "" && link.Location == "" {
			continue
		}
		out = append(out, link)
	}
	sort.Slice(out, func(i, j int) bool {
		if out[i].Row != out[j].Row {
			return out[i].Row < out[j].Row
		}
		return out[i].Col < out[j].Col
	})
	return out
}

// hyperlink reads a <hyperlink> element.
func hyperlink(se xml.StartElement) rawLink {
	var l rawLink
	for _, a := range se.Attr {
		switch {
		case a.Name.Local == "ref":
			l.ref = a.Value
		case a.Name.Local == "id" && a.Name.Space != "":
			l.relID = a.Value
		case a.Name.Local == "location":
			l.location = a.Value
		case a.Name.Local == "tooltip":
			l.tooltip = a.Value
		}
	}
	return l
}

// sheetParts lists what is attached to a worksheet: its Excel tables,
// drawing, notes, comment threads and hyperlink addresses. Parts that
// cannot be read are skipped.
func (p *pkg) sheetParts(sheetPart string) sheetParts {
	var out sheetParts
	dir := path.Dir(sheetPart)
	rels, err := p.rels(path.Join(dir, "_rels", path.Base(sheetPart)+".rels"), dir)
	if err != nil {
		return out
	}
	out.links = map[string]string{}
	for _, r := range rels {
		switch {
		case strings.HasSuffix(r.Type, "/drawing"):
			out.drawing = r.Target
		case strings.HasSuffix(r.Type, "/comments"):
			out.comments = r.Target
		case strings.HasSuffix(r.Type, "/threadedComment"):
			out.threads = r.Target
		case strings.HasSuffix(r.Type, "/hyperlink"):
			out.links[r.ID] = r.Target
		case strings.HasSuffix(r.Type, "/table"):
			if t, ok := p.table(r.Target); ok {
				out.tables = append(out.tables, t)
			}
		}
	}
	return out
}
