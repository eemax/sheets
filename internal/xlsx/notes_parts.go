package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// excelize writes a note into xl/comments<N>.xml where N is the number in
// the name of the sheet's drawing for notes (xl/drawings/vmlDrawing<N>.vml),
// and names the parts of a sheet's first note after how many notes parts
// there are. Excel names its parts that way, mostly; openpyxl names them
// otherwise (xl/comments/comment1.xml, commentsDrawing1.vml) and points at
// them by absolute paths, which excelize turns into broken parts. So before
// excelize opens a file, its notes parts are renamed to what excelize
// expects: sheets with notes get comments<N>.xml and vmlDrawing<N>.vml for
// N = 1, 2, ..., and other drawings of the kind (form controls, header
// pictures) numbers from 1001, out of the way of new ones.

var (
	partNumber    = regexp.MustCompile(`^xl/(?:comments|drawings/vmlDrawing)(\d+)\.(?:xml|vml)$`)
	legacyDrawing = regexp.MustCompile(`<legacyDrawing\b[^>]*?\b(?:r:)?id="([^"]*)"`)
	relElem       = regexp.MustCompile(`<Relationship\b[^>]*/>|<Relationship\b[^>]*>\s*</Relationship>`)
	relAttr       = regexp.MustCompile(`\b(Id|Type|Target|TargetMode)="([^"]*)"`)
)

const farNumbers = 1001

// notesPlan is how a package's notes parts are renamed.
type notesPlan struct {
	rename  map[string]string            // part → new part
	targets map[string]map[string]string // sheet rels part → relationship id → new target
}

// normalizeNotes renames notes parts as excelize expects them (see above),
// returning data as it is when they are named so already.
func normalizeNotes(data []byte) ([]byte, error) {
	if !bytes.Contains(data, []byte("vmlDrawing")) && !bytes.Contains(data, []byte(".vml")) {
		return data, nil
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return data, nil // not ours to judge; the reader reports it
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	plan, ok := planNotes(p, zr)
	vml := map[string][]byte{} // drawings with other prefixes, as renamed
	for _, zf := range zr.File {
		if !strings.HasSuffix(strings.ToLower(zf.Name), ".vml") {
			continue
		}
		b, err := p.read(zf.Name)
		if err != nil {
			continue
		}
		if c := canonicalVML(b); !bytes.Equal(c, b) {
			vml[zf.Name] = c
		}
	}
	if ok && len(vml) == 0 {
		return data, nil
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zf := range zr.File {
		name := zf.Name
		if to, ok := plan.rename[name]; ok {
			name = to
		}
		_, retargeted := plan.targets[zf.Name]
		_, redrawn := vml[zf.Name]
		if !retargeted && !redrawn && (zf.Name != "[Content_Types].xml" || len(plan.rename) == 0) {
			// As it is, under its new name if it has one: the sheets are
			// not unpacked.
			fh := zf.FileHeader
			fh.Name = name
			w, err := zw.CreateRaw(&fh)
			if err != nil {
				return nil, err
			}
			r, err := zf.OpenRaw()
			if err != nil {
				return nil, err
			}
			if _, err := io.Copy(w, r); err != nil {
				return nil, err
			}
			continue
		}
		r, err := zf.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return nil, err
		}
		if t, ok := plan.targets[zf.Name]; ok {
			b = retarget(b, t)
		}
		if c, ok := vml[zf.Name]; ok {
			b = c
		}
		if zf.Name == "[Content_Types].xml" {
			for from, to := range plan.rename {
				b = bytes.ReplaceAll(b, []byte(`PartName="/`+from+`"`), []byte(`PartName="/`+to+`"`))
			}
		}
		w, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(b); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// planNotes works out the renames; ok is true when none is needed.
func planNotes(p *pkg, zr *zip.Reader) (notesPlan, bool) {
	plan := notesPlan{rename: map[string]string{}, targets: map[string]map[string]string{}}
	_, _, sheets, err := p.workbookParts()
	if err != nil {
		return plan, true
	}
	type sheetNotes struct {
		rels             string // the sheet's relationships part
		commentsID       string
		comments         string
		vmlID, vml       string
		otherVML         map[string]string // other drawings of the kind: id → part
		commentsRelative bool
		vmlRelative      bool
	}
	var order []string
	for _, part := range sheets {
		order = append(order, part)
	}
	sort.Slice(order, func(i, j int) bool { return natural(order[i], order[j]) })
	var withNotes []*sheetNotes
	var others []*sheetNotes
	taken := map[string]bool{} // parts accounted for
	for _, part := range order {
		dir := path.Dir(part)
		relsPart := path.Join(dir, "_rels", path.Base(part)+".rels")
		raw, err := p.read(relsPart)
		if err != nil {
			continue
		}
		sheetXML, err := p.read(part)
		if err != nil {
			continue
		}
		legacy := ""
		if m := legacyDrawing.FindSubmatch(sheetXML); m != nil {
			legacy = string(m[1])
		}
		sn := &sheetNotes{rels: relsPart, otherVML: map[string]string{}}
		for _, el := range relElem.FindAll(raw, -1) {
			attrs := map[string]string{}
			for _, a := range relAttr.FindAllSubmatch(el, -1) {
				attrs[string(a[1])] = string(a[2])
			}
			target := attrs["Target"]
			resolved := strings.TrimPrefix(target, "/")
			if !strings.HasPrefix(target, "/") {
				resolved = path.Clean(path.Join(dir, target))
			}
			switch {
			case strings.HasSuffix(attrs["Type"], "/comments"):
				sn.commentsID, sn.comments, sn.commentsRelative = attrs["Id"], resolved, !strings.HasPrefix(target, "/")
			case strings.HasSuffix(attrs["Type"], "/vmlDrawing"):
				if attrs["Id"] == legacy {
					sn.vmlID, sn.vml, sn.vmlRelative = attrs["Id"], resolved, !strings.HasPrefix(target, "/")
				} else {
					sn.otherVML[attrs["Id"]] = resolved
				}
			}
		}
		switch {
		case sn.comments != "" && sn.vml != "":
			withNotes = append(withNotes, sn)
		case sn.vml != "" || len(sn.otherVML) > 0:
			if sn.vml != "" {
				sn.otherVML[sn.vmlID] = sn.vml
				sn.vml = ""
			}
			others = append(others, sn)
		}
	}
	// Already as excelize expects them?
	ok := true
	for i, sn := range withNotes {
		n := strconv.Itoa(i + 1)
		ok = ok && sn.comments == "xl/comments"+n+".xml" && sn.vml == "xl/drawings/vmlDrawing"+n+".vml" && sn.commentsRelative && sn.vmlRelative
		taken[sn.comments], taken[sn.vml] = true, true
	}
	for _, sn := range others {
		for _, v := range sn.otherVML {
			taken[v] = true
		}
	}
	free := len(withNotes) + len(sheets)
	for _, zf := range zr.File {
		if m := partNumber.FindStringSubmatch(zf.Name); m != nil {
			n, _ := strconv.Atoi(m[1])
			if !taken[zf.Name] && n <= free || n <= free && n > len(withNotes) {
				ok = false // in the way of the parts of new notes
			}
		}
	}
	if ok {
		return plan, true
	}
	exists := map[string]bool{}
	for _, zf := range zr.File {
		exists[zf.Name] = true
	}
	move := func(from, to string) {
		if from == to {
			return
		}
		plan.rename[from] = to
		rels := path.Join(path.Dir(from), "_rels", path.Base(from)+".rels")
		if exists[rels] {
			plan.rename[rels] = path.Join(path.Dir(to), "_rels", path.Base(to)+".rels")
		}
	}
	retargetIn := func(sn *sheetNotes, id, target string) {
		if plan.targets[sn.rels] == nil {
			plan.targets[sn.rels] = map[string]string{}
		}
		plan.targets[sn.rels][id] = target
	}
	for i, sn := range withNotes {
		n := strconv.Itoa(i + 1)
		move(sn.comments, "xl/comments"+n+".xml")
		move(sn.vml, "xl/drawings/vmlDrawing"+n+".vml")
		retargetIn(sn, sn.commentsID, "../comments"+n+".xml")
		retargetIn(sn, sn.vmlID, "../drawings/vmlDrawing"+n+".vml")
	}
	next := max(farNumbers, free+1)
	for _, sn := range others {
		ids := make([]string, 0, len(sn.otherVML))
		for id := range sn.otherVML {
			ids = append(ids, id)
		}
		sort.Strings(ids)
		for _, id := range ids {
			n := strconv.Itoa(next)
			next++
			move(sn.otherVML[id], "xl/drawings/vmlDrawing"+n+".vml")
			retargetIn(sn, id, "../drawings/vmlDrawing"+n+".vml")
		}
	}
	// Notes parts that no sheet uses, out of the way too.
	for _, zf := range zr.File {
		if m := partNumber.FindStringSubmatch(zf.Name); m != nil && !taken[zf.Name] {
			if n, _ := strconv.Atoi(m[1]); n <= free {
				move(zf.Name, strings.Replace(zf.Name, m[1]+".", strconv.Itoa(next)+".", 1))
				next++
			}
		}
	}
	return plan, false
}

// retarget sets the targets of relationships by id.
func retarget(b []byte, targets map[string]string) []byte {
	return relElem.ReplaceAllFunc(b, func(el []byte) []byte {
		id := ""
		for _, a := range relAttr.FindAllSubmatch(el, -1) {
			if string(a[1]) == "Id" {
				id = string(a[2])
			}
		}
		t, ok := targets[id]
		if !ok {
			return el
		}
		return regexp.MustCompile(`\bTarget="[^"]*"`).ReplaceAll(el, []byte(fmt.Sprintf(`Target="%s"`, t)))
	})
}

// natural orders part names with their numbers as numbers: sheet2 before
// sheet10.
func natural(a, b string) bool {
	na, nb := trailingNumber(a), trailingNumber(b)
	if na >= 0 && nb >= 0 && strings.TrimRight(strings.TrimSuffix(a, path.Ext(a)), "0123456789") == strings.TrimRight(strings.TrimSuffix(b, path.Ext(b)), "0123456789") {
		return na < nb
	}
	return a < b
}

func trailingNumber(s string) int {
	s = strings.TrimSuffix(s, path.Ext(s))
	i := len(s)
	for i > 0 && s[i-1] >= '0' && s[i-1] <= '9' {
		i--
	}
	n, err := strconv.Atoi(s[i:])
	if err != nil {
		return -1
	}
	return n
}

// VML drawings name their namespaces v, o and x in Excel's files; openpyxl
// writes ns0, ns1 and ns2, which excelize drops from the root when it
// rewrites a drawing, leaving prefixes that point nowhere. They are renamed
// to Excel's before excelize opens the file.
var (
	vmlNamespaces = map[string]string{
		"urn:schemas-microsoft-com:vml":           "v",
		"urn:schemas-microsoft-com:office:office": "o",
		"urn:schemas-microsoft-com:office:excel":  "x",
	}
	nsDecl = regexp.MustCompile(`xmlns:(\w+)="([^"]*)"`)
)

// canonicalVML renames a drawing's namespace prefixes to v, o and x.
func canonicalVML(b []byte) []byte {
	rename := map[string]string{}
	used := map[string]string{} // prefix → namespace
	for _, m := range nsDecl.FindAllSubmatch(b, -1) {
		used[string(m[1])] = string(m[2])
	}
	for prefix, ns := range used {
		want, ok := vmlNamespaces[ns]
		if !ok || prefix == want {
			continue
		}
		if other, taken := used[want]; taken && other != ns {
			return b // a prefix we would need means something else here
		}
		rename[prefix] = want
	}
	if len(rename) == 0 {
		return b
	}
	for from, to := range rename {
		b = regexp.MustCompile(`(</?|\s)`+regexp.QuoteMeta(from)+`:`).ReplaceAll(b, []byte("${1}"+to+":"))
		b = bytes.ReplaceAll(b, []byte("xmlns:"+from+`="`), []byte("xmlns:"+to+`="`))
	}
	// One declaration per prefix.
	seen := map[string]bool{}
	return nsDecl.ReplaceAllFunc(b, func(m []byte) []byte {
		k := string(m)
		if seen[k] {
			return nil
		}
		seen[k] = true
		return m
	})
}

var (
	noteShape   = regexp.MustCompile(`(?s)<v:shape\b[^>]*>.*?</v:shape>`)
	shapeRow    = regexp.MustCompile(`<x:Row>\s*(\d+)\s*</x:Row>`)
	shapeColumn = regexp.MustCompile(`<x:Column>\s*(\d+)\s*</x:Column>`)
	noteType    = regexp.MustCompile(`ObjectType="Note"`)
	idmapData   = regexp.MustCompile(`<o:idmap\b[^>]*\bdata="(\d+)`)
	shapeID     = regexp.MustCompile(`\bid="_x0000_s(\d+)"`)
	commentRef  = regexp.MustCompile(`<comment\b[^>]*\bref="([A-Z]+[0-9]+)"`)
)

const noteShapeType = `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`

// fitNoteShapes makes a drawing hold one note shape per note of its sheet:
// the shapes of notes that moved or went are removed, and notes without a
// shape get one. excelize finds a note's shape by an anchor that files
// from other tools do not have, or have otherwise, so it leaves shapes
// behind when notes move.
func fitNoteShapes(vml, comments []byte) []byte {
	want := map[[2]int]bool{} // 0-based row and column
	var order [][2]int
	for _, m := range commentRef.FindAllSubmatch(comments, -1) {
		if r, c, ok := grid.ParseCell(string(m[1])); ok {
			k := [2]int{r - 1, c - 1}
			if !want[k] {
				want[k] = true
				order = append(order, k)
			}
		}
	}
	// The last shape for a cell wins: excelize appends the newest.
	last := map[[2]int]int{}
	shapes := noteShape.FindAllIndex(vml, -1)
	for i, loc := range shapes {
		s := vml[loc[0]:loc[1]]
		if !noteType.Match(s) {
			continue
		}
		r, c := shapeRow.FindSubmatch(s), shapeColumn.FindSubmatch(s)
		if r == nil || c == nil {
			continue
		}
		rr, _ := strconv.Atoi(string(r[1]))
		cc, _ := strconv.Atoi(string(c[1]))
		last[[2]int{rr, cc}] = i
	}
	var out bytes.Buffer
	prev := 0
	have := map[[2]int]bool{}
	maxID := 0
	for _, m := range shapeID.FindAllSubmatch(vml, -1) {
		n, _ := strconv.Atoi(string(m[1]))
		maxID = max(maxID, n)
	}
	for i, loc := range shapes {
		s := vml[loc[0]:loc[1]]
		keep := true
		if noteType.Match(s) {
			r, c := shapeRow.FindSubmatch(s), shapeColumn.FindSubmatch(s)
			keep = false
			if r != nil && c != nil {
				rr, _ := strconv.Atoi(string(r[1]))
				cc, _ := strconv.Atoi(string(c[1]))
				k := [2]int{rr, cc}
				keep = want[k] && last[k] == i
				have[k] = have[k] || keep
			}
		}
		out.Write(vml[prev:loc[0]])
		if keep {
			out.Write(s)
		}
		prev = loc[1]
	}
	out.Write(vml[prev:])
	b := out.Bytes()
	var add bytes.Buffer
	base := 1
	if m := idmapData.FindSubmatch(b); m != nil {
		base, _ = strconv.Atoi(string(m[1]))
	}
	next := max(maxID+1, base*1024+1)
	for _, k := range order {
		if have[k] {
			continue
		}
		fmt.Fprintf(&add, `<v:shape id="_x0000_s%d" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:120pt;height:60pt;z-index:%d;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`+
			`<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>`+
			`<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>`+
			`<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%d, 15, %d, 10, %d, 15, %d, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%d</x:Row><x:Column>%d</x:Column></x:ClientData></v:shape>`,
			next, next-base*1024, k[1]+1, max(k[0]-1, 0), k[1]+3, k[0]+3, k[0], k[1])
		next++
	}
	if add.Len() == 0 {
		return b
	}
	if !bytes.Contains(b, []byte(`id="_x0000_t202"`)) {
		add.Reset()
		return fitNoteShapes(insertShapeType(b), comments)
	}
	end := bytes.LastIndex(b, []byte("</xml>"))
	if end < 0 {
		return b
	}
	return append(append(append([]byte{}, b[:end]...), add.Bytes()...), b[end:]...)
}

// insertShapeType adds the shape type notes use to a drawing.
func insertShapeType(b []byte) []byte {
	at := bytes.Index(b, []byte("</o:shapelayout>"))
	if at >= 0 {
		at += len("</o:shapelayout>")
	} else if at = bytes.Index(b, []byte(">")); at >= 0 && bytes.HasPrefix(bytes.TrimSpace(b), []byte("<xml")) {
		at++
	} else {
		return append(b, noteShapeType...) // not a drawing we know; leave it be
	}
	return append(append(append([]byte{}, b[:at]...), noteShapeType...), b[at:]...)
}
