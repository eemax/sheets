package render

import (
	"image"
	"image/color"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
	"golang.org/x/image/font"
	"golang.org/x/image/vector"
)

var (
	inkPrimary   = color.RGBA{0x0B, 0x0B, 0x0B, 255}
	inkSecondary = color.RGBA{0x52, 0x51, 0x4E, 255}
	inkMuted     = color.RGBA{0x89, 0x87, 0x81, 255}
	hairline     = color.RGBA{0xE1, 0xE0, 0xD9, 255}
	axisLine     = color.RGBA{0xC3, 0xC2, 0xB7, 255}
)

type series struct {
	name   string
	kind   string
	values []float64 // NaN where a cell holds no number
	xs     []float64 // scatter x values
	color  color.RGBA
	points []color.RGBA // pie slice colours
}

// resolve reads a range reference (Sheet!$B$2:$B$9) from the workbook.
func (sh *sheet) resolve(ref string) (*grid.Sheet, grid.Range, bool) {
	r, err := grid.ParseRef(ref, nil)
	if err != nil || r.Range == nil {
		return nil, grid.Range{}, false
	}
	s := sh.wb.Sheet(r.Sheet)
	if r.Sheet == "" {
		s = sh.wb.Sheet(sh.name)
	}
	if s == nil {
		return nil, grid.Range{}, false
	}
	return s, *r.Range, true
}

func (sh *sheet) numbers(ref string) []float64 {
	s, rg, ok := sh.resolve(ref)
	if !ok {
		return nil
	}
	var out []float64
	s2 := s
	for r := rg.R1; r <= rg.R2; r++ {
		for c := rg.C1; c <= rg.C2; c++ {
			v := sh.wb.Value(s2, r, c)
			if v != nil && v.Kind == grid.Number && !v.Pending {
				out = append(out, v.Num)
			} else {
				out = append(out, math.NaN())
			}
		}
	}
	return out
}

func (sh *sheet) labels(ref string) []string {
	s, rg, ok := sh.resolve(ref)
	if !ok {
		return nil
	}
	var out []string
	for r := rg.R1; r <= rg.R2; r++ {
		for c := rg.C1; c <= rg.C2; c++ {
			v := sh.wb.Value(s, r, c)
			if v == nil {
				out = append(out, "")
				continue
			}
			out = append(out, grid.Display(v, sh.wb.FormatOf(v), sh.wb.Date1904))
		}
	}
	return out
}

func palette(i int) color.RGBA {
	c, _ := hexColor(xlsx.Palette[i%len(xlsx.Palette)])
	return c
}

// chart draws a chart in its box.
func (p *painter) chart(sh *sheet, pc placedChart) {
	c := pc.chart
	x0, y0, x1, y1 := pc.x, pc.y, pc.x+pc.w, pc.y+pc.h
	p.fill(x0, y0, x1, y1, white)
	p.fill(x0, y0, x1, y0+1, hairline)
	p.fill(x0, y1-1, x1, y1, hairline)
	p.fill(x0, y0, x0+1, y1, hairline)
	p.fill(x1-1, y0, x1, y1, hairline)

	kind := c.Type
	if i := strings.IndexByte(kind, '+'); i >= 0 {
		kind = kind[:i]
	}
	stacked := strings.Contains(kind, "stacked")
	percent := strings.HasPrefix(kind, "100%")
	base := strings.TrimSpace(strings.TrimPrefix(strings.TrimPrefix(kind, "100% "), "stacked "))

	var ss []series
	var cats []string
	for i, s := range c.Series {
		sr := series{name: "Series " + itoa(i+1), kind: s.Kind, values: sh.numbers(s.Values), color: palette(i)}
		if s.Name != "" {
			if ls := sh.labels(s.Name); len(ls) > 0 && ls[0] != "" {
				sr.name = ls[0]
			}
		}
		if col, ok := hexColor(s.Color); ok {
			sr.color = col
		}
		for j, pcol := range s.PointColors {
			col, ok := hexColor(pcol)
			if !ok {
				col = palette(j)
			}
			sr.points = append(sr.points, col)
		}
		if sr.kind == "" {
			sr.kind = base
		}
		if s.Categories != "" {
			if base == "scatter" {
				sr.xs = sh.numbers(s.Categories)
			} else if cats == nil {
				cats = sh.labels(s.Categories)
			}
		}
		ss = append(ss, sr)
	}
	n := 0
	for _, s := range ss {
		n = max(n, len(s.values))
	}
	for len(cats) < n {
		cats = append(cats, itoa(len(cats)+1))
	}

	top := y0 + 10
	if c.Title != "" {
		lk := &look{size: 13, ink: inkPrimary}
		f := lookFace(lk, p.scale)
		t := truncateTo(c.Title, face(false, false, 13, 1), x1-x0-20)
		p.label(t, lk, f, (x0+x1)/2-textWidth(t, lk)/2, top+14)
		top += 26
	}
	// Legend entries: combined charts list each name once.
	var entries []series
	seen := map[string]bool{}
	for _, s := range ss {
		if !seen[s.name] {
			seen[s.name] = true
			entries = append(entries, s)
		}
	}
	right, bottom := x1-12, y1-10
	if base == "pie" {
		p.pie(ss, cats, x0+10, top, x1-10, bottom, c.Legend)
		return
	}
	if c.Legend != "" && len(entries) > 0 {
		switch c.Legend {
		case "r", "tr":
			right = p.legendColumn(entries, right, top) - 10
		case "b":
			bottom = p.legendRow(entries, x0, x1, bottom-14) - 6
		default:
			top = p.legendRow(entries, x0, x1, top+4) + 8
		}
	}

	// Value range.
	lo, hi := 0.0, 0.0
	first := true
	see := func(v float64) {
		if math.IsNaN(v) {
			return
		}
		if first {
			lo, hi, first = math.Min(0, v), math.Max(0, v), false
		}
		lo, hi = math.Min(lo, v), math.Max(hi, v)
	}
	if stacked {
		for i := 0; i < n; i++ {
			pos, neg := 0.0, 0.0
			for _, s := range ss {
				if i < len(s.values) && !math.IsNaN(s.values[i]) {
					if s.values[i] >= 0 {
						pos += s.values[i]
					} else {
						neg += s.values[i]
					}
				}
			}
			if percent {
				pos, neg = 1, 0
			}
			see(pos)
			see(neg)
		}
	} else {
		for _, s := range ss {
			for _, v := range s.values {
				see(v)
			}
		}
	}
	if hi == lo {
		hi = lo + 1
	}
	ticks := niceTicks(lo, hi, 5)
	lo, hi = ticks[0], ticks[len(ticks)-1]
	format := c.ValueFormat
	if percent {
		format = "0%"
	}
	lk := &look{size: 9, ink: inkMuted}
	f := lookFace(lk, p.scale)
	horizontal := base == "bar"

	// Plot box, leaving room for axis labels.
	labelW := 0
	for _, t := range ticks {
		labelW = max(labelW, textWidth(axisLabel(t, format), lk))
	}
	px0, px1, py0, py1 := x0+labelW+16, right, top, bottom-20
	// Bars: the value axis sits at the first category's end unless it
	// crosses at the last; categories run bottom-up unless reversed.
	labelsTop := horizontal && c.Reversed != c.CrossesMax
	if horizontal {
		catW := 0
		for _, cat := range cats {
			catW = max(catW, textWidth(cat, lk))
		}
		px0 = x0 + min(catW, (x1-x0)/3) + 16
		py1 = bottom - 18
		if labelsTop {
			py0, py1 = top+18, bottom
		}
	}
	if px1-px0 < 40 || py1-py0 < 40 {
		return // too small to draw
	}

	// Gridlines and value labels.
	for _, t := range ticks {
		label := axisLabel(t, format)
		if horizontal {
			x := px0 + int(float64(px1-px0)*(t-lo)/(hi-lo))
			p.fill(x, py0, x+1, py1, hairline)
			ly := py1 + 14
			if labelsTop {
				ly = py0 - 6
			}
			p.label(label, lk, f, x-textWidth(label, lk)/2, ly)
		} else {
			y := py1 - int(float64(py1-py0)*(t-lo)/(hi-lo))
			p.fill(px0, y, px1, y+1, hairline)
			p.label(label, lk, f, px0-6-textWidth(label, lk), y+4)
		}
	}
	if horizontal {
		p.fill(px0, py0, px0+1, py1, axisLine)
	} else {
		zero := py1 - int(float64(py1-py0)*(0-lo)/(hi-lo))
		p.fill(px0, zero, px1, zero+1, axisLine)
	}

	if base == "scatter" {
		p.scatter(ss, lk, f, px0, py0, px1, py1, lo, hi)
		return
	}

	// Categories.
	band := float64(px1-px0) / float64(max(n, 1))
	if horizontal {
		band = float64(py1-py0) / float64(max(n, 1))
	}
	catLk := &look{size: 9, ink: inkSecondary}
	catF := lookFace(catLk, p.scale)
	every := 1
	if !horizontal {
		widest := 0
		for _, cat := range cats {
			widest = max(widest, textWidth(cat, catLk)+6)
		}
		every = int(math.Ceil(float64(widest) / math.Max(band, 1)))
	}
	pos := func(i int) float64 { // centre of category i along its axis
		if horizontal && !c.Reversed {
			i = n - 1 - i
		}
		if !horizontal && c.Reversed {
			i = n - 1 - i
		}
		if horizontal {
			return float64(py0) + band*(float64(i)+0.5)
		}
		return float64(px0) + band*(float64(i)+0.5)
	}
	for i := 0; i < n; i++ {
		if i%every != 0 {
			continue
		}
		cat := cats[i]
		if horizontal {
			cat = truncateTo(cat, face(false, false, 9, 1), px0-x0-14)
			p.label(cat, catLk, catF, px0-8-textWidth(cat, catLk), int(pos(i))+4)
		} else {
			cat = truncateTo(cat, face(false, false, 9, 1), int(band*float64(every))-4)
			p.label(cat, catLk, catF, int(pos(i))-textWidth(cat, catLk)/2, py1+14)
		}
	}

	scaleV := func(v float64) float64 { // value to pixel along the value axis
		if horizontal {
			return float64(px0) + float64(px1-px0)*(v-lo)/(hi-lo)
		}
		return float64(py1) - float64(py1-py0)*(v-lo)/(hi-lo)
	}

	// Bars and columns.
	var bars []series
	for _, s := range ss {
		if s.kind == "column" || s.kind == "bar" {
			bars = append(bars, s)
		}
	}
	if len(bars) > 0 {
		groupW := band * 0.62
		each := groupW / float64(len(bars))
		if stacked {
			each = groupW
		}
		each = math.Min(each, 24*1.5)
		for i := 0; i < n; i++ {
			posSum, negSum := 0.0, 0.0
			total := 0.0
			for _, s := range bars {
				if i < len(s.values) && !math.IsNaN(s.values[i]) {
					total += math.Abs(s.values[i])
				}
			}
			for k, s := range bars {
				if i >= len(s.values) || math.IsNaN(s.values[i]) {
					continue
				}
				v := s.values[i]
				if percent && total > 0 {
					v /= total
				}
				from, to := 0.0, v
				var off float64
				if stacked {
					if v >= 0 {
						from, to = posSum, posSum+v
						posSum += v
					} else {
						from, to = negSum, negSum+v
						negSum += v
					}
					off = -each / 2
				} else {
					off = -each*float64(len(bars))/2 + each*float64(k)
				}
				a, b := scaleV(from), scaleV(to)
				cpos := pos(i)
				var r [4]float64
				if horizontal {
					r = [4]float64{math.Min(a, b), cpos + off, math.Max(a, b), cpos + off + each}
				} else {
					r = [4]float64{cpos + off, math.Min(a, b), cpos + off + each, math.Max(a, b)}
				}
				p.rect(r, s.color)
				// The surface gap between touching marks.
				p.strokeRect(r, white, 1)
			}
		}
	}

	// Areas, then lines on top.
	var prev []float64
	for _, s := range ss {
		if !strings.Contains(s.kind, "area") {
			continue
		}
		var pts, base [][2]float64
		for i := 0; i < n && i < len(s.values); i++ {
			v := s.values[i]
			if math.IsNaN(v) {
				v = 0
			}
			b := 0.0
			if stacked && prev != nil {
				b = prev[i]
			}
			pts = append(pts, [2]float64{pos(i), scaleV(b + v)})
			base = append([][2]float64{{pos(i), scaleV(b)}}, base...)
		}
		alpha := uint8(64)
		if stacked {
			alpha = 180
		}
		col := color.NRGBA{s.color.R, s.color.G, s.color.B, alpha}
		p.polygon(append(pts, base...), col)
		if stacked {
			next := make([]float64, n)
			for i := range next {
				if prev != nil {
					next[i] = prev[i]
				}
				if i < len(s.values) && !math.IsNaN(s.values[i]) {
					next[i] += s.values[i]
				}
			}
			prev = next
		}
	}
	for _, s := range ss {
		if s.kind != "line" {
			continue
		}
		var run [][2]float64
		flush := func() {
			if len(run) > 1 {
				p.polyline(run, 1.5, s.color)
			}
			run = nil
		}
		for i := 0; i < n && i < len(s.values); i++ {
			if math.IsNaN(s.values[i]) {
				flush()
				continue
			}
			run = append(run, [2]float64{pos(i), scaleV(s.values[i])})
		}
		flush()
		if n <= 12 && !strings.Contains(c.Type, "area") {
			for i := 0; i < n && i < len(s.values); i++ {
				if !math.IsNaN(s.values[i]) {
					p.circle(pos(i), scaleV(s.values[i]), 4.5, white)
					p.circle(pos(i), scaleV(s.values[i]), 3, s.color)
				}
			}
		}
	}
}

func (p *painter) label(s string, lk *look, f font.Face, x, y int) {
	d := &font.Drawer{Dst: p.img, Src: image.NewUniform(lk.ink), Face: f}
	drawString(d, s, float64(x+p.ox)*p.scale, float64(y+p.oy)*p.scale)
}

// legendRow draws legend entries in a centred row; it returns the row's
// bottom edge.
func (p *painter) legendRow(entries []series, x0, x1, y int) int {
	lk := &look{size: 9, ink: inkSecondary}
	f := lookFace(lk, p.scale)
	total := 0
	for _, e := range entries {
		total += 14 + textWidth(e.name, lk) + 14
	}
	x := (x0+x1)/2 - total/2
	for _, e := range entries {
		p.rect([4]float64{float64(x), float64(y + 2), float64(x + 10), float64(y + 12)}, e.color)
		p.label(e.name, lk, f, x+14, y+11)
		x += 14 + textWidth(e.name, lk) + 14
	}
	return y + 14
}

// legendColumn draws legend entries stacked at the right; it returns the
// column's left edge.
func (p *painter) legendColumn(entries []series, right, top int) int {
	lk := &look{size: 9, ink: inkSecondary}
	f := lookFace(lk, p.scale)
	w := 0
	for _, e := range entries {
		w = max(w, textWidth(e.name, lk))
	}
	x := right - w - 14
	for i, e := range entries {
		y := top + 10 + i*16
		p.rect([4]float64{float64(x), float64(y), float64(x + 10), float64(y + 10)}, e.color)
		p.label(e.name, lk, f, x+14, y+9)
	}
	return x
}

func (p *painter) pie(ss []series, cats []string, x0, y0, x1, y1 int, legend string) {
	if len(ss) == 0 {
		return
	}
	s := ss[0]
	total := 0.0
	for _, v := range s.values {
		if !math.IsNaN(v) && v > 0 {
			total += v
		}
	}
	if total == 0 {
		return
	}
	var entries []series
	for i, cat := range cats {
		col := palette(i)
		if i < len(s.points) {
			col = s.points[i]
		}
		entries = append(entries, series{name: cat, color: col})
	}
	right := x1
	if legend != "" {
		right = p.legendColumn(entries, x1, y0) - 10
	}
	cx, cy := float64(x0+right)/2, float64(y0+y1)/2
	r := math.Min(float64(right-x0), float64(y1-y0)) / 2 * 0.9
	angle := -math.Pi / 2 // twelve o'clock, clockwise
	lk := &look{size: 9, ink: white}
	f := lookFace(lk, p.scale)
	for i, v := range s.values {
		if math.IsNaN(v) || v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		pts := [][2]float64{{cx, cy}}
		steps := max(2, int(sweep*40))
		for k := 0; k <= steps; k++ {
			a := angle + sweep*float64(k)/float64(steps)
			pts = append(pts, [2]float64{cx + r*math.Cos(a), cy + r*math.Sin(a)})
		}
		p.polygon(pts, color.NRGBA{entries[i].color.R, entries[i].color.G, entries[i].color.B, 255})
		p.polyline(append(pts, [2]float64{cx, cy}), 1, white)
		if v/total >= 0.05 {
			mid := angle + sweep/2
			label := strconv.Itoa(int(math.Round(100*v/total))) + "%"
			lx := cx + 0.62*r*math.Cos(mid) - float64(textWidth(label, lk))/2
			ly := cy + 0.62*r*math.Sin(mid) + 4
			p.label(label, lk, f, int(lx), int(ly))
		}
		angle += sweep
	}
}

func (p *painter) scatter(ss []series, lk *look, f font.Face, px0, py0, px1, py1 int, lo, hi float64) {
	xlo, xhi, first := 0.0, 0.0, true
	for _, s := range ss {
		for _, x := range s.xs {
			if math.IsNaN(x) {
				continue
			}
			if first {
				xlo, xhi, first = x, x, false
			}
			xlo, xhi = math.Min(xlo, x), math.Max(xhi, x)
		}
	}
	if xhi == xlo {
		xhi = xlo + 1
	}
	ticks := niceTicks(xlo, xhi, 5)
	xlo, xhi = ticks[0], ticks[len(ticks)-1]
	for _, t := range ticks {
		x := px0 + int(float64(px1-px0)*(t-xlo)/(xhi-xlo))
		label := axisLabel(t, "")
		p.label(label, lk, f, x-textWidth(label, lk)/2, py1+14)
	}
	for _, s := range ss {
		for i := 0; i < len(s.values) && i < len(s.xs); i++ {
			if math.IsNaN(s.values[i]) || math.IsNaN(s.xs[i]) {
				continue
			}
			x := float64(px0) + float64(px1-px0)*(s.xs[i]-xlo)/(xhi-xlo)
			y := float64(py1) - float64(py1-py0)*(s.values[i]-lo)/(hi-lo)
			p.circle(x, y, 5, white)
			p.circle(x, y, 3.5, s.color)
		}
	}
}

// niceTicks returns about n round tick values covering lo..hi.
func niceTicks(lo, hi float64, n int) []float64 {
	span := hi - lo
	step := math.Pow(10, math.Floor(math.Log10(span/float64(n))))
	for _, m := range []float64{1, 2, 2.5, 5, 10} {
		if span/(step*m) <= float64(n) {
			step *= m
			break
		}
	}
	start := math.Floor(lo/step) * step
	end := math.Ceil(hi/step) * step
	var out []float64
	for v := start; v <= end+step/2; v += step {
		out = append(out, math.Round(v/step)*step)
	}
	return out
}

var quoted = regexp.MustCompile(`"([^"]*)"`)

// axisLabel formats a tick value roughly as the axis number format would:
// thousands (,"K") and millions (,,"M"), percentages, a currency prefix.
func axisLabel(v float64, code string) string {
	section := code
	if i := strings.IndexByte(code, ';'); i >= 0 {
		section = code[:i]
	}
	prefix, suffix := "", ""
	if m := quoted.FindAllStringSubmatch(section, -1); m != nil {
		for _, q := range m {
			switch {
			case q[1] == "K" || q[1] == "M":
				suffix = q[1]
			case strings.HasPrefix(section, `"`+q[1]+`"`):
				prefix = q[1]
			}
		}
	}
	switch {
	case strings.Contains(section, "%"):
		return strconv.FormatFloat(v*100, 'f', -1, 64) + "%"
	case strings.Contains(section, `,,"M"`):
		v /= 1e6
	case strings.Contains(section, `,"K"`):
		v /= 1e3
	}
	s := strconv.FormatFloat(math.Abs(v), 'f', -1, 64)
	if strings.Contains(section, "#,##") {
		s = group(s)
	}
	if v < 0 {
		s = "-" + s
	}
	return prefix + s + suffix
}

func group(s string) string {
	whole, frac, has := strings.Cut(s, ".")
	var b strings.Builder
	for i, r := range whole {
		if i > 0 && (len(whole)-i)%3 == 0 {
			b.WriteByte(',')
		}
		b.WriteRune(r)
	}
	if has {
		b.WriteString("." + frac)
	}
	return b.String()
}

// Shapes, anti-aliased, in pixels at scale 1.

func (p *painter) pt(x, y float64) (float32, float32) {
	return float32((x + float64(p.ox)) * p.scale), float32((y + float64(p.oy)) * p.scale)
}

func (p *painter) polygon(pts [][2]float64, c color.Color) {
	if len(pts) < 3 {
		return
	}
	b := p.img.Bounds()
	z := vector.NewRasterizer(b.Dx(), b.Dy())
	z.MoveTo(p.pt(pts[0][0], pts[0][1]))
	for _, q := range pts[1:] {
		z.LineTo(p.pt(q[0], q[1]))
	}
	z.ClosePath()
	z.Draw(p.img, b, image.NewUniform(c), image.Point{})
}

func (p *painter) rect(r [4]float64, c color.RGBA) {
	p.polygon([][2]float64{{r[0], r[1]}, {r[2], r[1]}, {r[2], r[3]}, {r[0], r[3]}}, c)
}

func (p *painter) strokeRect(r [4]float64, c color.RGBA, w float64) {
	p.polyline([][2]float64{{r[0], r[1]}, {r[2], r[1]}, {r[2], r[3]}, {r[0], r[3]}, {r[0], r[1]}}, w, c)
}

// polyline strokes a path of segments with round joins.
func (p *painter) polyline(pts [][2]float64, w float64, c color.Color) {
	for i := 0; i+1 < len(pts); i++ {
		a, b := pts[i], pts[i+1]
		dx, dy := b[0]-a[0], b[1]-a[1]
		l := math.Hypot(dx, dy)
		if l == 0 {
			continue
		}
		nx, ny := -dy/l*w/2, dx/l*w/2
		p.polygon([][2]float64{{a[0] + nx, a[1] + ny}, {b[0] + nx, b[1] + ny}, {b[0] - nx, b[1] - ny}, {a[0] - nx, a[1] - ny}}, c)
		if i > 0 {
			p.circle(a[0], a[1], w/2, c)
		}
	}
}

func (p *painter) circle(x, y, r float64, c color.Color) {
	var pts [][2]float64
	for k := 0; k < 24; k++ {
		a := 2 * math.Pi * float64(k) / 24
		pts = append(pts, [2]float64{x + r*math.Cos(a), y + r*math.Sin(a)})
	}
	p.polygon(pts, c)
}
