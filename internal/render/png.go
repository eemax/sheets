package render

import (
	"image"
	"image/color"
	"image/draw"
	"image/png"
	"io"
	"math"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"golang.org/x/image/font"
)

var (
	white      = color.RGBA{255, 255, 255, 255}
	gridColor  = color.RGBA{0xDD, 0xDD, 0xDD, 255}
	headerFill = color.RGBA{0xF3, 0xF3, 0xF1, 255}
	headerInk  = color.RGBA{0x6B, 0x6A, 0x66, 255}
	headerLine = color.RGBA{0xCF, 0xCE, 0xC9, 255}
	buttonFill = color.RGBA{0xF7, 0xF7, 0xF7, 255}
	buttonLine = color.RGBA{0xAB, 0xAB, 0xAB, 255}
	buttonInk  = color.RGBA{0x40, 0x40, 0x40, 255}
)

const (
	headerW = 40 // row number column
	headerH = 20 // column letter row
	padX    = 3
)

// PNG draws a sheet (or range) to w.
func PNG(path string, w io.Writer, opts Options) (*Result, error) {
	sh, err := load(path, opts)
	if err != nil {
		return nil, err
	}
	img, res := sh.draw(opts)
	return res, png.Encode(w, img)
}

func (sh *sheet) result(width, height int) *Result {
	res := &Result{Sheet: sh.name, Range: sh.rg.String(), Charts: len(sh.charts), Width: width, Height: height}
	if sh.of != sh.rg {
		res.Of = sh.of.String()
	}
	if len(sh.missing) > 0 {
		res.Missing = strings.Join(sh.missing, "; ")
	}
	return res
}

func (sh *sheet) draw(opts Options) (*image.RGBA, *Result) {
	scale := opts.Scale
	if scale <= 0 {
		scale = 2
	}
	ox, oy := 0, 0
	if opts.Headers {
		ox, oy = headerW, headerH
	}
	gw, gh := sh.colX[len(sh.colX)-1], sh.rowY[len(sh.rowY)-1]
	for _, c := range sh.charts {
		gw, gh = max(gw, c.x+c.w), max(gh, c.y+c.h)
	}
	W, H := ox+gw+1, oy+gh+1
	img := image.NewRGBA(image.Rect(0, 0, int(float64(W)*scale), int(float64(H)*scale)))
	draw.Draw(img, img.Bounds(), image.NewUniform(white), image.Point{}, draw.Src)
	p := &painter{img: img, scale: scale, ox: ox, oy: oy}

	ncols, nrows := len(sh.colX)-1, len(sh.rowY)-1
	rect := func(r, c int) (int, int, int, int) {
		ce := sh.cells[[2]int{r, c}]
		x0, y0 := sh.colX[c-sh.rg.C1], sh.rowY[r-sh.rg.R1]
		x1, y1 := sh.colX[c-sh.rg.C1+1], sh.rowY[r-sh.rg.R1+1]
		if ce != nil && !ce.span.Empty() && ce.span.R2 > 0 {
			x1 = sh.edgeX(min(ce.span.C2, sh.rg.C2) + 1)
			y1 = sh.edgeY(min(ce.span.R2, sh.rg.R2) + 1)
		}
		return x0, y0, x1, y1
	}

	// Fills, then gridlines where cells have no fill, then borders.
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			ce := sh.cells[[2]int{r, c}]
			if ce != nil && !ce.covered && ce.look != nil && ce.look.fill != nil {
				x0, y0, x1, y1 := rect(r, c)
				p.fill(x0, y0, x1, y1, *ce.look.fill)
			}
		}
	}
	if sh.gridlines {
		for i := 0; i <= ncols; i++ {
			p.fill(sh.colX[i], 0, sh.colX[i]+1, gh, gridColor)
		}
		for i := 0; i <= nrows; i++ {
			p.fill(0, sh.rowY[i], gw, sh.rowY[i]+1, gridColor)
		}
		// Merged areas and filled cells hide the gridlines under them.
		for r := sh.rg.R1; r <= sh.rg.R2; r++ {
			for c := sh.rg.C1; c <= sh.rg.C2; c++ {
				ce := sh.cells[[2]int{r, c}]
				if ce == nil || ce.covered || ce.look == nil {
					continue
				}
				x0, y0, x1, y1 := rect(r, c)
				if ce.look.fill != nil {
					p.fill(x0+1, y0+1, x1, y1, *ce.look.fill)
				} else if ce.span.R2 > 0 && (ce.span.Rows() > 1 || ce.span.Cols() > 1) {
					p.fill(x0+1, y0+1, x1, y1, white)
				}
			}
		}
	}
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			ce := sh.cells[[2]int{r, c}]
			if ce == nil || ce.covered || ce.look == nil {
				continue
			}
			x0, y0, x1, y1 := rect(r, c)
			if x0 == x1 || y0 == y1 {
				continue // hidden
			}
			b := ce.look.borders
			if b[0].width > 0 {
				p.fill(x0, y0, x0+b[0].width, y1+1, b[0].color)
			}
			if b[1].width > 0 {
				p.fill(x0, y0, x1+1, y0+b[1].width, b[1].color)
			}
			if b[2].width > 0 {
				p.fill(x1+1-b[2].width, y0, x1+1, y1+1, b[2].color)
			}
			if b[3].width > 0 {
				p.fill(x0, y1+1-b[3].width, x1+1, y1+1, b[3].color)
			}
		}
	}

	// Text.
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			ce := sh.cells[[2]int{r, c}]
			if ce == nil || ce.covered || ce.text == "" {
				continue
			}
			lk := ce.look
			if lk == nil {
				lk = defaultLook
			}
			x0, y0, x1, y1 := rect(r, c)
			if x0 == x1 || y0 == y1 {
				continue // hidden
			}
			align := ce.align(lk)
			if align == "centerContinuous" {
				x1 = sh.colX[sh.continuous(r, c)-sh.rg.C1+1]
			}
			clipR := x1
			// Text too wide for its cell runs on over the empty cells to
			// its right, hiding their gridlines.
			last := sh.runsOn(r, c, align)
			for cc := c + 1; cc <= last; cc++ {
				if sh.gridlines && !sh.verticalBorder(r, cc) {
					bg := white
					if n := sh.cells[[2]int{r, cc}]; n != nil && n.look != nil && n.look.fill != nil {
						bg = *n.look.fill
					}
					p.fill(clipR, y0+1, clipR+1, y1, bg)
				}
				clipR = sh.colX[cc-sh.rg.C1+1]
			}
			p.text(ce, lk, x0, y0, x1, y1, clipR, align)
		}
	}

	// Filter buttons over the text, as Excel draws them.
	for cell := range sh.buttons {
		x0, y0, x1, y1 := rect(cell[0], cell[1])
		if x0 < x1 && y0 < y1 {
			p.filterButton(x1, y1)
		}
	}

	// Charts on top.
	for _, c := range sh.charts {
		p.chart(sh, c)
	}

	if opts.Headers {
		p.headers(sh, gw, gh)
	}
	return img, sh.result(W, H)
}

// align is how a cell's text sits: as set, or by the type of its value.
func (ce *cell) align(lk *look) string {
	if lk.halign == "" || lk.halign == "general" {
		if ce.number {
			return "right"
		}
		return "left"
	}
	return lk.halign
}

// runsOn returns the last column a cell's text spills into: left-aligned
// text too wide for its cell runs on over empty cells to its right.
func (sh *sheet) runsOn(r, c int, align string) int {
	ce := sh.cells[[2]int{r, c}]
	lk := ce.look
	if lk == nil {
		lk = defaultLook
	}
	if ce.text == "" || ce.number || lk.wrap || align != "left" || ce.span.Cols() > 1 {
		return c
	}
	w := textWidth(ce.text, lk) + 2*padX
	last := c
	for cc := c + 1; cc <= sh.rg.C2 && w > sh.colX[cc-sh.rg.C1]-sh.colX[c-sh.rg.C1]; cc++ {
		if n := sh.cells[[2]int{r, cc}]; n != nil && (n.text != "" || n.covered) {
			break
		}
		last = cc
	}
	return last
}

// verticalBorder reports whether a border is set on the left edge of a
// cell, by it or by the cell before it.
func (sh *sheet) verticalBorder(r, c int) bool {
	if n := sh.cells[[2]int{r, c}]; n != nil && n.look != nil && n.look.borders[0].width > 0 {
		return true
	}
	p := sh.cells[[2]int{r, c - 1}]
	return p != nil && p.look != nil && p.look.borders[2].width > 0
}

// painter draws in pixels at scale 1, offset by the headers.
type painter struct {
	img    *image.RGBA
	scale  float64
	ox, oy int
}

func (p *painter) px(v int) int { return int(math.Round(float64(v) * p.scale)) }

func (p *painter) fill(x0, y0, x1, y1 int, c color.RGBA) {
	r := image.Rect(p.px(x0+p.ox), p.px(y0+p.oy), p.px(x1+p.ox), p.px(y1+p.oy))
	draw.Draw(p.img, r, image.NewUniform(c), image.Point{}, draw.Over)
}

// text draws a cell's text inside its box, clipped at clipR.
func (p *painter) text(ce *cell, lk *look, x0, y0, x1, y1, clipR int, align string) {
	f := lookFace(lk, p.scale)
	clip := image.Rect(p.px(x0+p.ox), p.px(y0+p.oy), p.px(clipR+p.ox), p.px(y1+p.oy+1))
	d := &font.Drawer{Dst: p.img.SubImage(clip).(*image.RGBA), Src: image.NewUniform(lk.ink), Face: f}
	width := x1 - x0 - 2*padX
	lines := []string{ce.text}
	if lk.wrap {
		lines = wrapText(ce.text, lk, width)
	} else if ce.number && textWidth(ce.text, lk) > width {
		lines = []string{strings.Repeat("#", max(1, width/max(textWidth("#", lk), 1)))} // as Excel
	}
	lh := lineHeight(lk)
	asc := float64(f.Metrics().Ascent.Ceil()) / p.scale
	block := lh * float64(len(lines))
	top := float64(y1) - 3 - block // bottom aligned
	switch lk.valign {
	case "top":
		top = float64(y0) + 2
	case "center":
		top = float64(y0) + (float64(y1-y0)-block)/2
	}
	for i, line := range lines {
		w := float64(textWidth(line, lk))
		x := float64(x0 + padX)
		switch align {
		case "right":
			x = float64(x1-padX) - w
		case "center", "centerContinuous":
			x = float64(x0) + (float64(x1-x0)-w)/2
		}
		base := top + float64(i)*lh + asc + (lh-float64(f.Metrics().Height.Ceil())/p.scale)/2
		drawString(d, line, (x+float64(p.ox))*p.scale, (base+float64(p.oy))*p.scale)
	}
}

// filterButton draws a filter drop-down in the bottom-right corner of the
// cell ending at (x1, y1): a 16-pixel square, with a down arrow.
func (p *painter) filterButton(x1, y1 int) {
	const size = 16
	x0, y0 := x1-size, y1-size
	p.fill(x0, y0, x1, y1, buttonLine)
	p.fill(x0+1, y0+1, x1-1, y1-1, buttonFill)
	cx, cy := x0+size/2, y0+size/2-1
	for i := 0; i < 4; i++ {
		p.fill(cx-4+i, cy+i, cx+4-i, cy+i+1, buttonInk)
	}
}

// headers draws column letters and row numbers.
func (p *painter) headers(sh *sheet, gw, gh int) {
	ox, oy := p.ox, p.oy
	p.ox, p.oy = 0, 0
	defer func() { p.ox, p.oy = ox, oy }()
	p.fill(0, 0, ox+gw+1, oy, headerFill)
	p.fill(0, 0, ox, oy+gh+1, headerFill)
	p.fill(0, oy-1, ox+gw+1, oy, headerLine)
	p.fill(ox-1, 0, ox, oy+gh+1, headerLine)
	lk := &look{size: 9, ink: headerInk}
	f := lookFace(lk, p.scale)
	d := &font.Drawer{Dst: p.img, Src: image.NewUniform(headerInk), Face: f}
	asc := float64(f.Metrics().Ascent.Ceil()) / p.scale
	for i := 0; i+1 < len(sh.colX); i++ {
		if sh.colX[i] == sh.colX[i+1] {
			continue // hidden
		}
		name := grid.ColName(sh.rg.C1 + i)
		x := ox + (sh.colX[i]+sh.colX[i+1])/2 - textWidth(name, lk)/2
		drawString(d, name, float64(x)*p.scale, (float64(oy)/2+asc/2-1)*p.scale)
		p.fill(ox+sh.colX[i+1], 0, ox+sh.colX[i+1]+1, oy, headerLine)
	}
	for i := 0; i+1 < len(sh.rowY); i++ {
		if sh.rowY[i] == sh.rowY[i+1] {
			continue
		}
		num := itoa(sh.rg.R1 + i)
		x := ox/2 - textWidth(num, lk)/2
		y := oy + (sh.rowY[i]+sh.rowY[i+1])/2
		drawString(d, num, float64(x)*p.scale, (float64(y)+asc/2-1)*p.scale)
		p.fill(0, oy+sh.rowY[i+1], ox, oy+sh.rowY[i+1]+1, headerLine)
	}
}

func itoa(n int) string { return grid.FormatNumber(float64(n)) }
