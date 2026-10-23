package render

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"html"
	"image"
	"image/color"
	"image/draw"
	"image/png"
	"io"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// HTML writes a sheet (or range) as a self-contained HTML page: a table
// carrying the cells' looks, and the charts as images.
func HTML(path string, w io.Writer, opts Options) (*Result, error) {
	sh, err := load(path, opts)
	if err != nil {
		return nil, err
	}
	var b strings.Builder
	css := func(c color.RGBA) string { return fmt.Sprintf("#%02x%02x%02x", c.R, c.G, c.B) }
	b.WriteString(`<!doctype html><html><head><meta charset="utf-8"><title>` + html.EscapeString(sh.name) + `</title><style>
body{margin:16px;background:#fff;font-family:Calibri,Carlito,"Segoe UI",system-ui,sans-serif;color:#000}
.wrap{position:relative;display:inline-block}
table{border-collapse:collapse;table-layout:fixed}
td{padding:0 3px;white-space:nowrap;overflow:hidden;vertical-align:bottom;font-size:11pt}
th{background:#f3f3f1;color:#6b6a66;font:9pt system-ui,sans-serif;font-weight:normal;border:1px solid #cfcec9}
.g td{border:1px solid #ddd}
img.chart{position:absolute}
td.fb{position:relative}
td.fb::after{content:"";position:absolute;right:0;bottom:0;width:14px;height:14px;border:1px solid #ababab;background:#f7f7f7 no-repeat center/8px 4px url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 8 4'%3E%3Cpath d='M0 0h8L4 4z' fill='%23404040'/%3E%3C/svg%3E")}
</style></head><body><div class="wrap">`)
	cls := ""
	if sh.gridlines {
		cls = ` class="g"`
	}
	// Fixed layout keeps the column widths only with the table's own set.
	tw := sh.colX[len(sh.colX)-1]
	if opts.Headers {
		tw += headerW
	}
	fmt.Fprintf(&b, `<table%s style="width:%dpx">`, cls, tw)
	b.WriteString("<colgroup>")
	if opts.Headers {
		fmt.Fprintf(&b, `<col style="width:%dpx">`, headerW)
	}
	for i := 0; i+1 < len(sh.colX); i++ {
		w := sh.colX[i+1] - sh.colX[i]
		if w == 0 {
			b.WriteString(`<col style="width:0;visibility:collapse">`)
			continue
		}
		fmt.Fprintf(&b, `<col style="width:%dpx">`, w)
	}
	b.WriteString("</colgroup>")
	if opts.Headers {
		b.WriteString("<tr><th></th>")
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			fmt.Fprintf(&b, "<th>%s</th>", grid.ColName(c))
		}
		b.WriteString("</tr>")
	}
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		if h := sh.rowY[r-sh.rg.R1+1] - sh.rowY[r-sh.rg.R1]; h > 0 {
			fmt.Fprintf(&b, `<tr style="height:%dpx">`, h)
		} else {
			b.WriteString(`<tr style="visibility:collapse">`)
		}
		if opts.Headers {
			fmt.Fprintf(&b, "<th>%d</th>", r)
		}
		skip := 0 // cells a centre-across-selection label covers
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			ce := sh.cells[[2]int{r, c}]
			if ce != nil && ce.covered || c <= skip {
				continue
			}
			var style []string
			attrs := ""
			text := ""
			if ce != nil {
				text = ce.text
				lk := ce.look
				if lk == nil {
					lk = defaultLook
				}
				if ce.span.R2 > 0 && (ce.span.Rows() > 1 || ce.span.Cols() > 1) {
					attrs = fmt.Sprintf(` rowspan="%d" colspan="%d"`, min(ce.span.R2, sh.rg.R2)-r+1, min(ce.span.C2, sh.rg.C2)-c+1)
				} else if lk.halign == "centerContinuous" && ce.text != "" {
					if skip = sh.continuous(r, c); skip > c {
						attrs = fmt.Sprintf(` colspan="%d"`, skip-c+1)
						// The right edge is the last covered cell's.
						if last := sh.cells[[2]int{r, skip}]; last.look != nil {
							lk2 := *lk
							lk2.borders[2] = last.look.borders[2]
							lk = &lk2
						}
					}
				}
				if lk.bold {
					style = append(style, "font-weight:bold")
				}
				if lk.italic {
					style = append(style, "font-style:italic")
				}
				if lk.size != 11 {
					style = append(style, fmt.Sprintf("font-size:%gpt", lk.size))
				}
				if lk.ink != (color.RGBA{0, 0, 0, 255}) {
					style = append(style, "color:"+css(lk.ink))
				}
				if lk.fill != nil {
					style = append(style, "background:"+css(*lk.fill))
				}
				for i, side := range []string{"left", "top", "right", "bottom"} {
					if bd := lk.borders[i]; bd.width > 0 {
						kind := "solid"
						if bd.width == 3 {
							kind = "double"
						}
						style = append(style, fmt.Sprintf("border-%s:%dpx %s %s", side, bd.width, kind, css(bd.color)))
					}
				}
				align := ce.align(lk)
				if align == "centerContinuous" {
					align = "center"
				}
				if align != "left" {
					style = append(style, "text-align:"+align)
				}
				if last := sh.runsOn(r, c, align); last > c {
					// Text runs on over the empty cells to its right.
					style = append(style, "overflow:visible")
					text = fmt.Sprintf(`<div style="width:%dpx;overflow:hidden">%s</div>`, sh.colX[last-sh.rg.C1+1]-sh.colX[c-sh.rg.C1]-2*padX, html.EscapeString(ce.text))
				} else {
					text = html.EscapeString(text)
				}
				if v := lk.valign; v == "top" || v == "center" {
					style = append(style, "vertical-align:"+map[string]string{"top": "top", "center": "middle"}[v])
				}
				if lk.wrap {
					style = append(style, "white-space:pre-wrap")
				}
			}
			s := ""
			if len(style) > 0 {
				s = ` style="` + strings.Join(style, ";") + `"`
			}
			if sh.buttons[[2]int{r, c}] {
				attrs += ` class="fb"`
			}
			fmt.Fprintf(&b, "<td%s%s>%s</td>", attrs, s, text)
		}
		b.WriteString("</tr>")
	}
	b.WriteString("</table>")
	ox, oy := 0, 0
	if opts.Headers {
		ox, oy = headerW+1, headerH+1
	}
	for _, c := range sh.charts {
		img := image.NewRGBA(image.Rect(0, 0, c.w*2, c.h*2))
		draw.Draw(img, img.Bounds(), image.NewUniform(white), image.Point{}, draw.Src)
		p := &painter{img: img, scale: 2, ox: -c.x, oy: -c.y}
		p.chart(sh, c)
		var buf bytes.Buffer
		if err := png.Encode(&buf, img); err != nil {
			return nil, err
		}
		fmt.Fprintf(&b, `<img class="chart" alt="%s" style="left:%dpx;top:%dpx;width:%dpx;height:%dpx" src="data:image/png;base64,%s">`,
			html.EscapeString(c.chart.Title), c.x+ox, c.y+oy, c.w, c.h, base64.StdEncoding.EncodeToString(buf.Bytes()))
	}
	b.WriteString("</div></body></html>\n")
	gw, gh := sh.colX[len(sh.colX)-1], sh.rowY[len(sh.rowY)-1]
	if _, err := io.WriteString(w, b.String()); err != nil {
		return nil, err
	}
	return sh.result(gw+ox, gh+oy), nil
}
