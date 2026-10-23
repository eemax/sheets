package render

import (
	"math"
	"strings"
	"sync"

	"golang.org/x/image/font"
	"golang.org/x/image/font/gofont/gobold"
	"golang.org/x/image/font/gofont/gobolditalic"
	"golang.org/x/image/font/gofont/goitalic"
	"golang.org/x/image/font/gofont/goregular"
	"golang.org/x/image/font/opentype"
	"golang.org/x/image/math/fixed"
)

// The Go fonts stand in for Calibri, which is narrower: text is set a
// little smaller so widths come out close to Excel's.
const fontScale = 0.88

var (
	fontsOnce sync.Once
	fonts     [4]*opentype.Font // regular, bold, italic, bold italic
	faceMu    sync.Mutex
	faces     = map[faceKey]font.Face{}
)

type faceKey struct {
	bold, italic bool
	px           float64
}

func loadFonts() {
	fontsOnce.Do(func() {
		for i, data := range [][]byte{goregular.TTF, gobold.TTF, goitalic.TTF, gobolditalic.TTF} {
			f, err := opentype.Parse(data)
			if err != nil {
				panic(err) // embedded fonts always parse
			}
			fonts[i] = f
		}
	})
}

// face returns a font face for a size in points, at a pixel scale.
func face(bold, italic bool, pt, scale float64) font.Face {
	loadFonts()
	px := pt * 96 / 72 * fontScale * scale
	k := faceKey{bold, italic, px}
	faceMu.Lock()
	defer faceMu.Unlock()
	if f, ok := faces[k]; ok {
		return f
	}
	i := 0
	if bold {
		i |= 1
	}
	if italic {
		i |= 2
	}
	f, err := opentype.NewFace(fonts[i], &opentype.FaceOptions{Size: px, DPI: 72, Hinting: font.HintingFull})
	if err != nil {
		panic(err)
	}
	faces[k] = f
	return f
}

func lookFace(lk *look, scale float64) font.Face {
	return face(lk.bold, lk.italic, lk.size, scale)
}

// textWidth measures text in pixels at scale 1.
func textWidth(s string, lk *look) int {
	return font.MeasureString(lookFace(lk, 1), s).Ceil()
}

func lineHeight(lk *look) float64 {
	return lk.size * 96 / 72 * 1.25
}

// rowFit is the height Excel gives a row holding lines of text in a look:
// 20 pixels for a line of 11 point, more for larger type.
func rowFit(lk *look, lines int) int {
	return int(math.Round(float64(lines)*lineHeight(lk) + 2))
}

// wrapText breaks text into lines no wider than width pixels.
func wrapText(s string, lk *look, width int) []string {
	var out []string
	for _, para := range strings.Split(s, "\n") {
		words := strings.Fields(para)
		if len(words) == 0 {
			out = append(out, "")
			continue
		}
		line := words[0]
		for _, w := range words[1:] {
			if textWidth(line+" "+w, lk) <= width {
				line += " " + w
				continue
			}
			out = append(out, line)
			line = w
		}
		out = append(out, line)
	}
	return out
}

// drawString draws s with its baseline at (x, y) in pixels.
func drawString(d *font.Drawer, s string, x, y float64) {
	d.Dot = fixed.Point26_6{X: fixed.Int26_6(x * 64), Y: fixed.Int26_6(y * 64)}
	d.DrawString(s)
}

// truncateTo shortens s with an ellipsis to fit width pixels.
func truncateTo(s string, f font.Face, width int) string {
	if font.MeasureString(f, s).Ceil() <= width {
		return s
	}
	r := []rune(s)
	for len(r) > 0 && font.MeasureString(f, string(r)+"…").Ceil() > width {
		r = r[:len(r)-1]
	}
	if len(r) == 0 {
		return ""
	}
	return string(r) + "…"
}
