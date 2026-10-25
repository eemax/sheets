package engine

import (
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// lookupEq tells whether a value matches a lookup value exactly: same
// type, text without regard to case and with wildcards when wild.
func lookupEq(x, v Value, wild bool) bool {
	if x.Kind == Empty {
		x = num(0)
	}
	switch x.Kind {
	case Text:
		if v.Kind != Text {
			return false
		}
		if wild && strings.ContainsAny(x.Str, "*?~") {
			return wildMatch(foldCase(x.Str), foldCase(v.Str))
		}
		return foldCase(x.Str) == foldCase(v.Str)
	case Number:
		return v.Kind == Number && cmpFloat(x.Num, v.Num) == 0
	case Bool:
		return v.Kind == Bool && x.Bool == v.Bool
	}
	return false
}

// approxIndex finds, as Excel's binary search does on sorted values, the
// last position whose value is <= x (asc) or >= x (descending), among
// values of x's type; errors and other types are passed over. -1 when
// there is none.
func approxIndex(vals []Value, x Value, desc bool) int {
	if x.Kind == Empty {
		x = num(0)
	}
	var idx []int
	for i, v := range vals {
		if v.Kind == x.Kind {
			idx = append(idx, i)
		}
	}
	lo, hi, found := 0, len(idx)-1, -1
	for lo <= hi {
		mid := (lo + hi) / 2
		r := compare(vals[idx[mid]], x)
		if !desc && r <= 0 || desc && r >= 0 {
			found = idx[mid]
			lo = mid + 1
		} else {
			hi = mid - 1
		}
	}
	return found
}

// lineOf reads a value as a row or column of values; ok is false for a
// two-dimensional one.
func (c *ctx) lineOf(v Value) ([]Value, bool) {
	a := c.toArr(v)
	if a.Rows != 1 && a.Cols != 1 {
		return nil, false
	}
	return a.V, true
}

func init() {
	reg("MATCH", &fnSpec{min: 2, max: 3, args: []argKind{aVal, aRef, aVal}, f: func(c *ctx, a []Value) Value {
		x := a[0]
		if x.Kind == Error {
			return x
		}
		if a[1].Kind == Error {
			return a[1]
		}
		vals, ok := c.lineOf(a[1])
		if !ok {
			return errv(ErrNA)
		}
		typ := 1.0
		if len(a) > 2 {
			if isMissing(a[2]) {
				typ = 0
			} else {
				t, err := toNumber(a[2])
				if err != nil {
					return *err
				}
				typ = t
			}
		}
		switch {
		case typ == 0:
			for i, v := range vals {
				if lookupEq(x, v, true) {
					return num(float64(i + 1))
				}
			}
			return errv(ErrNA)
		case typ > 0:
			if i := approxIndex(vals, x, false); i >= 0 {
				return num(float64(i + 1))
			}
		default:
			if i := approxIndex(vals, x, true); i >= 0 {
				return num(float64(i + 1))
			}
		}
		return errv(ErrNA)
	}})
	vh := func(vertical bool) *fnSpec {
		return &fnSpec{min: 3, max: 4, args: []argKind{aVal, aRef, aVal, aVal}, f: func(c *ctx, a []Value) Value {
			x := a[0]
			if x.Kind == Error {
				return x
			}
			if a[1].Kind == Error {
				return a[1]
			}
			idx, err := toNumber(a[2])
			if err != nil {
				return *err
			}
			approx := true
			if len(a) > 3 {
				if isMissing(a[3]) {
					approx = false
				} else {
					b, err := toBool(a[3])
					if err != nil {
						return *err
					}
					approx = b
				}
			}
			t := c.toArr(a[1])
			n := int(idx)
			width := t.Cols
			if !vertical {
				width = t.Rows
			}
			if n < 1 {
				return errv(ErrValue)
			}
			if n > width {
				return errv(ErrRef)
			}
			length := t.Rows
			if !vertical {
				length = t.Cols
			}
			key := make([]Value, length)
			for i := range key {
				if vertical {
					key[i] = t.At(i, 0)
				} else {
					key[i] = t.At(0, i)
				}
			}
			pos := -1
			if approx {
				pos = approxIndex(key, x, false)
			} else {
				for i, v := range key {
					if lookupEq(x, v, true) {
						pos = i
						break
					}
				}
			}
			if pos < 0 {
				return errv(ErrNA)
			}
			if vertical {
				return t.At(pos, n-1)
			}
			return t.At(n-1, pos)
		}}
	}
	reg("VLOOKUP", vh(true))
	reg("HLOOKUP", vh(false))
	reg("INDEX", &fnSpec{min: 1, max: 4, args: []argKind{aArr, aVal, aVal, aVal}, f: indexFn})
	reg("LOOKUP", &fnSpec{min: 2, max: 3, args: []argKind{aVal, aArr, aArr}, f: func(c *ctx, a []Value) Value {
		x := a[0]
		if x.Kind == Error {
			return x
		}
		if a[1].Kind == Error {
			return a[1]
		}
		look := c.toArr(a[1])
		var key, res []Value
		if len(a) > 2 && !isMissing(a[2]) {
			key, _ = c.lineOf(a[1])
			if key == nil {
				key = look.V
			}
			r := c.toArr(a[2])
			res = r.V
		} else if look.Cols > look.Rows {
			// The array form: search the first row, return from the last.
			for j := 0; j < look.Cols; j++ {
				key = append(key, look.At(0, j))
				res = append(res, look.At(look.Rows-1, j))
			}
		} else {
			for i := 0; i < look.Rows; i++ {
				key = append(key, look.At(i, 0))
				res = append(res, look.At(i, look.Cols-1))
			}
		}
		i := approxIndex(key, x, false)
		if i < 0 {
			return errv(ErrNA)
		}
		if i >= len(res) {
			return errv(ErrNA)
		}
		return res[i]
	}})
	reg("XLOOKUP", &fnSpec{min: 3, max: 6, args: []argKind{aVal, aArr, aArr, aVal, aVal, aVal}, f: xlookup})
	reg("XMATCH", &fnSpec{min: 2, max: 4, args: []argKind{aVal, aArr, aVal, aVal}, f: func(c *ctx, a []Value) Value {
		x := a[0]
		if x.Kind == Error {
			return x
		}
		vals, ok := c.lineOf(a[1])
		if !ok {
			return errv(ErrValue)
		}
		mode, err := numArg(a, 2, 0)
		if err != nil {
			return *err
		}
		search, err := numArg(a, 3, 1)
		if err != nil {
			return *err
		}
		i := xfind(vals, x, int(mode), int(search))
		if i < 0 {
			return errv(ErrNA)
		}
		return num(float64(i + 1))
	}})
	reg("OFFSET", &fnSpec{min: 3, max: 5, args: []argKind{aRef, aRef, aRef, aRef, aRef}, f: func(c *ctx, a []Value) Value {
		if a[0].Kind == Error {
			return a[0]
		}
		if a[0].Kind != Ref {
			return errv(ErrValue)
		}
		for i := 1; i < len(a); i++ {
			if c.lifts(a[i]) {
				return c.fail(unsupported("OFFSET with several offsets at once gives several references, which the sheets calculator does not follow"))
			}
			if !isMissing(a[i]) {
				a[i] = c.scalar(a[i])
			}
		}
		base := a[0].Ref.Areas[0]
		dr, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		dc, err := toNumber(a[2])
		if err != nil {
			return *err
		}
		h, err := numArg(a, 3, float64(base.Rows()))
		if err != nil {
			return *err
		}
		w, err := numArg(a, 4, float64(base.Cols()))
		if err != nil {
			return *err
		}
		if int(h) == 0 || int(w) == 0 {
			return errv(ErrRef)
		}
		r1, c1 := base.R1+int(dr), base.C1+int(dc)
		r2, c2 := r1+int(h)-1, c1+int(w)-1
		if int(h) < 0 {
			r1, r2 = r1+int(h)+1, r1
		}
		if int(w) < 0 {
			c1, c2 = c1+int(w)+1, c1
		}
		if r1 < 1 || c1 < 1 || r2 > grid.MaxRows || c2 > grid.MaxCols {
			return errv(ErrRef)
		}
		return Value{Kind: Ref, Ref: &RefV{Sheet: a[0].Ref.Sheet, Areas: []Area{{r1, c1, r2, c2}}}}
	}})
	reg("INDIRECT", &fnSpec{min: 1, max: 2, f: func(c *ctx, a []Value) Value {
		text, err := toText(a[0])
		if err != nil {
			return *err
		}
		a1, err := boolArg(a, 1, true)
		if err != nil {
			return *err
		}
		return c.indirect(strings.TrimSpace(text), a1)
	}})
	reg("ADDRESS", &fnSpec{min: 2, max: 5, f: func(c *ctx, a []Value) Value {
		r, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		col, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		absNum, err := numArg(a, 2, 1)
		if err != nil {
			return *err
		}
		a1, err := boolArg(a, 3, true)
		if err != nil {
			return *err
		}
		ri, ci := int(r), int(col)
		if ri < 1 || ci < 1 || ri > grid.MaxRows || ci > grid.MaxCols {
			return errv(ErrValue)
		}
		var s string
		ab := int(absNum)
		if a1 {
			cs, rs := grid.ColName(ci), strconv.Itoa(ri)
			switch ab {
			case 1:
				s = "$" + cs + "$" + rs
			case 2:
				s = cs + "$" + rs
			case 3:
				s = "$" + cs + rs
			default:
				s = cs + rs
			}
		} else {
			rs, cs := "R"+strconv.Itoa(ri), "C"+strconv.Itoa(ci)
			if ab == 2 || ab == 4 {
				cs = "C[" + strconv.Itoa(ci) + "]"
			}
			if ab == 3 || ab == 4 {
				rs = "R[" + strconv.Itoa(ri) + "]"
			}
			s = rs + cs
		}
		if sheet, ok := opt(a, 4); ok {
			name, err := toText(sheet)
			if err != nil {
				return *err
			}
			s = grid.QuoteSheet(name) + "!" + s
		}
		return str(s)
	}})
	reg("TRANSPOSE", &fnSpec{min: 1, max: 1, args: []argKind{aArr}, f: func(c *ctx, a []Value) Value {
		if a[0].Kind == Error {
			return a[0]
		}
		x := c.toArr(a[0])
		out := newArr(x.Cols, x.Rows)
		for i := 0; i < x.Rows; i++ {
			for j := 0; j < x.Cols; j++ {
				out.V[j*x.Rows+i] = blankZero(x.At(i, j))
			}
		}
		return arrOf(out)
	}})
	reg("SEQUENCE", &fnSpec{min: 1, max: 4, f: func(c *ctx, a []Value) Value {
		r, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		cols, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		start, err := numArg(a, 2, 1)
		if err != nil {
			return *err
		}
		step, err := numArg(a, 3, 1)
		if err != nil {
			return *err
		}
		ri, ci := int(r), int(cols)
		if ri < 1 || ci < 1 || ri*ci > 1<<22 {
			return errv(ErrCalc)
		}
		out := newArr(ri, ci)
		for i := range out.V {
			out.V[i] = num(start + float64(i)*step)
		}
		return arrOf(out)
	}})
	reg("UNIQUE", &fnSpec{min: 1, max: 3, args: []argKind{aArr, aVal, aVal}, f: func(c *ctx, a []Value) Value {
		x := c.toArr(a[0])
		byCol, err := boolArg(a, 1, false)
		if err != nil {
			return *err
		}
		once, err := boolArg(a, 2, false)
		if err != nil {
			return *err
		}
		for i, v := range x.V {
			x.V[i] = blankZero(v)
		}
		items, n := slices(x, byCol)
		keyOf := func(vs []Value) string {
			var b strings.Builder
			for _, v := range vs {
				b.WriteString(strconv.Itoa(int(v.Kind)))
				t, _ := toText(v)
				b.WriteString(foldCase(t))
				b.WriteByte(0)
			}
			return b.String()
		}
		count := map[string]int{}
		for _, it := range items {
			count[keyOf(it)]++
		}
		seen := map[string]bool{}
		var keep [][]Value
		for _, it := range items {
			k := keyOf(it)
			if seen[k] || once && count[k] > 1 {
				continue
			}
			seen[k] = true
			keep = append(keep, it)
		}
		if len(keep) == 0 {
			return errv(ErrCalc)
		}
		return arrOf(unslices(keep, n, byCol))
	}})
	reg("SORT", &fnSpec{min: 1, max: 4, args: []argKind{aArr, aVal, aVal, aVal}, f: func(c *ctx, a []Value) Value {
		x := c.toArr(a[0])
		idx, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		order, err := numArg(a, 2, 1)
		if err != nil {
			return *err
		}
		byCol, err := boolArg(a, 3, false)
		if err != nil {
			return *err
		}
		items, n := slices(x, byCol)
		k := int(idx) - 1
		if k < 0 || k >= n {
			return errv(ErrValue)
		}
		sort.SliceStable(items, func(i, j int) bool {
			r := sortCompare(items[i][k], items[j][k])
			if order < 0 {
				return r > 0
			}
			return r < 0
		})
		return arrOf(unslices(items, n, byCol))
	}})
	reg("SORTBY", &fnSpec{min: 2, max: -1, args: []argKind{aArr, aArr, aVal}, f: func(c *ctx, a []Value) Value {
		x := c.toArr(a[0])
		type key struct {
			vals  []Value
			order float64
		}
		var keys []key
		for i := 1; i < len(a); i += 2 {
			by := c.toArr(a[i])
			order := 1.0
			if i+1 < len(a) {
				o, err := toNumber(c.scalar(a[i+1]))
				if err != nil {
					return *err
				}
				order = o
			}
			keys = append(keys, key{by.V, order})
		}
		rows := make([]int, x.Rows)
		for i := range rows {
			rows[i] = i
		}
		sort.SliceStable(rows, func(p, q int) bool {
			for _, k := range keys {
				if rows[p] >= len(k.vals) || rows[q] >= len(k.vals) {
					continue
				}
				r := sortCompare(k.vals[rows[p]], k.vals[rows[q]])
				if r != 0 {
					if k.order < 0 {
						return r > 0
					}
					return r < 0
				}
			}
			return false
		})
		out := newArr(x.Rows, x.Cols)
		for i, r := range rows {
			copy(out.V[i*x.Cols:(i+1)*x.Cols], x.V[r*x.Cols:(r+1)*x.Cols])
		}
		return arrOf(out)
	}})
	reg("FILTER", &fnSpec{min: 2, max: 3, args: []argKind{aArr, aArr, aVal}, f: func(c *ctx, a []Value) Value {
		if a[0].Kind == Error {
			return a[0]
		}
		if a[1].Kind == Error {
			return a[1]
		}
		x := c.toArr(a[0])
		inc := c.toArr(a[1])
		// A column of tests keeps rows; a row of tests keeps columns.
		var byCol bool
		switch {
		case inc.Cols == 1 && inc.Rows == x.Rows:
		case inc.Rows == 1 && inc.Cols == x.Cols:
			byCol = true
		default:
			return errv(ErrValue)
		}
		items, n := slices(x, byCol)
		var keep [][]Value
		for i, it := range items {
			v := inc.V[i]
			if v.Kind == Error {
				return v
			}
			b, err := toBool(v)
			if err != nil {
				return *err
			}
			if b {
				keep = append(keep, it)
			}
		}
		if len(keep) == 0 {
			if v, ok := opt(a, 2); ok {
				return v
			}
			return errv(ErrCalc)
		}
		return arrOf(unslices(keep, n, byCol))
	}})
	takeDrop := func(take bool) *fnSpec {
		return &fnSpec{min: 2, max: 3, args: []argKind{aArr, aVal, aVal}, f: func(c *ctx, a []Value) Value {
			x := c.toArr(a[0])
			// Left out: TAKE keeps every row or column, DROP drops none.
			defRows, defCols := float64(x.Rows), float64(x.Cols)
			if !take {
				defRows, defCols = 0, 0
			}
			r, err := numArg(a, 1, defRows)
			if err != nil {
				return *err
			}
			cc, err := numArg(a, 2, defCols)
			if err != nil {
				return *err
			}
			r1, r2 := span(x.Rows, int(r), take)
			c1, c2 := span(x.Cols, int(cc), take)
			if r2 <= r1 || c2 <= c1 {
				return errv(ErrCalc)
			}
			out := newArr(r2-r1, c2-c1)
			for i := r1; i < r2; i++ {
				for j := c1; j < c2; j++ {
					out.V[(i-r1)*(c2-c1)+j-c1] = x.At(i, j)
				}
			}
			return arrOf(out)
		}}
	}
	reg("TAKE", takeDrop(true))
	reg("DROP", takeDrop(false))
	chooseLines := func(cols bool) *fnSpec {
		return &fnSpec{min: 2, max: -1, args: []argKind{aArr, aVal}, f: func(c *ctx, a []Value) Value {
			x := c.toArr(a[0])
			items, n := slices(x, cols)
			var keep [][]Value
			for _, v := range a[1:] {
				k, err := toNumber(v)
				if err != nil {
					return *err
				}
				i := int(k)
				if i < 0 {
					i = len(items) + i + 1
				}
				if i < 1 || i > len(items) {
					return errv(ErrValue)
				}
				keep = append(keep, items[i-1])
			}
			return arrOf(unslices(keep, n, cols))
		}}
	}
	reg("CHOOSECOLS", chooseLines(true))
	reg("CHOOSEROWS", chooseLines(false))
	stack := func(vertical bool) *fnSpec {
		return &fnSpec{min: 1, max: -1, args: []argKind{aArr}, f: func(c *ctx, a []Value) Value {
			var parts []*Arr
			rows, cols := 0, 0
			for _, v := range a {
				x := c.toArr(v)
				parts = append(parts, x)
				if vertical {
					rows += x.Rows
					cols = max(cols, x.Cols)
				} else {
					cols += x.Cols
					rows = max(rows, x.Rows)
				}
			}
			out := newArr(rows, cols)
			for i := range out.V {
				out.V[i] = errv(ErrNA)
			}
			off := 0
			for _, x := range parts {
				for i := 0; i < x.Rows; i++ {
					for j := 0; j < x.Cols; j++ {
						if vertical {
							out.V[(off+i)*cols+j] = x.At(i, j)
						} else {
							out.V[i*cols+off+j] = x.At(i, j)
						}
					}
				}
				if vertical {
					off += x.Rows
				} else {
					off += x.Cols
				}
			}
			return arrOf(out)
		}}
	}
	reg("VSTACK", stack(true))
	reg("HSTACK", stack(false))
	toLine := func(col bool) *fnSpec {
		return &fnSpec{min: 1, max: 3, args: []argKind{aArr, aVal, aVal}, f: func(c *ctx, a []Value) Value {
			x := c.toArr(a[0])
			ignore, err := numArg(a, 1, 0)
			if err != nil {
				return *err
			}
			byCol, err := boolArg(a, 2, false)
			if err != nil {
				return *err
			}
			var vals []Value
			add := func(v Value) {
				switch {
				case int(ignore)&1 != 0 && v.Kind == Empty, int(ignore)&2 != 0 && v.Kind == Error:
					return
				}
				vals = append(vals, v)
			}
			if byCol {
				for j := 0; j < x.Cols; j++ {
					for i := 0; i < x.Rows; i++ {
						add(x.At(i, j))
					}
				}
			} else {
				for _, v := range x.V {
					add(v)
				}
			}
			var out *Arr
			if col {
				out = newArr(len(vals), 1)
			} else {
				out = newArr(1, len(vals))
			}
			copy(out.V, vals)
			return arrOf(out)
		}}
	}
	reg("TOCOL", toLine(true))
	reg("TOROW", toLine(false))
}

func blankZero(v Value) Value {
	if v.Kind == Empty {
		return num(0)
	}
	return v
}

// span gives the part TAKE or DROP keeps of n lines: k from the start, -k
// from the end.
func span(n, k int, take bool) (int, int) {
	switch {
	case take && k >= 0:
		return 0, min(k, n)
	case take:
		return max(0, n+k), n
	case k >= 0:
		return min(k, n), n
	}
	return 0, max(0, n+k)
}

// slices splits an array into rows (or columns, byCol).
func slices(x *Arr, byCol bool) ([][]Value, int) {
	var out [][]Value
	if byCol {
		for j := 0; j < x.Cols; j++ {
			col := make([]Value, x.Rows)
			for i := 0; i < x.Rows; i++ {
				col[i] = x.At(i, j)
			}
			out = append(out, col)
		}
		return out, x.Rows
	}
	for i := 0; i < x.Rows; i++ {
		out = append(out, append([]Value(nil), x.V[i*x.Cols:(i+1)*x.Cols]...))
	}
	return out, x.Cols
}

func unslices(items [][]Value, n int, byCol bool) *Arr {
	if byCol {
		out := newArr(n, len(items))
		for j, col := range items {
			for i := 0; i < n && i < len(col); i++ {
				out.V[i*len(items)+j] = col[i]
			}
		}
		return out
	}
	out := newArr(len(items), n)
	for i, row := range items {
		copy(out.V[i*n:], row)
	}
	return out
}

// sortCompare orders values as SORT does: numbers, text, logical values,
// errors, then blanks last.
func sortCompare(a, b Value) int {
	a, b = blankZero(a), blankZero(b)
	ra, rb := sortRank(a), sortRank(b)
	if ra != rb {
		if ra < rb {
			return -1
		}
		return 1
	}
	if a.Kind == Empty || a.Kind == Error {
		return 0
	}
	return compare(a, b)
}

func sortRank(v Value) int {
	switch v.Kind {
	case Number:
		return 0
	case Text:
		return 1
	case Bool:
		return 2
	case Error:
		return 3
	}
	return 4
}

// indexFn is INDEX(array, row, col, area): a reference into a reference,
// or a value of an array. Row or column 0 takes the whole column or row.
func indexFn(c *ctx, a []Value) Value {
	v := a[0]
	if v.Kind == Error {
		return v
	}
	rowV, hasRow := opt(a, 1)
	colV, hasCol := opt(a, 2)
	var r, col int
	if hasRow {
		if rowV.Kind == Error {
			return rowV
		}
		n, err := toNumber(rowV)
		if err != nil {
			return *err
		}
		r = int(n)
	}
	if hasCol {
		if colV.Kind == Error {
			return colV
		}
		n, err := toNumber(colV)
		if err != nil {
			return *err
		}
		col = int(n)
	}
	if r < 0 || col < 0 {
		return errv(ErrValue)
	}
	if v.Kind == Ref {
		areaN := 1
		if av, ok := opt(a, 3); ok {
			n, err := toNumber(av)
			if err != nil {
				return *err
			}
			areaN = int(n)
		}
		if areaN < 1 || areaN > len(v.Ref.Areas) {
			return errv(ErrRef)
		}
		ar := v.Ref.Areas[areaN-1]
		rows, cols := ar.Rows(), ar.Cols()
		// One index into a single row or column runs along it.
		if hasRow && !hasCol && rows == 1 && cols > 1 {
			r, col = 1, r
		}
		if !hasCol && cols == 1 && !hasRow {
			col = 1
		}
		if r > rows || col > cols {
			return errv(ErrRef)
		}
		out := ar
		if r > 0 {
			out.R1, out.R2 = ar.R1+r-1, ar.R1+r-1
		}
		if col > 0 {
			out.C1, out.C2 = ar.C1+col-1, ar.C1+col-1
		}
		if !hasCol && cols > 1 && r > 0 && rows > 1 {
			// INDEX(A1:C5, 2) on a two-dimensional range: the whole row.
			out.C1, out.C2 = ar.C1, ar.C2
		}
		return Value{Kind: Ref, Ref: &RefV{Sheet: v.Ref.Sheet, Areas: []Area{out}}}
	}
	x := c.toArr(v)
	if hasRow && !hasCol && x.Rows == 1 && x.Cols > 1 {
		r, col = 1, r
	}
	if !hasCol && x.Cols == 1 {
		col = 1
	}
	if r > x.Rows || col > x.Cols {
		return errv(ErrRef)
	}
	switch {
	case r > 0 && col > 0:
		return x.At(r-1, col-1)
	case r > 0:
		out := newArr(1, x.Cols)
		copy(out.V, x.V[(r-1)*x.Cols:r*x.Cols])
		return arrOf(out)
	case col > 0:
		out := newArr(x.Rows, 1)
		for i := 0; i < x.Rows; i++ {
			out.V[i] = x.At(i, col-1)
		}
		return arrOf(out)
	}
	return arrOf(x)
}

// xfind is XLOOKUP's and XMATCH's search: mode 0 exact, -1 exact or next
// smaller, 1 exact or next larger, 2 wildcards; search 1 from the first,
// -1 from the last, 2 and -2 binary.
func xfind(vals []Value, x Value, mode, search int) int {
	if search == 2 || search == -2 {
		if mode == 0 || mode == 2 {
			i := approxIndex(vals, x, search == -2)
			if i >= 0 && lookupEq(x, vals[i], false) {
				return i
			}
			return -1
		}
	}
	order := make([]int, len(vals))
	for i := range order {
		order[i] = i
	}
	if search == -1 {
		for i, j := 0, len(order)-1; i < j; i, j = i+1, j-1 {
			order[i], order[j] = order[j], order[i]
		}
	}
	for _, i := range order {
		if lookupEq(x, vals[i], mode == 2) {
			return i
		}
	}
	if mode == -1 || mode == 1 {
		best := -1
		for _, i := range order {
			v := vals[i]
			if v.Kind != x.Kind || v.Kind == Empty {
				continue
			}
			r := compare(v, x)
			if mode == -1 && r < 0 && (best < 0 || compare(v, vals[best]) > 0) ||
				mode == 1 && r > 0 && (best < 0 || compare(v, vals[best]) < 0) {
				best = i
			}
		}
		return best
	}
	return -1
}

func xlookup(c *ctx, a []Value) Value {
	x := a[0]
	if x.Kind == Error {
		return x
	}
	if a[1].Kind == Error {
		return a[1]
	}
	if a[2].Kind == Error {
		return a[2]
	}
	look := c.toArr(a[1])
	if look.Rows != 1 && look.Cols != 1 {
		return errv(ErrValue)
	}
	mode, err := numArg(a, 4, 0)
	if err != nil {
		return *err
	}
	search, err := numArg(a, 5, 1)
	if err != nil {
		return *err
	}
	i := xfind(look.V, x, int(mode), int(search))
	if i < 0 {
		if v, ok := opt(a, 3); ok {
			return v
		}
		return errv(ErrNA)
	}
	vertical := look.Cols == 1 && look.Rows > 1 || look.Rows == 1 && look.Cols == 1
	ret := a[2]
	if ret.Kind == Ref && len(ret.Ref.Areas) == 1 {
		ar := ret.Ref.Areas[0]
		out := ar
		if vertical {
			if i >= ar.Rows() {
				return errv(ErrValue)
			}
			out.R1, out.R2 = ar.R1+i, ar.R1+i
		} else {
			if i >= ar.Cols() {
				return errv(ErrValue)
			}
			out.C1, out.C2 = ar.C1+i, ar.C1+i
		}
		return Value{Kind: Ref, Ref: &RefV{Sheet: ret.Ref.Sheet, Areas: []Area{out}}}
	}
	r := c.toArr(ret)
	if vertical {
		if i >= r.Rows {
			return errv(ErrValue)
		}
		if r.Cols == 1 {
			return r.At(i, 0)
		}
		out := newArr(1, r.Cols)
		copy(out.V, r.V[i*r.Cols:(i+1)*r.Cols])
		return arrOf(out)
	}
	if i >= r.Cols {
		return errv(ErrValue)
	}
	if r.Rows == 1 {
		return r.At(0, i)
	}
	out := newArr(r.Rows, 1)
	for k := 0; k < r.Rows; k++ {
		out.V[k] = r.At(k, i)
	}
	return arrOf(out)
}

// indirect reads text as a reference: A1 style (Sheet!A1:B2, A:A, 1:1) or
// R1C1, or a defined name.
func (c *ctx) indirect(text string, a1 bool) Value {
	if text == "" {
		return errv(ErrRef)
	}
	if !a1 {
		return c.r1c1(text)
	}
	toks, err := lex(text)
	if err != nil || len(toks) < 2 {
		return errv(ErrRef)
	}
	p := &parser{toks: toks}
	n, err := p.expr(0)
	if err != nil || p.peek().kind != tEOF {
		return errv(ErrRef)
	}
	switch n.kind {
	case nRef, nName:
	case nBinary:
		if n.str != ":" {
			return errv(ErrRef)
		}
	default:
		return errv(ErrRef)
	}
	v := c.eval(n)
	if v.Kind != Ref {
		return errv(ErrRef)
	}
	return v
}

func (c *ctx) r1c1(text string) Value {
	sheet := c.sheet
	if i := strings.LastIndexByte(text, '!'); i >= 0 {
		sheet = strings.Trim(text[:i], "'")
		text = text[i+1:]
	}
	parse := func(s string) (int, int, bool) {
		s = strings.ToUpper(s)
		if !strings.HasPrefix(s, "R") {
			return 0, 0, false
		}
		k := strings.IndexByte(s, 'C')
		if k < 0 {
			return 0, 0, false
		}
		rel := func(part string, base int) (int, bool) {
			if part == "" {
				return base, true
			}
			if strings.HasPrefix(part, "[") && strings.HasSuffix(part, "]") {
				n, err := strconv.Atoi(part[1 : len(part)-1])
				return base + n, err == nil
			}
			n, err := strconv.Atoi(part)
			return n, err == nil
		}
		r, ok1 := rel(s[1:k], c.row)
		col, ok2 := rel(s[k+1:], c.col)
		return r, col, ok1 && ok2
	}
	parts := strings.Split(text, ":")
	r1, c1, ok := parse(parts[0])
	if !ok {
		return errv(ErrRef)
	}
	r2, c2 := r1, c1
	if len(parts) == 2 {
		if r2, c2, ok = parse(parts[1]); !ok {
			return errv(ErrRef)
		}
	}
	if _, _, ok := c.e.book.Extent(sheet); !ok {
		return errv(ErrRef)
	}
	return Value{Kind: Ref, Ref: &RefV{Sheet: sheet, Areas: []Area{norm(Area{r1, c1, r2, c2})}}}
}
