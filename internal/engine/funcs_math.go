package engine

import (
	"math"
	"math/rand"
	"sort"
	"strconv"
	"strings"
)

// collect gathers the numbers of aggregate arguments, as SUM does: in
// references and arrays only numbers count (text, logical values and
// blanks are skipped); typed directly, logical values and text that reads
// as a number count too, other text is #VALUE!. With all, text and
// logical values in references count as well (as 0 and 1), as the A
// functions do. Errors stop it, unless skipErr.
func (c *ctx) collect(a []Value, all, skipErr bool, fn func(x float64)) *Value {
	var bad *Value
	for _, v := range a {
		if isMissing(v) {
			continue
		}
		c.each(v, func(x Value, fromRef bool) bool {
			switch x.Kind {
			case Number:
				fn(x.Num)
			case Error:
				if skipErr {
					return true
				}
				bad = &x
				return false
			case Bool:
				if !fromRef || all {
					if x.Bool {
						fn(1)
					} else {
						fn(0)
					}
				}
			case Text:
				if fromRef {
					if all {
						fn(0)
					}
					return true
				}
				n, ok := textNumber(x.Str)
				if !ok {
					e := errv(ErrValue)
					bad = &e
					return false
				}
				fn(n)
			case Empty:
				if !fromRef && !all {
					fn(0)
				}
			}
			return true
		})
		if bad != nil {
			return bad
		}
	}
	return nil
}

func (c *ctx) numbers(a []Value, all bool) ([]float64, *Value) {
	var xs []float64
	err := c.collect(a, all, false, func(x float64) { xs = append(xs, x) })
	return xs, err
}

func aggregate(all bool, f func(xs []float64) Value) *fnSpec {
	return &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		xs, err := c.numbers(a, all)
		if err != nil {
			return *err
		}
		return f(xs)
	}}
}

func sum(xs []float64) float64 {
	s := 0.0
	for _, x := range xs {
		s = addRound(s, x)
	}
	return s
}

func mean(xs []float64) float64 { return sum(xs) / float64(len(xs)) }

func variance(xs []float64, sample bool) Value {
	n := float64(len(xs))
	if n == 0 || sample && n < 2 {
		return errv(ErrDiv0)
	}
	m := mean(xs)
	ss := 0.0
	for _, x := range xs {
		ss += (x - m) * (x - m)
	}
	if sample {
		return num(ss / (n - 1))
	}
	return num(ss / n)
}

func init() {
	reg("SUM", aggregate(false, func(xs []float64) Value { return num(sum(xs)) }))
	reg("SUMSQ", aggregate(false, func(xs []float64) Value {
		s := 0.0
		for _, x := range xs {
			s += x * x
		}
		return num(s)
	}))
	reg("PRODUCT", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return num(0)
		}
		p := 1.0
		for _, x := range xs {
			p *= x
		}
		return numOrErr(p)
	}))
	avg := func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrDiv0)
		}
		return num(mean(xs))
	}
	reg("AVERAGE", aggregate(false, avg))
	reg("AVERAGEA", aggregate(true, avg))
	extreme := func(isMax bool) func(xs []float64) Value {
		return func(xs []float64) Value {
			if len(xs) == 0 {
				return num(0)
			}
			m := xs[0]
			for _, x := range xs[1:] {
				if isMax && x > m || !isMax && x < m {
					m = x
				}
			}
			return num(m)
		}
	}
	reg("MAX", aggregate(false, extreme(true)))
	reg("MIN", aggregate(false, extreme(false)))
	reg("MAXA", aggregate(true, extreme(true)))
	reg("MINA", aggregate(true, extreme(false)))
	reg("MEDIAN", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrNum)
		}
		s := append([]float64(nil), xs...)
		sort.Float64s(s)
		n := len(s)
		if n%2 == 1 {
			return num(s[n/2])
		}
		return num((s[n/2-1] + s[n/2]) / 2)
	}))
	reg("MODE MODE.SNGL", aggregate(false, func(xs []float64) Value {
		best, count := 0.0, 1
		for i, x := range xs {
			n := 0
			for _, y := range xs[i:] {
				if y == x {
					n++
				}
			}
			if n > count {
				best, count = x, n
			}
		}
		if count < 2 {
			return errv(ErrNA)
		}
		return num(best)
	}))
	reg("VAR VAR.S", aggregate(false, func(xs []float64) Value { return variance(xs, true) }))
	reg("VARP VAR.P", aggregate(false, func(xs []float64) Value { return variance(xs, false) }))
	reg("VARA", aggregate(true, func(xs []float64) Value { return variance(xs, true) }))
	reg("VARPA", aggregate(true, func(xs []float64) Value { return variance(xs, false) }))
	sd := func(sample bool) func(xs []float64) Value {
		return func(xs []float64) Value {
			v := variance(xs, sample)
			if v.Kind == Error {
				return v
			}
			return num(math.Sqrt(v.Num))
		}
	}
	reg("STDEV STDEV.S", aggregate(false, sd(true)))
	reg("STDEVP STDEV.P", aggregate(false, sd(false)))
	reg("STDEVA", aggregate(true, sd(true)))
	reg("STDEVPA", aggregate(true, sd(false)))
	reg("AVEDEV", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrNum)
		}
		m := mean(xs)
		s := 0.0
		for _, x := range xs {
			s += math.Abs(x - m)
		}
		return num(s / float64(len(xs)))
	}))
	reg("DEVSQ", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrNum)
		}
		m := mean(xs)
		s := 0.0
		for _, x := range xs {
			s += (x - m) * (x - m)
		}
		return num(s)
	}))
	reg("GEOMEAN", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrNum)
		}
		s := 0.0
		for _, x := range xs {
			if x <= 0 {
				return errv(ErrNum)
			}
			s += math.Log(x)
		}
		return num(math.Exp(s / float64(len(xs))))
	}))
	reg("HARMEAN", aggregate(false, func(xs []float64) Value {
		if len(xs) == 0 {
			return errv(ErrNum)
		}
		s := 0.0
		for _, x := range xs {
			if x <= 0 {
				return errv(ErrNum)
			}
			s += 1 / x
		}
		return num(float64(len(xs)) / s)
	}))
	reg("COUNT", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		n := 0
		for _, v := range a {
			if isMissing(v) {
				continue
			}
			c.each(v, func(x Value, fromRef bool) bool {
				switch x.Kind {
				case Number:
					n++
				case Bool:
					if !fromRef {
						n++
					}
				case Text:
					if !fromRef {
						if _, ok := textNumber(x.Str); ok {
							n++
						}
					}
				}
				return true
			})
		}
		return num(float64(n))
	}})
	reg("COUNTA", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		n := 0
		for _, v := range a {
			if isMissing(v) {
				continue
			}
			c.each(v, func(x Value, fromRef bool) bool {
				if x.Kind != Empty || !fromRef {
					n++
				}
				return true
			})
		}
		return num(float64(n))
	}})
	reg("COUNTBLANK", &fnSpec{min: 1, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		if a[0].Kind != Ref {
			return errv(ErrValue)
		}
		n := 0
		total := 0
		for _, ar := range a[0].Ref.Areas {
			total += ar.Rows() * ar.Cols()
		}
		filled := 0
		c.each(a[0], func(x Value, _ bool) bool {
			if x.Kind == Text && x.Str == "" {
				n++
			}
			if x.Kind != Empty {
				filled++
			}
			return true
		})
		return num(float64(total - filled + n))
	}})
	kth := func(largest bool) *fnSpec {
		return &fnSpec{min: 2, max: 2, args: []argKind{aRef, aVal}, f: func(c *ctx, a []Value) Value {
			xs, err := c.numbers(a[:1], false)
			if err != nil {
				return *err
			}
			k, err := toNumber(a[1])
			if err != nil {
				return *err
			}
			ki := int(math.Ceil(k))
			if ki < 1 || ki > len(xs) {
				return errv(ErrNum)
			}
			s := append([]float64(nil), xs...)
			sort.Float64s(s)
			if largest {
				return num(s[len(s)-ki])
			}
			return num(s[ki-1])
		}}
	}
	reg("LARGE", kth(true))
	reg("SMALL", kth(false))
	reg("RANK RANK.EQ RANK.AVG", &fnSpec{min: 2, max: 3, args: []argKind{aVal, aRef, aVal}, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		xs, err := c.numbers(a[1:2], false)
		if err != nil {
			return *err
		}
		asc, err := numArg(a, 2, 0)
		if err != nil {
			return *err
		}
		before, same := 0, 0
		for _, y := range xs {
			switch {
			case y == x:
				same++
			case asc == 0 && y > x || asc != 0 && y < x:
				before++
			}
		}
		if same == 0 {
			return errv(ErrNA)
		}
		return num(float64(before + 1))
	}})
	pct := func(exclusive bool, quart bool) *fnSpec {
		return &fnSpec{min: 2, max: 2, args: []argKind{aRef, aVal}, f: func(c *ctx, a []Value) Value {
			xs, err := c.numbers(a[:1], false)
			if err != nil {
				return *err
			}
			k, err := toNumber(a[1])
			if err != nil {
				return *err
			}
			if quart {
				k = math.Trunc(k)
				if k < 0 || k > 4 || exclusive && (k <= 0 || k >= 4) {
					return errv(ErrNum)
				}
				k /= 4
			}
			return percentile(xs, k, exclusive)
		}}
	}
	reg("PERCENTILE PERCENTILE.INC", pct(false, false))
	reg("PERCENTILE.EXC", pct(true, false))
	reg("QUARTILE QUARTILE.INC", pct(false, true))
	reg("QUARTILE.EXC", pct(true, true))
	paired := func(f func(xs, ys []float64) Value) *fnSpec {
		return &fnSpec{min: 2, max: 2, args: []argKind{aArr, aArr}, f: func(c *ctx, a []Value) Value {
			x, y := c.toArr(a[0]), c.toArr(a[1])
			if len(x.V) != len(y.V) {
				return errv(ErrNA)
			}
			var xs, ys []float64
			for i := range x.V {
				if x.V[i].Kind == Error {
					return x.V[i]
				}
				if y.V[i].Kind == Error {
					return y.V[i]
				}
				if x.V[i].Kind == Number && y.V[i].Kind == Number {
					xs, ys = append(xs, x.V[i].Num), append(ys, y.V[i].Num)
				}
			}
			return f(xs, ys)
		}}
	}
	cov := func(xs, ys []float64) (float64, float64, float64) {
		mx, my := mean(xs), mean(ys)
		var sxy, sxx, syy float64
		for i := range xs {
			sxy += (xs[i] - mx) * (ys[i] - my)
			sxx += (xs[i] - mx) * (xs[i] - mx)
			syy += (ys[i] - my) * (ys[i] - my)
		}
		return sxy, sxx, syy
	}
	reg("CORREL PEARSON", paired(func(xs, ys []float64) Value {
		if len(xs) < 2 {
			return errv(ErrDiv0)
		}
		sxy, sxx, syy := cov(xs, ys)
		if sxx == 0 || syy == 0 {
			return errv(ErrDiv0)
		}
		return num(sxy / math.Sqrt(sxx*syy))
	}))
	reg("RSQ", paired(func(xs, ys []float64) Value {
		if len(xs) < 2 {
			return errv(ErrDiv0)
		}
		sxy, sxx, syy := cov(xs, ys)
		if sxx == 0 || syy == 0 {
			return errv(ErrDiv0)
		}
		r := sxy / math.Sqrt(sxx*syy)
		return num(r * r)
	}))
	// SLOPE(known_y, known_x)
	reg("SLOPE", paired(func(ys, xs []float64) Value {
		if len(xs) < 2 {
			return errv(ErrDiv0)
		}
		sxy, sxx, _ := cov(xs, ys)
		if sxx == 0 {
			return errv(ErrDiv0)
		}
		return num(sxy / sxx)
	}))
	reg("INTERCEPT", paired(func(ys, xs []float64) Value {
		if len(xs) < 2 {
			return errv(ErrDiv0)
		}
		sxy, sxx, _ := cov(xs, ys)
		if sxx == 0 {
			return errv(ErrDiv0)
		}
		return num(mean(ys) - sxy/sxx*mean(xs))
	}))
	reg("COVARIANCE.P COVAR", paired(func(xs, ys []float64) Value {
		if len(xs) == 0 {
			return errv(ErrDiv0)
		}
		sxy, _, _ := cov(xs, ys)
		return num(sxy / float64(len(xs)))
	}))
	reg("COVARIANCE.S", paired(func(xs, ys []float64) Value {
		if len(xs) < 2 {
			return errv(ErrDiv0)
		}
		sxy, _, _ := cov(xs, ys)
		return num(sxy / float64(len(xs)-1))
	}))
	reg("FORECAST FORECAST.LINEAR", &fnSpec{min: 3, max: 3, args: []argKind{aVal, aArr, aArr}, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		s := fns["SLOPE"].f(c, a[1:])
		i := fns["INTERCEPT"].f(c, a[1:])
		if s.Kind == Error {
			return s
		}
		if i.Kind == Error {
			return i
		}
		return num(i.Num + s.Num*x)
	}})

	reg("SUMPRODUCT", &fnSpec{min: 1, max: -1, args: []argKind{aArr}, f: func(c *ctx, a []Value) Value {
		arrs := make([]*Arr, len(a))
		for i, v := range a {
			if v.Kind == Error {
				return v
			}
			arrs[i] = c.toArr(v)
			if arrs[i].Rows != arrs[0].Rows || arrs[i].Cols != arrs[0].Cols {
				return errv(ErrValue)
			}
		}
		total := 0.0
		for k := range arrs[0].V {
			p := 1.0
			for _, ar := range arrs {
				x := ar.V[k]
				switch x.Kind {
				case Error:
					return x
				case Number:
					p *= x.Num
				default:
					p = 0
				}
			}
			total += p
		}
		return num(total)
	}})
	sumx := func(f func(x, y float64) float64) *fnSpec {
		return &fnSpec{min: 2, max: 2, args: []argKind{aArr, aArr}, f: func(c *ctx, a []Value) Value {
			x, y := c.toArr(a[0]), c.toArr(a[1])
			if len(x.V) != len(y.V) {
				return errv(ErrNA)
			}
			s := 0.0
			for i := range x.V {
				if x.V[i].Kind == Number && y.V[i].Kind == Number {
					s += f(x.V[i].Num, y.V[i].Num)
				}
			}
			return num(s)
		}}
	}
	reg("SUMX2MY2", sumx(func(x, y float64) float64 { return x*x - y*y }))
	reg("SUMX2PY2", sumx(func(x, y float64) float64 { return x*x + y*y }))
	reg("SUMXMY2", sumx(func(x, y float64) float64 { return (x - y) * (x - y) }))
	reg("MMULT", &fnSpec{min: 2, max: 2, args: []argKind{aArr, aArr}, f: func(c *ctx, a []Value) Value {
		x, y := c.toArr(a[0]), c.toArr(a[1])
		if x.Cols != y.Rows {
			return errv(ErrValue)
		}
		out := newArr(x.Rows, y.Cols)
		for i := 0; i < x.Rows; i++ {
			for j := 0; j < y.Cols; j++ {
				s := 0.0
				for k := 0; k < x.Cols; k++ {
					p, q := x.At(i, k), y.At(k, j)
					if p.Kind != Number || q.Kind != Number {
						out.V[i*y.Cols+j] = errv(ErrValue)
						goto next
					}
					s += p.Num * q.Num
				}
				out.V[i*y.Cols+j] = num(s)
			next:
			}
		}
		return arrOf(out)
	}})
	reg("FREQUENCY", &fnSpec{min: 2, max: 2, args: []argKind{aArr, aArr}, f: func(c *ctx, a []Value) Value {
		data, err := c.numbers(a[:1], false)
		if err != nil {
			return *err
		}
		bins, err := c.numbers(a[1:], false)
		if err != nil {
			return *err
		}
		order := make([]int, len(bins))
		for i := range order {
			order[i] = i
		}
		sort.SliceStable(order, func(i, j int) bool { return bins[order[i]] < bins[order[j]] })
		counts := make([]float64, len(bins)+1)
		for _, x := range data {
			placed := false
			for _, i := range order {
				if x <= bins[i] {
					counts[i]++
					placed = true
					break
				}
			}
			if !placed {
				counts[len(bins)]++
			}
		}
		out := newArr(len(counts), 1)
		for i, n := range counts {
			out.V[i] = num(n)
		}
		return arrOf(out)
	}})
	reg("SUBTOTAL", &fnSpec{min: 2, max: -1, args: []argKind{aVal, aRef}, f: func(c *ctx, a []Value) Value {
		f, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		return subtotal(c, int(f)%100, a[1:], false)
	}})
	reg("AGGREGATE", &fnSpec{min: 3, max: -1, args: []argKind{aVal, aVal, aArr, aArr}, f: aggregateFn})

	// Conditional aggregates.
	// Criteria are taken as given; ifs reads them (see there).
	reg("SUMIF", &fnSpec{min: 2, max: 3, args: []argKind{aRef, aRef, aRef}, f: func(c *ctx, a []Value) Value {
		target := a[0]
		if len(a) > 2 && !isMissing(a[2]) {
			target = a[2]
		}
		return ifs(c, "sum", target, a[0:1], a[1:2])
	}})
	reg("AVERAGEIF", &fnSpec{min: 2, max: 3, args: []argKind{aRef, aRef, aRef}, f: func(c *ctx, a []Value) Value {
		target := a[0]
		if len(a) > 2 && !isMissing(a[2]) {
			target = a[2]
		}
		return ifs(c, "average", target, a[0:1], a[1:2])
	}})
	reg("COUNTIF", &fnSpec{min: 2, max: 2, args: []argKind{aRef, aRef}, f: func(c *ctx, a []Value) Value {
		return ifs(c, "count", Value{}, a[0:1], a[1:2])
	}})
	multi := func(op string, hasTarget bool) *fnSpec {
		kinds := []argKind{aRef, aVal}
		min := 2
		if hasTarget {
			kinds = []argKind{aRef, aRef, aVal}
			min = 3
		}
		return &fnSpec{min: min, max: -1, args: kinds, f: func(c *ctx, a []Value) Value {
			target, rest := Value{}, a
			if hasTarget {
				target, rest = a[0], a[1:]
			}
			if len(rest)%2 != 0 {
				return errv(ErrValue)
			}
			var ranges, crits []Value
			for i := 0; i < len(rest); i += 2 {
				ranges = append(ranges, rest[i])
				crits = append(crits, rest[i+1])
			}
			return ifs(c, op, target, ranges, crits)
		}}
	}
	// Criteria ranges alternate with criteria: aRef, aVal, aRef, aVal...
	reg("COUNTIFS", withAlternating(multi("count", false), 0))
	reg("SUMIFS", withAlternating(multi("sum", true), 1))
	reg("AVERAGEIFS", withAlternating(multi("average", true), 1))
	reg("MAXIFS", withAlternating(multi("max", true), 1))
	reg("MINIFS", withAlternating(multi("min", true), 1))

	// Rounding and arithmetic.
	reg("ABS", fn1(func(x float64) Value { return num(math.Abs(x)) }))
	reg("SIGN", fn1(func(x float64) Value {
		switch {
		case x > 0:
			return num(1)
		case x < 0:
			return num(-1)
		}
		return num(0)
	}))
	reg("INT", fn1(func(x float64) Value { return num(math.Floor(roundSig(x, 15))) }))
	reg("TRUNC", fn2(1, 0, func(x, d float64) Value { return num(roundDigits(x, int(d), -1)) }))
	reg("ROUND", fn2(2, 0, func(x, d float64) Value { return num(roundDigits(x, int(d), 0)) }))
	reg("ROUNDUP", fn2(2, 0, func(x, d float64) Value { return num(roundDigits(x, int(d), 1)) }))
	reg("ROUNDDOWN", fn2(2, 0, func(x, d float64) Value { return num(roundDigits(x, int(d), -1)) }))
	reg("MROUND", fn2(2, 0, func(x, m float64) Value {
		if m == 0 {
			return num(0)
		}
		if x*m < 0 {
			return errv(ErrNum)
		}
		return num(roundDigits(x/m, 0, 0) * m)
	}))
	reg("CEILING", fn2(1, 1, func(x, s float64) Value {
		if s == 0 {
			return num(0)
		}
		if x > 0 && s < 0 {
			return errv(ErrNum)
		}
		return num(math.Ceil(roundSig(x/s, 15)) * s)
	}))
	reg("FLOOR", fn2(1, 1, func(x, s float64) Value {
		if s == 0 {
			if x == 0 {
				return num(0)
			}
			return errv(ErrDiv0)
		}
		if x > 0 && s < 0 {
			return errv(ErrNum)
		}
		return num(math.Floor(roundSig(x/s, 15)) * s)
	}))
	reg("CEILING.MATH", &fnSpec{min: 1, max: 3, f: func(c *ctx, a []Value) Value { return mathRound(a, true) }})
	reg("FLOOR.MATH", &fnSpec{min: 1, max: 3, f: func(c *ctx, a []Value) Value { return mathRound(a, false) }})
	reg("CEILING.PRECISE ISO.CEILING", fn2(1, 1, func(x, s float64) Value {
		s = math.Abs(s)
		if s == 0 {
			return num(0)
		}
		return num(math.Ceil(roundSig(x/s, 15)) * s)
	}))
	reg("FLOOR.PRECISE", fn2(1, 1, func(x, s float64) Value {
		s = math.Abs(s)
		if s == 0 {
			return num(0)
		}
		return num(math.Floor(roundSig(x/s, 15)) * s)
	}))
	reg("MOD", fn2(2, 0, func(x, d float64) Value {
		if d == 0 {
			return errv(ErrDiv0)
		}
		r := x - d*math.Floor(roundSig(x/d, 15))
		if math.Abs(r) < 1e-15*math.Max(1, math.Abs(x)) {
			r = 0
		}
		return num(r)
	}))
	reg("QUOTIENT", fn2(2, 0, func(x, d float64) Value {
		if d == 0 {
			return errv(ErrDiv0)
		}
		return num(math.Trunc(x / d))
	}))
	reg("POWER", fn2(2, 0, func(x, y float64) Value { return binop("^", num(x), num(y)) }))
	reg("SQRT", fn1(func(x float64) Value {
		if x < 0 {
			return errv(ErrNum)
		}
		return num(math.Sqrt(x))
	}))
	reg("SQRTPI", fn1(func(x float64) Value {
		if x < 0 {
			return errv(ErrNum)
		}
		return num(math.Sqrt(x * math.Pi))
	}))
	reg("EXP", fn1(func(x float64) Value { return numOrErr(math.Exp(x)) }))
	reg("LN", fn1(func(x float64) Value {
		if x <= 0 {
			return errv(ErrNum)
		}
		return num(math.Log(x))
	}))
	reg("LOG10", fn1(func(x float64) Value {
		if x <= 0 {
			return errv(ErrNum)
		}
		return num(math.Log10(x))
	}))
	reg("LOG", fn2(1, 10, func(x, b float64) Value {
		if x <= 0 || b <= 0 {
			return errv(ErrNum)
		}
		if b == 1 {
			return errv(ErrDiv0)
		}
		return num(math.Log(x) / math.Log(b))
	}))
	reg("PI", &fnSpec{min: 0, max: 0, f: func(*ctx, []Value) Value { return num(math.Pi) }})
	reg("RAND", &fnSpec{min: 0, max: 0, f: func(*ctx, []Value) Value { return num(rand.Float64()) }})
	reg("RANDBETWEEN", fn2(2, 0, func(lo, hi float64) Value {
		lo, hi = math.Ceil(lo), math.Floor(hi)
		if lo > hi {
			return errv(ErrNum)
		}
		return num(lo + math.Floor(rand.Float64()*(hi-lo+1)))
	}))
	reg("EVEN", fn1(func(x float64) Value {
		r := math.Ceil(math.Abs(x)/2) * 2
		return num(math.Copysign(r, x))
	}))
	reg("ODD", fn1(func(x float64) Value {
		r := math.Ceil((math.Abs(x)-1)/2)*2 + 1
		if r < 1 {
			r = 1
		}
		if x < 0 {
			return num(-r)
		}
		return num(r)
	}))
	reg("FACT", fn1(func(x float64) Value {
		if x < 0 {
			return errv(ErrNum)
		}
		r := 1.0
		for i := 2.0; i <= math.Trunc(x); i++ {
			r *= i
		}
		return numOrErr(r)
	}))
	reg("COMBIN", fn2(2, 0, func(n, k float64) Value {
		n, k = math.Trunc(n), math.Trunc(k)
		if n < 0 || k < 0 || k > n {
			return errv(ErrNum)
		}
		r := 1.0
		for i := 0.0; i < k; i++ {
			r = r * (n - i) / (i + 1)
		}
		return num(math.Round(r))
	}))
	reg("PERMUT", fn2(2, 0, func(n, k float64) Value {
		n, k = math.Trunc(n), math.Trunc(k)
		if n < 0 || k < 0 || k > n {
			return errv(ErrNum)
		}
		r := 1.0
		for i := 0.0; i < k; i++ {
			r *= n - i
		}
		return num(r)
	}))
	gcd := func(a, b float64) float64 {
		for b != 0 {
			a, b = b, math.Mod(a, b)
		}
		return a
	}
	reg("GCD", aggregate(false, func(xs []float64) Value {
		g := 0.0
		for _, x := range xs {
			if x < 0 {
				return errv(ErrNum)
			}
			g = gcd(g, math.Trunc(x))
		}
		return num(g)
	}))
	reg("LCM", aggregate(false, func(xs []float64) Value {
		l := 1.0
		for _, x := range xs {
			x = math.Trunc(x)
			if x < 0 {
				return errv(ErrNum)
			}
			if x == 0 {
				return num(0)
			}
			l = l * x / gcd(l, x)
		}
		return num(l)
	}))
	reg("DEGREES", fn1(func(x float64) Value { return num(x * 180 / math.Pi) }))
	reg("RADIANS", fn1(func(x float64) Value { return num(x * math.Pi / 180) }))
	for name, f := range map[string]func(float64) float64{"SIN": math.Sin, "COS": math.Cos, "TAN": math.Tan, "ATAN": math.Atan,
		"SINH": math.Sinh, "COSH": math.Cosh, "TANH": math.Tanh, "ASINH": math.Asinh} {
		f := f
		reg(name, fn1(func(x float64) Value { return numOrErr(f(x)) }))
	}
	reg("ASIN", fn1(func(x float64) Value {
		if x < -1 || x > 1 {
			return errv(ErrNum)
		}
		return num(math.Asin(x))
	}))
	reg("ACOS", fn1(func(x float64) Value {
		if x < -1 || x > 1 {
			return errv(ErrNum)
		}
		return num(math.Acos(x))
	}))
	reg("ATAN2", fn2(2, 0, func(x, y float64) Value {
		if x == 0 && y == 0 {
			return errv(ErrDiv0)
		}
		return num(math.Atan2(y, x))
	}))
}

// withAlternating sets an *IFS function's argument kinds: after the first
// skip arguments, ranges and criteria alternate.
func withAlternating(s *fnSpec, skip int) *fnSpec {
	s.args = nil
	kinds := make([]argKind, 0, 255)
	for i := 0; i < skip; i++ {
		kinds = append(kinds, aRef)
	}
	for len(kinds) < 255 {
		kinds = append(kinds, aRef, aRef)
	}
	s.args = kinds
	return s
}

// roundDigits rounds to d decimal places (negative: to tens, hundreds):
// mode 0 half away from zero, 1 away from zero, -1 toward zero. It works
// on the number to 15 significant digits, so 2.675 rounds to 2.68 as in
// Excel.
func roundDigits(x float64, d int, mode int) float64 {
	if x == 0 || math.IsInf(x, 0) || math.IsNaN(x) {
		return x
	}
	if d > 15 {
		d = 15
	}
	scale := math.Pow(10, float64(abs(d)))
	var f float64
	if d >= 0 {
		f = roundSig(x*scale, 15)
	} else {
		f = roundSig(x/scale, 15)
	}
	var r float64
	switch mode {
	case 0:
		r = math.Floor(math.Abs(f) + 0.5)
	case 1:
		r = math.Ceil(math.Abs(f))
	default:
		r = math.Floor(math.Abs(f))
	}
	r = math.Copysign(r, f)
	if d >= 0 {
		// Dividing gives the nearest double to the decimal result.
		v, _ := strconv.ParseFloat(strconv.FormatFloat(r/scale, 'g', 15, 64), 64)
		return v
	}
	return r * scale
}

func abs(n int) int {
	if n < 0 {
		return -n
	}
	return n
}

func mathRound(a []Value, ceil bool) Value {
	x, err := toNumber(a[0])
	if err != nil {
		return *err
	}
	s, err := numArg(a, 1, 1)
	if err != nil {
		return *err
	}
	mode, err := numArg(a, 2, 0)
	if err != nil {
		return *err
	}
	s = math.Abs(s)
	if s == 0 {
		return num(0)
	}
	q := roundSig(x/s, 15)
	// Positive numbers round up (CEILING) or down (FLOOR); negative ones
	// the same way unless mode is set, which turns them around.
	up := ceil
	if x < 0 && mode != 0 {
		up = !up
	}
	r := math.Floor(q)
	if up {
		r = math.Ceil(q)
	}
	return num(r * s)
}

func percentile(xs []float64, k float64, exclusive bool) Value {
	n := float64(len(xs))
	if n == 0 || k < 0 || k > 1 {
		return errv(ErrNum)
	}
	s := append([]float64(nil), xs...)
	sort.Float64s(s)
	var pos float64
	if exclusive {
		pos = k*(n+1) - 1
		if pos < 0 || pos > n-1 {
			return errv(ErrNum)
		}
	} else {
		pos = k * (n - 1)
	}
	lo := math.Floor(pos)
	frac := pos - lo
	i := int(lo)
	if i+1 >= len(s) {
		return num(s[len(s)-1])
	}
	return num(s[i] + frac*(s[i+1]-s[i]))
}

// subtotal is SUBTOTAL's function f (1-11) over references, skipping
// cells that hold SUBTOTAL or AGGREGATE formulas themselves.
func subtotal(c *ctx, f int, refs []Value, skipErr bool) Value {
	var vals []Value
	for _, r := range refs {
		if r.Kind == Error {
			return r
		}
		if r.Kind != Ref {
			vals = append(vals, r)
			continue
		}
		for _, area := range r.Ref.Areas {
			ar := c.bounded(r.Ref.Sheet, area)
			sub := newArr(max(ar.Rows(), 0), max(ar.Cols(), 0))
			for i := 0; i < sub.Rows; i++ {
				for j := 0; j < sub.Cols; j++ {
					row, col := ar.R1+i, ar.C1+j
					if _, text := c.e.book.Cell(r.Ref.Sheet, row, col); text != "" {
						up := strings.ToUpper(text)
						if strings.Contains(up, "SUBTOTAL(") || strings.Contains(up, "AGGREGATE(") {
							continue
						}
					}
					sub.V[i*sub.Cols+j] = c.cell(r.Ref.Sheet, row, col)
				}
			}
			vals = append(vals, arrOf(sub))
		}
	}
	if skipErr {
		for _, v := range vals {
			if v.Kind == Array {
				for i, x := range v.Arr.V {
					if x.Kind == Error {
						v.Arr.V[i] = Value{}
					}
				}
			}
		}
	}
	names := []string{"", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PRODUCT", "STDEV", "STDEVP", "SUM", "VAR", "VARP"}
	if f < 1 || f > 11 {
		return errv(ErrValue)
	}
	return fns[names[f]].f(c, vals)
}

func aggregateFn(c *ctx, a []Value) Value {
	f, err := toNumber(a[0])
	if err != nil {
		return *err
	}
	opt, err := toNumber(a[1])
	if err != nil {
		return *err
	}
	o := int(opt)
	skipErr := o == 2 || o == 3 || o == 6 || o == 7
	fi := int(f)
	if fi >= 1 && fi <= 13 {
		names := []string{"", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PRODUCT", "STDEV.S", "STDEV.P", "SUM", "VAR.S", "VAR.P", "MEDIAN", "MODE.SNGL"}
		if fi <= 11 && o <= 3 {
			return subtotal(c, fi, a[2:], skipErr)
		}
		var vals []Value
		for _, v := range a[2:] {
			vals = append(vals, filterErrors(c, v, skipErr))
		}
		return fns[names[fi]].f(c, vals)
	}
	if fi < 14 || fi > 19 || len(a) < 4 {
		return errv(ErrValue)
	}
	data := filterErrors(c, a[2], skipErr)
	k := c.scalar(a[3])
	names := map[int]string{14: "LARGE", 15: "SMALL", 16: "PERCENTILE.INC", 17: "QUARTILE.INC", 18: "PERCENTILE.EXC", 19: "QUARTILE.EXC"}
	if isMulti(a[3]) {
		ka := c.toArr(a[3])
		out := newArr(ka.Rows, ka.Cols)
		for i, kv := range ka.V {
			out.V[i] = fns[names[fi]].f(c, []Value{data, kv})
		}
		return arrOf(out)
	}
	return fns[names[fi]].f(c, []Value{data, k})
}

// filterErrors reads a value as an array, without its errors when skip.
func filterErrors(c *ctx, v Value, skip bool) Value {
	if !skip {
		return v
	}
	a := c.toArr(v)
	out := newArr(a.Rows, a.Cols)
	for i, x := range a.V {
		if x.Kind != Error {
			out.V[i] = x
		}
	}
	return arrOf(out)
}
