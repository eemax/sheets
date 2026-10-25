package engine

import (
	"math"
)

// Financial functions, with Excel's conventions: money paid out is
// negative; type 1 pays at the start of each period.

func finArgs(a []Value, n int, defs ...float64) ([]float64, *Value) {
	out := make([]float64, len(a))
	for i := range out {
		def := 0.0
		if i >= n && i-n < len(defs) {
			def = defs[i-n]
		}
		x, err := numArg(a, i, def)
		if err != nil {
			return nil, err
		}
		out[i] = x
	}
	for len(out) < n+len(defs) {
		out = append(out, defs[len(out)-n])
	}
	return out, nil
}

func fin(n, max int, defs []float64, f func(x []float64) Value) *fnSpec {
	return &fnSpec{min: n, max: max, f: func(c *ctx, a []Value) Value {
		x, err := finArgs(a, n, defs...)
		if err != nil {
			return *err
		}
		return f(x)
	}}
}

// annuity: pv*(1+r)^n + pmt*(1+r*t)*((1+r)^n-1)/r + fv = 0
func pmtOf(r, n, pv, fv, t float64) float64 {
	if r == 0 {
		return -(pv + fv) / n
	}
	g := math.Pow(1+r, n)
	return -(pv*g + fv) * r / ((1 + r*t) * (g - 1))
}

func fvOf(r, n, pmt, pv, t float64) float64 {
	if r == 0 {
		return -(pv + pmt*n)
	}
	g := math.Pow(1+r, n)
	return -(pv*g + pmt*(1+r*t)*(g-1)/r)
}

func ipmtOf(r, per, n, pv, fv, t float64) float64 {
	p := pmtOf(r, n, pv, fv, t)
	var ip float64
	if per == 1 {
		if t == 1 {
			return 0
		}
		ip = -pv
	} else if t == 1 {
		ip = fvOf(r, per-2, p, pv, 1) - p
	} else {
		ip = fvOf(r, per-1, p, pv, 0)
	}
	return ip * r
}

func init() {
	reg("PMT", fin(3, 5, []float64{0, 0}, func(x []float64) Value {
		if x[1] == 0 {
			return errv(ErrNum)
		}
		return numOrErr(pmtOf(x[0], x[1], x[2], x[3], x[4]))
	}))
	reg("FV", fin(3, 5, []float64{0, 0}, func(x []float64) Value { return numOrErr(fvOf(x[0], x[1], x[2], x[3], x[4])) }))
	reg("PV", fin(3, 5, []float64{0, 0}, func(x []float64) Value {
		r, n, pmt, fv, t := x[0], x[1], x[2], x[3], x[4]
		if r == 0 {
			return num(-(fv + pmt*n))
		}
		g := math.Pow(1+r, n)
		return numOrErr(-(fv + pmt*(1+r*t)*(g-1)/r) / g)
	}))
	reg("NPER", fin(3, 5, []float64{0, 0}, func(x []float64) Value {
		r, pmt, pv, fv, t := x[0], x[1], x[2], x[3], x[4]
		if r == 0 {
			if pmt == 0 {
				return errv(ErrNum)
			}
			return num(-(pv + fv) / pmt)
		}
		z := pmt * (1 + r*t) / r
		v := (z - fv) / (z + pv)
		if v <= 0 {
			return errv(ErrNum)
		}
		return numOrErr(math.Log(v) / math.Log(1+r))
	}))
	reg("IPMT", fin(4, 6, []float64{0, 0}, func(x []float64) Value {
		r, per, n, pv, fv, t := x[0], x[1], x[2], x[3], x[4], x[5]
		if per < 1 || per > n {
			return errv(ErrNum)
		}
		return numOrErr(ipmtOf(r, per, n, pv, fv, t))
	}))
	reg("PPMT", fin(4, 6, []float64{0, 0}, func(x []float64) Value {
		r, per, n, pv, fv, t := x[0], x[1], x[2], x[3], x[4], x[5]
		if per < 1 || per > n {
			return errv(ErrNum)
		}
		return numOrErr(pmtOf(r, n, pv, fv, t) - ipmtOf(r, per, n, pv, fv, t))
	}))
	reg("RATE", fin(3, 6, []float64{0, 0, 0.1}, func(x []float64) Value {
		n, pmt, pv, fv, t, guess := x[0], x[1], x[2], x[3], x[4], x[5]
		f := func(r float64) float64 {
			if math.Abs(r) < 1e-12 {
				return pv + pmt*n + fv
			}
			g := math.Pow(1+r, n)
			return pv*g + pmt*(1+r*t)*(g-1)/r + fv
		}
		r, ok := newton(f, guess)
		if !ok {
			return errv(ErrNum)
		}
		return num(r)
	}))
	reg("NPV", &fnSpec{min: 2, max: -1, args: []argKind{aVal, aRef}, f: func(c *ctx, a []Value) Value {
		r, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		xs, err := c.numbers(a[1:], false)
		if err != nil {
			return *err
		}
		s := 0.0
		for i, x := range xs {
			s += x / math.Pow(1+r, float64(i+1))
		}
		return numOrErr(s)
	}})
	reg("IRR", &fnSpec{min: 1, max: 2, args: []argKind{aArr, aVal}, f: func(c *ctx, a []Value) Value {
		xs, err := c.numbers(a[:1], false)
		if err != nil {
			return *err
		}
		guess, err := numArg(a, 1, 0.1)
		if err != nil {
			return *err
		}
		f := func(r float64) float64 {
			s := 0.0
			for i, x := range xs {
				s += x / math.Pow(1+r, float64(i))
			}
			return s
		}
		r, ok := newton(f, guess)
		if !ok {
			return errv(ErrNum)
		}
		return num(r)
	}})
	reg("MIRR", &fnSpec{min: 3, max: 3, args: []argKind{aArr, aVal, aVal}, f: func(c *ctx, a []Value) Value {
		xs, err := c.numbers(a[:1], false)
		if err != nil {
			return *err
		}
		fr, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		rr, err := toNumber(a[2])
		if err != nil {
			return *err
		}
		n := float64(len(xs))
		var pos, neg float64
		for i, x := range xs {
			if x > 0 {
				pos += x * math.Pow(1+rr, n-float64(i)-1)
			} else {
				neg += x / math.Pow(1+fr, float64(i))
			}
		}
		if neg == 0 || pos == 0 {
			return errv(ErrDiv0)
		}
		return numOrErr(math.Pow(-pos/neg, 1/(n-1)) - 1)
	}})
	reg("XNPV", &fnSpec{min: 3, max: 3, args: []argKind{aVal, aArr, aArr}, f: func(c *ctx, a []Value) Value {
		r, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		vs, ds, bad := pairs(c, a[1], a[2])
		if bad != nil {
			return *bad
		}
		s := 0.0
		for i := range vs {
			s += vs[i] / math.Pow(1+r, (ds[i]-ds[0])/365)
		}
		return numOrErr(s)
	}})
	reg("XIRR", &fnSpec{min: 2, max: 3, args: []argKind{aArr, aArr, aVal}, f: func(c *ctx, a []Value) Value {
		vs, ds, bad := pairs(c, a[0], a[1])
		if bad != nil {
			return *bad
		}
		guess, err := numArg(a, 2, 0.1)
		if err != nil {
			return *err
		}
		f := func(r float64) float64 {
			s := 0.0
			for i := range vs {
				s += vs[i] / math.Pow(1+r, (ds[i]-ds[0])/365)
			}
			return s
		}
		r, ok := newton(f, guess)
		if !ok {
			return errv(ErrNum)
		}
		return num(r)
	}})
	reg("SLN", fin(3, 3, nil, func(x []float64) Value {
		if x[2] == 0 {
			return errv(ErrDiv0)
		}
		return num((x[0] - x[1]) / x[2])
	}))
	reg("SYD", fin(4, 4, nil, func(x []float64) Value {
		cost, salvage, life, per := x[0], x[1], x[2], x[3]
		if life <= 0 || per <= 0 || per > life {
			return errv(ErrNum)
		}
		return num((cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)))
	}))
	reg("DDB", fin(4, 5, []float64{2}, func(x []float64) Value {
		cost, salvage, life, per, factor := x[0], x[1], x[2], x[3], x[4]
		if life <= 0 || per <= 0 || per > life {
			return errv(ErrNum)
		}
		book := cost
		dep := 0.0
		for p := 1.0; p <= per; p++ {
			dep = math.Min(book*factor/life, math.Max(0, book-salvage))
			book -= dep
		}
		return num(dep)
	}))
	reg("EFFECT", fin(2, 2, nil, func(x []float64) Value {
		if x[0] <= 0 || x[1] < 1 {
			return errv(ErrNum)
		}
		n := math.Trunc(x[1])
		return num(math.Pow(1+x[0]/n, n) - 1)
	}))
	reg("NOMINAL", fin(2, 2, nil, func(x []float64) Value {
		if x[0] <= 0 || x[1] < 1 {
			return errv(ErrNum)
		}
		n := math.Trunc(x[1])
		return num(n * (math.Pow(1+x[0], 1/n) - 1))
	}))
}

func pairs(c *ctx, values, dates Value) ([]float64, []float64, *Value) {
	v, d := c.toArr(values), c.toArr(dates)
	if len(v.V) != len(d.V) {
		e := errv(ErrNum)
		return nil, nil, &e
	}
	var vs, ds []float64
	for i := range v.V {
		x, err := toNumber(v.V[i])
		if err != nil {
			return nil, nil, err
		}
		t, err := dateArg(d.V[i])
		if err != nil {
			return nil, nil, err
		}
		vs, ds = append(vs, x), append(ds, math.Floor(t))
	}
	if len(vs) < 2 {
		e := errv(ErrNum)
		return nil, nil, &e
	}
	return vs, ds, nil
}

// newton finds a root of f near guess, as Excel's iterative functions do.
func newton(f func(float64) float64, guess float64) (float64, bool) {
	r := guess
	for i := 0; i < 100; i++ {
		y := f(r)
		if math.IsNaN(y) || math.IsInf(y, 0) {
			return 0, false
		}
		// A central difference, for a slope good to the last digits.
		h := 1e-6 * math.Max(1e-3, math.Abs(r))
		d := (f(r+h) - f(r-h)) / (2 * h)
		if d == 0 {
			return 0, false
		}
		next := r - y/d
		if math.Abs(next-r) <= 1e-15*math.Max(1, math.Abs(next)) {
			return next, true
		}
		r = next
		if r <= -1 {
			r = -0.999999
		}
	}
	return 0, false
}
