package engine

import (
	"math"
	"strings"
	"time"
)

// dateArg reads a date: a serial, or text Excel reads as a date. Negative
// serials and ones past 9999 are #NUM!.
func dateArg(v Value) (float64, *Value) {
	if v.Kind == Text {
		n, ok, dependent := readNumber(v.Str)
		switch {
		case dependent:
			e := languageDependent(v.Str)
			return 0, &e
		case ok:
			return n, nil
		}
		e := errv(ErrValue)
		return 0, &e
	}
	n, err := toNumber(v)
	if err != nil {
		return 0, err
	}
	if n < 0 || n >= 2958466 {
		e := errv(ErrNum)
		return 0, &e
	}
	return n, nil
}

func dateFn(min, max int, f func(c *ctx, d []float64, a []Value) Value) *fnSpec {
	return &fnSpec{min: min, max: max, f: func(c *ctx, a []Value) Value {
		d := make([]float64, 0, len(a))
		for i, v := range a {
			if i >= min && isMissing(v) {
				break
			}
			n, err := dateArg(v)
			if err != nil {
				return *err
			}
			d = append(d, n)
		}
		return f(c, d, a)
	}}
}

func init() {
	reg("DATE", &fnSpec{min: 3, max: 3, f: func(c *ctx, a []Value) Value {
		var n [3]float64
		for i := range n {
			x, err := toNumber(a[i])
			if err != nil {
				return *err
			}
			n[i] = x
		}
		return dateSerial(n[0], n[1], n[2])
	}})
	reg("TIME", &fnSpec{min: 3, max: 3, f: func(c *ctx, a []Value) Value {
		var n [3]float64
		for i := range n {
			x, err := toNumber(a[i])
			if err != nil {
				return *err
			}
			n[i] = math.Trunc(x)
		}
		secs := n[0]*3600 + n[1]*60 + n[2]
		if secs < 0 {
			return errv(ErrNum)
		}
		return num(math.Mod(secs, 86400) / 86400)
	}})
	reg("DATEVALUE", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		s, err := toText(a[0])
		if err != nil {
			return *err
		}
		n, ok, dependent := readNumber(s)
		switch {
		case dependent:
			return languageDependent(s)
		case ok:
			return num(math.Floor(n))
		}
		return notNumber(s)
	}})
	reg("TIMEVALUE", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		s, err := toText(a[0])
		if err != nil {
			return *err
		}
		n, ok, dependent := readNumber(s)
		switch {
		case dependent:
			return languageDependent(s)
		case ok:
			return num(n - math.Floor(n))
		}
		return notNumber(s)
	}})
	part := func(f func(v float64) float64) *fnSpec {
		return dateFn(1, 1, func(c *ctx, d []float64, a []Value) Value { return num(f(d[0])) })
	}
	reg("YEAR", part(func(v float64) float64 { y, _, _ := civil(v); return float64(y) }))
	reg("MONTH", part(func(v float64) float64 { _, m, _ := civil(v); return float64(m) }))
	reg("DAY", part(func(v float64) float64 { _, _, d := civil(v); return float64(d) }))
	reg("HOUR", part(func(v float64) float64 { h, _, _ := timeOf(v); return float64(h) }))
	reg("MINUTE", part(func(v float64) float64 { _, m, _ := timeOf(v); return float64(m) }))
	reg("SECOND", part(func(v float64) float64 { _, _, s := timeOf(v); return float64(s) }))
	reg("TODAY", &fnSpec{min: 0, max: 0, f: func(c *ctx, a []Value) Value {
		t := c.e.Now()
		return num(serialOf(t.Year(), int(t.Month()), t.Day()))
	}})
	reg("NOW", &fnSpec{min: 0, max: 0, f: func(c *ctx, a []Value) Value {
		t := c.e.Now()
		return num(serialOf(t.Year(), int(t.Month()), t.Day()) + float64(t.Hour()*3600+t.Minute()*60+t.Second())/86400)
	}})
	reg("WEEKDAY", dateFn(1, 2, func(c *ctx, d []float64, a []Value) Value {
		typ, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		w := weekday(d[0]) // 0 Sunday
		switch int(typ) {
		case 1, 17:
			return num(float64(w + 1))
		case 2, 11:
			return num(float64((w+6)%7 + 1))
		case 3:
			return num(float64((w + 6) % 7))
		case 12, 13, 14, 15, 16:
			start := int(typ) - 10 // 12: Tuesday is 1
			return num(float64((w-start+7)%7 + 1))
		}
		return errv(ErrNum)
	}))
	reg("WEEKNUM", dateFn(1, 2, func(c *ctx, d []float64, a []Value) Value {
		typ, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		t := int(typ)
		if t == 21 {
			return num(float64(isoWeek(d[0])))
		}
		start := map[int]int{1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0}
		s, ok := start[t]
		if !ok {
			return errv(ErrNum)
		}
		y, _, _ := civil(d[0])
		jan1 := serialOf(y, 1, 1)
		off := (weekday(jan1) - s + 7) % 7
		return num(math.Floor((math.Floor(d[0])-jan1+float64(off))/7) + 1)
	}))
	reg("ISOWEEKNUM", dateFn(1, 1, func(c *ctx, d []float64, a []Value) Value { return num(float64(isoWeek(d[0]))) }))
	months := func(end bool) *fnSpec {
		return &fnSpec{min: 2, max: 2, f: func(c *ctx, a []Value) Value {
			d, err := dateArg(a[0])
			if err != nil {
				return *err
			}
			m, err := toNumber(a[1])
			if err != nil {
				return *err
			}
			return addMonths(d, m, end)
		}}
	}
	reg("EDATE", months(false))
	reg("EOMONTH", months(true))
	reg("DAYS", dateFn(2, 2, func(c *ctx, d []float64, a []Value) Value {
		return num(math.Floor(d[0]) - math.Floor(d[1]))
	}))
	reg("DAYS360", dateFn(2, 3, func(c *ctx, d []float64, a []Value) Value {
		eu, err := boolArg(a, 2, false)
		if err != nil {
			return *err
		}
		return num(days360(d[0], d[1], eu))
	}))
	reg("DATEDIF", &fnSpec{min: 3, max: 3, f: func(c *ctx, a []Value) Value {
		s, err := dateArg(a[0])
		if err != nil {
			return *err
		}
		e, err := dateArg(a[1])
		if err != nil {
			return *err
		}
		unit, err := toText(a[2])
		if err != nil {
			return *err
		}
		s, e = math.Floor(s), math.Floor(e)
		if s > e {
			return errv(ErrNum)
		}
		y1, m1, d1 := civil(s)
		y2, m2, d2 := civil(e)
		months := (y2-y1)*12 + m2 - m1
		if d2 < d1 {
			months--
		}
		switch strings.ToUpper(unit) {
		case "Y":
			return num(float64(months / 12))
		case "M":
			return num(float64(months))
		case "D":
			return num(e - s)
		case "YM":
			return num(float64(months % 12))
		case "YD":
			// Days since the last anniversary.
			y := y2
			if m2 < m1 || m2 == m1 && d2 < d1 {
				y--
			}
			anniv, ok := civilSerial(y, m1, min(d1, daysIn(y, m1)))
			if !ok {
				return errv(ErrNum)
			}
			return num(e - anniv)
		case "MD":
			if d2 >= d1 {
				return num(float64(d2 - d1))
			}
			pm, py := m2-1, y2
			if pm < 1 {
				pm, py = 12, y2-1
			}
			return num(float64(daysIn(py, pm) - d1 + d2))
		}
		return errv(ErrNum)
	}})
	reg("YEARFRAC", dateFn(2, 3, func(c *ctx, d []float64, a []Value) Value {
		basis, err := numArg(a, 2, 0)
		if err != nil {
			return *err
		}
		return yearFrac(d[0], d[1], int(basis))
	}))
	reg("NETWORKDAYS", &fnSpec{min: 2, max: 3, args: []argKind{aVal, aVal, aRef}, f: func(c *ctx, a []Value) Value {
		return networkDays(c, a, "0000011")
	}})
	reg("NETWORKDAYS.INTL", &fnSpec{min: 2, max: 4, args: []argKind{aVal, aVal, aVal, aRef}, f: func(c *ctx, a []Value) Value {
		mask, err := weekendMask(a, 2)
		if err != nil {
			return *err
		}
		b := []Value{a[0], a[1]}
		if len(a) > 3 {
			b = append(b, a[3])
		}
		return networkDays(c, b, mask)
	}})
	reg("WORKDAY", &fnSpec{min: 2, max: 3, args: []argKind{aVal, aVal, aRef}, f: func(c *ctx, a []Value) Value {
		return workday(c, a, "0000011")
	}})
	reg("WORKDAY.INTL", &fnSpec{min: 2, max: 4, args: []argKind{aVal, aVal, aVal, aRef}, f: func(c *ctx, a []Value) Value {
		mask, err := weekendMask(a, 2)
		if err != nil {
			return *err
		}
		b := []Value{a[0], a[1]}
		if len(a) > 3 {
			b = append(b, a[3])
		}
		return workday(c, b, mask)
	}})
}

func isoWeek(serial float64) int {
	y, m, d := civil(serial)
	_, w := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC).ISOWeek()
	return w
}

func addMonths(start, months float64, endOfMonth bool) Value {
	y, m, d := civil(start)
	total := y*12 + m - 1 + int(math.Trunc(months))
	ny, nm := floorDiv(total, 12), total-floorDiv(total, 12)*12+1
	if ny < 1900 || ny > 9999 {
		return errv(ErrNum)
	}
	if endOfMonth {
		return num(serialOf(ny, nm, daysIn(ny, nm)))
	}
	return num(serialOf(ny, nm, min(d, daysIn(ny, nm))))
}

func days360(s, e float64, european bool) float64 {
	y1, m1, d1 := civil(s)
	y2, m2, d2 := civil(e)
	if european {
		if d1 == 31 {
			d1 = 30
		}
		if d2 == 31 {
			d2 = 30
		}
	} else {
		lastFeb := func(y, m, d int) bool { return m == 2 && d == daysIn(y, 2) }
		if lastFeb(y1, m1, d1) {
			if lastFeb(y2, m2, d2) {
				d2 = 30
			}
			d1 = 30
		}
		if d2 == 31 && d1 >= 30 {
			d2 = 30
		}
		if d1 == 31 {
			d1 = 30
		}
	}
	return float64((y2-y1)*360 + (m2-m1)*30 + d2 - d1)
}

func isLeap(y int) bool { return y%4 == 0 && (y%100 != 0 || y%400 == 0) }

func yearFrac(s, e float64, basis int) Value {
	s, e = math.Floor(s), math.Floor(e)
	if s > e {
		s, e = e, s
	}
	switch basis {
	case 0:
		y1, m1, d1 := civil(s)
		y2, m2, d2 := civil(e)
		if d1 == 31 && d2 == 31 {
			d1, d2 = 30, 30
		} else if d1 == 31 {
			d1 = 30
		} else if d1 == 30 && d2 == 31 {
			d2 = 30
		} else if m1 == 2 && m2 == 2 && d1 == daysIn(y1, 2) && d2 == daysIn(y2, 2) {
			d1, d2 = 30, 30
		} else if m1 == 2 && d1 == daysIn(y1, 2) {
			d1 = 30
		}
		return num(float64((y2-y1)*360+(m2-m1)*30+d2-d1) / 360)
	case 1:
		y1, m1, d1 := civil(s)
		y2, m2, d2 := civil(e)
		if y1 == y2 || y2 == y1+1 && (m1 > m2 || m1 == m2 && d1 >= d2) {
			days := 365.0
			if y1 == y2 && isLeap(y1) {
				days = 366
			} else if y1 != y2 {
				// A 29 February within the span makes it a leap year.
				if isLeap(y1) && (m1 < 3) || isLeap(y2) && (m2 > 2 || m2 == 2 && d2 == 29) {
					days = 366
				}
			}
			return num((e - s) / days)
		}
		total := 0.0
		for y := y1; y <= y2; y++ {
			if isLeap(y) {
				total += 366
			} else {
				total += 365
			}
		}
		return num((e - s) / (total / float64(y2-y1+1)))
	case 2:
		return num((e - s) / 360)
	case 3:
		return num((e - s) / 365)
	case 4:
		return num(days360(s, e, true) / 360)
	}
	return errv(ErrNum)
}

// weekendMask reads NETWORKDAYS.INTL's weekend: a number (1 Saturday and
// Sunday, 11 Sunday only...) or a seven-character mask from Monday.
func weekendMask(a []Value, i int) (string, *Value) {
	v, ok := opt(a, i)
	if !ok {
		return "0000011", nil
	}
	if v.Kind == Text {
		if len(v.Str) == 7 && strings.Trim(v.Str, "01") == "" && v.Str != "1111111" {
			return v.Str, nil
		}
		e := errv(ErrValue)
		return "", &e
	}
	n, err := toNumber(v)
	if err != nil {
		return "", err
	}
	masks := map[int]string{1: "0000011", 2: "1000001", 3: "1100000", 4: "0110000", 5: "0011000", 6: "0001100", 7: "0000110",
		11: "0000001", 12: "1000000", 13: "0100000", 14: "0010000", 15: "0001000", 16: "0000100", 17: "0000010"}
	m, ok := masks[int(n)]
	if !ok {
		e := errv(ErrNum)
		return "", &e
	}
	return m, nil
}

func holidays(c *ctx, v Value) (map[float64]bool, *Value) {
	h := map[float64]bool{}
	var bad *Value
	c.each(v, func(x Value, _ bool) bool {
		switch x.Kind {
		case Number:
			h[math.Floor(x.Num)] = true
		case Error:
			bad = &x
			return false
		case Text:
			if n, ok := parseDateTime(x.Str); ok {
				h[math.Floor(n)] = true
			}
		}
		return true
	})
	return h, bad
}

// isWorkday tells whether a serial is a working day under a weekend mask
// (Monday first).
func isWorkday(serial float64, mask string, hol map[float64]bool) bool {
	w := (weekday(serial) + 6) % 7 // 0 Monday
	return mask[w] == '0' && !hol[serial]
}

func networkDays(c *ctx, a []Value, mask string) Value {
	s, err := dateArg(a[0])
	if err != nil {
		return *err
	}
	e, err := dateArg(a[1])
	if err != nil {
		return *err
	}
	hol := map[float64]bool{}
	if len(a) > 2 && !isMissing(a[2]) {
		h, err := holidays(c, a[2])
		if err != nil {
			return *err
		}
		hol = h
	}
	s, e = math.Floor(s), math.Floor(e)
	sign := 1.0
	if s > e {
		s, e, sign = e, s, -1
	}
	n := 0
	for d := s; d <= e; d++ {
		if isWorkday(d, mask, hol) {
			n++
		}
	}
	return num(sign * float64(n))
}

func workday(c *ctx, a []Value, mask string) Value {
	s, err := dateArg(a[0])
	if err != nil {
		return *err
	}
	days, err := toNumber(a[1])
	if err != nil {
		return *err
	}
	hol := map[float64]bool{}
	if len(a) > 2 && !isMissing(a[2]) {
		h, err := holidays(c, a[2])
		if err != nil {
			return *err
		}
		hol = h
	}
	d := math.Floor(s)
	n := int(math.Trunc(days))
	step := 1.0
	if n < 0 {
		step, n = -1, -n
	}
	for n > 0 {
		d += step
		if isWorkday(d, mask, hol) {
			n--
		}
		if d < 0 || d > 2958465 {
			return errv(ErrNum)
		}
	}
	return num(d)
}
