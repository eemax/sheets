package engine

import (
	"math"
	"strconv"
	"strings"
	"time"
)

// Dates are serial numbers: days since 1899-12-30 in the 1900 system,
// where Excel also counts a 29 February 1900 that never was (serial 60), so
// serials from 61 on are one day ahead of a plain count.

// serialOf gives the serial of a civil date.
func serialOf(y, m, d int) float64 {
	t := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	days := math.Floor(t.Sub(time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)).Hours()/24 + 0.5)
	if days >= 60 { // on or after 1 March 1900, past the phantom 29 February
		days++
	}
	return days
}

// civil gives the date of a serial (its whole part): year, month, day.
// Serial 60 is the phantom 29 February 1900; serial 0 is 0 January 1900.
func civil(serial float64) (y, m, d int) {
	n := int(math.Floor(serial))
	switch {
	case n == 60:
		return 1900, 2, 29
	case n == 0:
		return 1900, 1, 0
	case n > 60:
		n--
	}
	t := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC).AddDate(0, 0, n)
	return t.Year(), int(t.Month()), t.Day()
}

// timeOf gives the time of day of a serial: hour, minute, second, rounded
// to the second as Excel shows it.
func timeOf(serial float64) (h, mi, s int) {
	frac := serial - math.Floor(serial)
	secs := int(math.Floor(frac*86400 + 0.5))
	if secs >= 86400 {
		secs = 86399
	}
	return secs / 3600, secs / 60 % 60, secs % 60
}

func validDateSerial(serial float64) bool { return serial >= 0 && serial < 2958466 }

var monthNames = []string{"january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"}

func monthByName(s string) int {
	s = strings.ToLower(strings.TrimSuffix(s, "."))
	if len(s) < 3 {
		return 0
	}
	for i, m := range monthNames {
		if strings.HasPrefix(m, s) {
			return i + 1
		}
	}
	return 0
}

// parseDateTime reads text as a date, a time, or both, the way Excel does
// in an English (US) locale: 1/2/2024 is 2 January, 2024-01-02, Jan 2,
// 2024, 2-Jan-2024, January 2024, 13:30, 1:30 PM, and combinations.
func parseDateTime(s string) (float64, bool) {
	t := strings.TrimSpace(s)
	if t == "" {
		return 0, false
	}
	// A time part: the last field holding a colon, with an optional AM/PM.
	fields := strings.Fields(t)
	var timeVal float64
	hasTime := false
	if n := len(fields); n > 0 {
		last := strings.ToUpper(fields[n-1])
		ampm := ""
		if last == "AM" || last == "PM" || last == "A" || last == "P" {
			ampm, fields = last, fields[:n-1]
		} else if strings.HasSuffix(last, "AM") || strings.HasSuffix(last, "PM") {
			ampm = last[len(last)-2:]
			fields[n-1] = fields[n-1][:len(fields[n-1])-2]
		}
		if k := len(fields); k > 0 && (strings.Contains(fields[k-1], ":") || ampm != "") {
			v, ok := parseTime(fields[k-1], ampm)
			if !ok {
				return 0, false
			}
			timeVal, hasTime, fields = v, true, fields[:k-1]
		}
	}
	if len(fields) == 0 {
		if hasTime {
			return timeVal, true
		}
		return 0, false
	}
	d, ok := parseDate(strings.Join(fields, " "))
	if !ok {
		return 0, false
	}
	return d + timeVal, true
}

func parseTime(s, ampm string) (float64, bool) {
	parts := strings.Split(s, ":")
	if len(parts) > 3 {
		return 0, false
	}
	var vals [3]float64
	for i, p := range parts {
		if p == "" {
			return 0, false
		}
		v, err := strconv.ParseFloat(p, 64)
		if err != nil || v < 0 {
			return 0, false
		}
		if i < 2 && strings.Contains(p, ".") {
			return 0, false
		}
		vals[i] = v
	}
	h, m, sec := vals[0], vals[1], vals[2]
	if len(parts) == 1 && ampm == "" {
		return 0, false
	}
	if m >= 60 || sec >= 60 {
		return 0, false
	}
	switch ampm {
	case "AM", "A":
		if h > 12 || h == 0 {
			return 0, false
		}
		if h == 12 {
			h = 0
		}
	case "PM", "P":
		if h > 12 || h == 0 {
			return 0, false
		}
		if h != 12 {
			h += 12
		}
	}
	if h >= 10000 {
		return 0, false
	}
	return (h*3600 + m*60 + sec) / 86400, true
}

func parseDate(s string) (float64, bool) {
	t := strings.TrimSpace(strings.ReplaceAll(s, ",", " "))
	t = splitMonthWords(t)
	// Numeric forms: m/d/y, m/d, y-m-d, m-d-y, d.m.y is not English.
	if strings.ContainsAny(t, "/-") && !strings.ContainsAny(t, " ") {
		sep := "/"
		if !strings.Contains(t, "/") {
			sep = "-"
		}
		parts := strings.Split(t, sep)
		allNum := true
		nums := make([]int, len(parts))
		for i, p := range parts {
			v, err := strconv.Atoi(p)
			if err != nil {
				allNum = false
				break
			}
			nums[i] = v
		}
		if allNum {
			switch len(parts) {
			case 3:
				if len(parts[0]) == 4 { // y-m-d
					return civilSerial(nums[0], nums[1], nums[2])
				}
				if nums[0] != nums[1] && nums[0] <= 12 && nums[1] <= 12 {
					// 11/05/1999 is 5 November or 11 May, by the language
					// Excel is set to: not to be guessed.
					return 0, false
				}
				return civilSerial(fullYear(nums[2], len(parts[2])), nums[0], nums[1])
			case 2:
				// m/d of this year, or m/yyyy.
				if len(parts[1]) == 4 {
					return civilSerial(nums[1], nums[0], 1)
				}
				return civilSerial(time.Now().Year(), nums[0], nums[1])
			}
			return 0, false
		}
		// 2-Jan-2024, Jan-2024, 2-Jan
		fields := strings.Split(t, sep)
		return wordDate(fields)
	}
	return wordDate(strings.Fields(strings.ReplaceAll(t, "-", " ")))
}

// wordDate reads dates with a month name: 2 Jan 2024, Jan 2 2024, January
// 2024, 2 Jan.
func wordDate(fields []string) (float64, bool) {
	if len(fields) == 1 {
		// A month alone (January): its first day this year.
		if m := monthByName(fields[0]); m > 0 {
			if _, err := strconv.Atoi(fields[0]); err != nil {
				return civilSerial(time.Now().Year(), m, 1)
			}
		}
		return 0, false
	}
	if len(fields) > 3 {
		return 0, false
	}
	mi := -1
	for i, f := range fields {
		if monthByName(f) > 0 {
			if _, err := strconv.Atoi(f); err != nil {
				mi = i
				break
			}
		}
	}
	if mi < 0 {
		return 0, false
	}
	m := monthByName(fields[mi])
	var rest []string
	for i, f := range fields {
		if i != mi {
			rest = append(rest, f)
		}
	}
	nums := make([]int, len(rest))
	for i, r := range rest {
		v, err := strconv.Atoi(strings.TrimSuffix(strings.TrimSuffix(strings.TrimSuffix(strings.TrimSuffix(r, "st"), "nd"), "rd"), "th"))
		if err != nil {
			return 0, false
		}
		nums[i] = v
	}
	switch len(nums) {
	case 1:
		if len(rest[0]) == 4 || nums[0] > 31 { // January 2024
			return civilSerial(fullYear(nums[0], len(rest[0])), m, 1)
		}
		return civilSerial(time.Now().Year(), m, nums[0])
	case 2:
		d, y := nums[0], nums[1]
		dl, yl := len(rest[0]), len(rest[1])
		if mi == 0 || len(rest[0]) == 4 { // Jan 2 2024; or 2024 Jan 2
			if len(rest[0]) == 4 {
				d, y, dl, yl = nums[1], nums[0], len(rest[1]), len(rest[0])
			}
		}
		_ = dl
		return civilSerial(fullYear(y, yl), m, d)
	}
	return 0, false
}

func fullYear(y, digits int) int {
	if digits <= 2 {
		if y < 30 {
			return 2000 + y
		}
		return 1900 + y
	}
	return y
}

func civilSerial(y, m, d int) (float64, bool) {
	if m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 9999 {
		return 0, false
	}
	if d > daysIn(y, m) && !(y == 1900 && m == 2 && d == 29) {
		return 0, false
	}
	if y == 1900 && m == 2 && d == 29 {
		return 60, true
	}
	return serialOf(y, m, d), true
}

func daysIn(y, m int) int {
	return time.Date(y, time.Month(m)+1, 0, 0, 0, 0, 0, time.UTC).Day()
}

// dateSerial is DATE(y, m, d): months and days roll over, years under 1900
// count from 1900.
func dateSerial(y, m, d float64) Value {
	yi, mi, di := int(math.Floor(y)), int(math.Floor(m)), int(math.Floor(d))
	if yi < 0 || yi >= 10000 {
		return errv(ErrNum)
	}
	if yi < 1900 {
		yi += 1900
	}
	// Months roll into years, days into months: serials count on evenly
	// from the first of a month (the phantom 29 February included).
	total := yi*12 + mi - 1
	yi = floorDiv(total, 12)
	mi = total - yi*12 + 1
	if yi < 1900 || yi > 9999 {
		return errv(ErrNum)
	}
	s := serialOf(yi, mi, 1) + float64(di-1)
	if s < 0 || s >= 2958466 {
		return errv(ErrNum)
	}
	return num(s)
}

// splitMonthWords puts spaces between digits and letters, so that dates
// written together (1Jan2024, Jan2024) read like 1 Jan 2024.
func splitMonthWords(t string) string {
	var b strings.Builder
	prev := rune(0)
	for _, r := range t {
		if prev != 0 && (isDigitRune(prev) && isLetterRune(r) || isLetterRune(prev) && isDigitRune(r)) {
			b.WriteByte(' ')
		}
		b.WriteRune(r)
		prev = r
	}
	return b.String()
}

func isDigitRune(r rune) bool  { return r >= '0' && r <= '9' }
func isLetterRune(r rune) bool { return r >= 'a' && r <= 'z' || r >= 'A' && r <= 'Z' }

// dayFirstDate tells whether text is a date only when read day first:
// 21/11/2024, or 21/11/24 14:28.
func dayFirstDate(s string) bool {
	fields := strings.Fields(strings.TrimSpace(s))
	if len(fields) == 0 || len(fields) > 3 {
		return false
	}
	parts := strings.FieldsFunc(fields[0], func(r rune) bool { return r == '/' || r == '-' || r == '.' })
	if len(parts) < 2 || len(parts) > 3 {
		return false
	}
	nums := make([]int, len(parts))
	for i, p := range parts {
		n, err := strconv.Atoi(p)
		if err != nil {
			return false
		}
		nums[i] = n
	}
	if len(nums) == 3 && len(parts[0]) == 4 {
		return false // year first
	}
	// Day first only (21/11/2024), or either way (11/05/2024).
	return nums[0] >= 1 && nums[0] <= 31 && nums[1] >= 1 && nums[1] <= 12 && (nums[0] > 12 || nums[0] != nums[1] && len(nums) == 3)
}
