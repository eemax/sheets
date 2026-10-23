package grid

import (
	"math"
	"strconv"
	"strings"
	"time"
)

// SerialToTime converts a spreadsheet serial number to a time.
func SerialToTime(serial float64, date1904 bool) time.Time {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	if date1904 {
		epoch = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	} else if serial < 61 {
		// Serials before 1900-03-01 predate Excel's fictitious 1900-02-29.
		epoch = epoch.AddDate(0, 0, 1)
	}
	days := math.Floor(serial)
	ms := math.Round((serial - days) * 86400000)
	return epoch.AddDate(0, 0, int(days)).Add(time.Duration(ms) * time.Millisecond)
}

// TimeToSerial converts a time to a spreadsheet serial number (1900 system).
func TimeToSerial(t time.Time) float64 {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t = time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), time.UTC)
	d := t.Sub(epoch)
	serial := float64(d/(24*time.Hour)) + float64(d%(24*time.Hour))/float64(24*time.Hour)
	if serial < 61 {
		serial--
	}
	return serial
}

// FormatNumber renders a float the way a spreadsheet would show it in General
// format: at most 15 significant digits, no exponent for ordinary magnitudes.
func FormatNumber(v float64) string {
	if v == 0 {
		return "0"
	}
	if math.IsNaN(v) || math.IsInf(v, 0) {
		return strconv.FormatFloat(v, 'g', -1, 64)
	}
	r, _ := strconv.ParseFloat(strconv.FormatFloat(v, 'g', 15, 64), 64)
	if a := math.Abs(r); a >= 1e21 || a < 1e-9 {
		return strconv.FormatFloat(r, 'g', -1, 64)
	}
	return strconv.FormatFloat(r, 'f', -1, 64)
}

// Display renders a cell value as plain, unambiguous text for agents: ISO
// dates, percentages with a % sign, other numbers without separators.
func Display(c *Cell, f Format, date1904 bool) string {
	switch c.Kind {
	case Text, Error:
		return c.Text
	case Bool:
		if c.Bool {
			return "TRUE"
		}
		return "FALSE"
	case Number:
		switch f.Kind {
		case Date, DateTime, Time:
			return FormatTemporal(c.Num, f.Kind, date1904)
		case Percent:
			return FormatNumber(c.Num*100) + "%"
		}
		return FormatNumber(c.Num)
	}
	return ""
}

// FormatTemporal renders a date/time serial in ISO style.
func FormatTemporal(serial float64, k FormatKind, date1904 bool) string {
	t := SerialToTime(serial, date1904)
	switch k {
	case Date:
		return t.Format("2006-01-02")
	case Time:
		if t.Second() != 0 {
			return t.Format("15:04:05")
		}
		return t.Format("15:04")
	}
	if t.Second() != 0 {
		return t.Format("2006-01-02 15:04:05")
	}
	if t.Hour() == 0 && t.Minute() == 0 {
		return t.Format("2006-01-02")
	}
	return t.Format("2006-01-02 15:04")
}

// Truncate shortens s to at most n runes, marking the cut with an ellipsis.
func Truncate(s string, n int) string {
	if n <= 0 || len(s) <= n {
		return s
	}
	r := []rune(s)
	if len(r) <= n {
		return s
	}
	return string(r[:n-1]) + "…"
}

// Escape makes a value safe for tab-separated output.
func Escape(s string) string {
	if !strings.ContainsAny(s, "\t\n\r") {
		return s
	}
	return strings.NewReplacer("\t", `\t`, "\n", `\n`, "\r", `\r`).Replace(s)
}
