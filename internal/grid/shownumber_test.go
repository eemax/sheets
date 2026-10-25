package grid

import "testing"

func TestShowNumber(t *testing.T) {
	for _, c := range []struct {
		in   float64
		want string
	}{
		{24836587.209999993, "24836587.21"},
		{-24836587.209999993, "-24836587.21"},
		{1234567.1200000001, "1234567.12"},
		{12345678.300000004, "12345678.3"},
		{0.1 + 0.2, "0.3"},
		{1.0 / 3, "0.333333333333333"},
		{1.0000001, "1.0000001"},
		{123456789012345, "123456789012345"},
		{2.5, "2.5"},
		{99999.999999999, "99999.999999999"}, // as written: 14 digits
	} {
		if got := ShowNumber(c.in); got != c.want {
			t.Errorf("ShowNumber(%v) = %s (FormatNumber %s), want %s", c.in, got, FormatNumber(c.in), c.want)
		}
	}
}
