package formula

import (
	"reflect"
	"testing"
)

func TestDepsOf(t *testing.T) {
	for _, tt := range []struct {
		f                  string
		refs               []string // as Sheet!Range, own sheet as !Range
		names, tables      []string
		volatile, position bool
		external           bool
	}{
		{f: "SUM(B2:B9)*Rate", refs: []string{"!B2:B9"}, names: []string{"Rate"}},
		{f: "Data!$A$1+'My Sheet'!C3:D4", refs: []string{"Data!A1", "My Sheet!C3:D4"}},
		{f: "SUM(A:A)+SUM(3:4)", refs: []string{"!A1:A1048576", "!A3:XFD4"}},
		{f: "SUM(Sales[Units])+[@Price]", tables: []string{"Sales", ""}},
		{f: `IF(A1>0,"Rate",NOW())`, refs: []string{"!A1"}, volatile: true},
		{f: "INDIRECT(\"B\"&ROW())", volatile: true, position: true},
		{f: "OFFSET(A1,1,1)", refs: []string{"!A1"}, volatile: true},
		{f: "LET(x,B2,y,C2,x*y+Tax)", refs: []string{"!B2", "!C2"}, names: []string{"Tax"}},
		{f: "_xlfn.LET(_xlpm.x,B2,_xlpm.x*2)", refs: []string{"!B2"}},
		{f: "Double(B2)", refs: []string{"!B2"}, names: []string{"Double"}},
		{f: "[1]Rates!B2*2", refs: []string{"Rates!B2"}, external: true},
		{f: "'[Book.xlsx]Rates'!B2+A1", refs: []string{"[Book.xlsx]Rates!B2", "!A1"}, external: true},
		{f: "SUM(Jan:Dec!B2)", refs: []string{"!JAN1:JAN1048576", "Dec!B2"}, volatile: true}, // volatile covers it
		{f: "IF(A1,#N/A,#DIV/0!)&#NAME?", refs: []string{"!A1"}},
		{f: "TRUE+1.5E+3+FALSE", refs: nil},
		{f: "IF(ISERROR(A1),#N/A,#REF!)", refs: []string{"!A1"}},
		{f: "COLUMNS(A1:C1)", refs: []string{"!A1:C1"}, position: true},
		{f: "Sheet2!Total*2", names: []string{"Total"}},
	} {
		d := DepsOf(tt.f)
		var refs []string
		for _, r := range d.Refs {
			refs = append(refs, r.Sheet+"!"+r.Range.String())
		}
		if !reflect.DeepEqual(refs, tt.refs) || !reflect.DeepEqual(d.Names, tt.names) || !reflect.DeepEqual(d.Tables, tt.tables) ||
			d.Volatile != tt.volatile || d.Position != tt.position || d.External != tt.external {
			t.Errorf("DepsOf(%q) = refs %v names %v tables %q volatile %v position %v external %v", tt.f, refs, d.Names, d.Tables, d.Volatile, d.Position, d.External)
		}
	}
}

func TestErrorLiterals(t *testing.T) {
	for f, want := range map[string][]string{
		"SUM(#REF!)+#N/A":         {"#REF!", "#N/A"},
		`IF(A1="#N/A",1,#DIV/0!)`: {"#DIV/0!"},
		"Data!#REF!*2":            {"#REF!"},
		"SUM(A1:A3)":              nil,
	} {
		if got := ErrorLiterals(f); !reflect.DeepEqual(got, want) {
			t.Errorf("ErrorLiterals(%q) = %q, want %q", f, got, want)
		}
	}
}

func TestArrayUse(t *testing.T) {
	multi := func(n string) bool { return n == "Units" }
	for f, array := range map[string]bool{
		"SUM(A1:A9)":                            false,
		"SUMPRODUCT(B2:B4,C2:C4)":               false,
		"VLOOKUP(\"x\",Data!A2:C9,2,FALSE)":     false,
		"SUMIFS(C:C,A:A,\"East\")":              false,
		"A1*2+B1":                               false,
		"SUM((A1:A9))":                          false,
		"SUMPRODUCT((A2:A4=\"East\")*B2:B4)":    true,
		"SUM(A1:A9*2)":                          true,
		"SUMPRODUCT(--(A2:A4=\"East\"),B2:B4)":  true,
		"A1:A9":                                 true,
		"ABS(A1:A9)":                            true,
		"IF(A1:A3>0,1,0)":                       true,
		"SUM(Units)":                            false,
		"SUM(Units*2)":                          true,
		"ROUND(SUM(Data!B2:B9)/COUNT(B2:B9),2)": false,
		"INDEX(A1:C9,MATCH(\"x\",A1:A9,0),2)":   false,
		"(A1:A9)*2":                             true,
		"TEXT(A1,\"0\")&\" of \"&ROWS(A1:A9)":   false,
		"UNIQUE(A2:A4)":                         false,
		"FILTER(A2:A9,B2:B9)":                   false,
		"SORTBY(A2:A9,B2:B9,-1)":                false,
		"SUMIFS(C:C,A:A,\"East\",B:B,\">2\")":   false,
		"COUNTIFS(A:A,\"x\",B:B,1)":             false,
		"SUMPRODUCT({1,2,3},A1:A3)":             false,
		"NETWORKDAYS(A1,B1,H1:H9)":              false,
		"VLOOKUP(A2:A4,D:E,2,FALSE)":            true,
		"COUNTIF(A:A,B2:B4)":                    true,
		"SUMIF(A:A,B2:B4,C:C)":                  true,
		"INDEX(A:A,B2:B4)":                      true,
		"NETWORKDAYS(A1:A3,B1)":                 true,
	} {
		if got := ArrayUse(f, multi) != ""; got != array {
			t.Errorf("ArrayUse(%q) = %v (%s), want %v", f, got, ArrayUse(f, multi), array)
		}
	}
}
