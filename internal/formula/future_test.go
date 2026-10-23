package formula

import "testing"

func TestToFile(t *testing.T) {
	tests := map[string]string{
		`XLOOKUP("b",A1:A2,B1:B2)`:    `_xlfn.XLOOKUP("b",A1:A2,B1:B2)`,
		`xlookup(A1,B:B,C:C)`:         `_xlfn.XLOOKUP(A1,B:B,C:C)`,
		`SUM(FILTER(A1:A9,B1:B9>0))`:  `SUM(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))`,
		`STDEV.S(A1:A9)+STDEV(A1:A9)`: `_xlfn.STDEV.S(A1:A9)+STDEV(A1:A9)`,
		`_xlfn.IFS(A1>0,1,TRUE,0)`:    `_xlfn.IFS(A1>0,1,TRUE,0)`,
		`"XLOOKUP(" & A1`:             `"XLOOKUP(" & A1`,
		`LET(x, A1*2, y, x+1, x*y)`:   `_xlfn.LET(_xlpm.x, A1*2, _xlpm.y, _xlpm.x+1, _xlpm.x*_xlpm.y)`,
		`LET(f,LAMBDA(n,n+1),f(2))`:   `_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(2))`,
		`MAP(A1:A3,LAMBDA(v,v*2))`:    `_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))`,
		`'IFS sheet'!A1+Sheet2!B2`:    `'IFS sheet'!A1+Sheet2!B2`,
		`SUM(A1:A3)`:                  `SUM(A1:A3)`,
	}
	for in, want := range tests {
		if got := ToFile(in); got != want {
			t.Errorf("ToFile(%s)\n got %s\nwant %s", in, got, want)
		}
		if back := FromFile(want); ToFile(back) != want {
			t.Errorf("round trip of %s gave %s", want, back)
		}
	}
	if got := FromFile(`_xlfn.LET(_xlpm.x,_xlfn._xlws.SORT(A1:A3),_xlpm.x)`); got != `LET(x,SORT(A1:A3),x)` {
		t.Errorf("FromFile = %s", got)
	}
}

func TestSpills(t *testing.T) {
	for f, want := range map[string]bool{
		`UNIQUE(A1:A9)`:           true,
		`_xlfn._xlws.SORT(A1:A9)`: true,
		`SUM(A1:A9)`:              false,
		`XLOOKUP(A1,B:B,C:C)`:     false,
		`"UNIQUE("&A1`:            false,
		`COUNTA(UNIQUE(A1:A9))`:   true, // a scalar, but stored dynamic like Excel 365 does
	} {
		if got := Spills(f); got != want {
			t.Errorf("Spills(%s) = %v", f, got)
		}
	}
}

func TestSheetRefs(t *testing.T) {
	if got := RenameSheet(`Data!A1*2+'Data'!B2+SUM(data!C1:C3)+Other!A1+[1]Data!A1`, "Data", "Sales 2026"); got != `'Sales 2026'!A1*2+'Sales 2026'!B2+SUM('Sales 2026'!C1:C3)+Other!A1+[1]Data!A1` {
		t.Errorf("RenameSheet = %s", got)
	}
	if got := RenameSheet(`'Bob''s'!A1`, "Bob's", "Plain"); got != `Plain!A1` {
		t.Errorf("RenameSheet quoted = %s", got)
	}
	if got := RenameSheet(`"Data!A1"&A1`, "Data", "X"); got != `"Data!A1"&A1` {
		t.Errorf("RenameSheet in string = %s", got)
	}
	if !RefersToSheet(`SUM('My Data'!A:A)`, "my data") || RefersToSheet(`A1+B1`, "Data") {
		t.Error("RefersToSheet")
	}
	if got := DropSheet(`Data!A1*2+SUM(Data!B1:B9)+Keep!A1`, "Data"); got != `#REF!*2+SUM(#REF!)+Keep!A1` {
		t.Errorf("DropSheet = %s", got)
	}
}
