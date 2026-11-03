#!/usr/bin/env bash
# Builds workbooks to open in real Excel, the one check the test suite
# cannot do. See excel-check/CHECKLIST.md after running:
#
#   scripts/excel-check.sh [output-dir]
#
# Needs Go, and python3 (its standard library only) for 22.
set -euo pipefail

root=$(cd "$(dirname "$0")/.." && pwd)
out=${1:-"$root/excel-check"}
bin=$(mktemp -d)/sheets
trap 'rm -rf "$(dirname "$bin")"' EXIT
(cd "$root" && go build -o "$bin" ./cmd/sheets)
export SHEETS_UNDO=off
rm -rf "$out" && mkdir -p "$out"
cd "$out"
s() { "$bin" "$@" >/dev/null; }

# 1. The house style on a typical report.
printf '%s\n' \
  'Q3 sales by region' \
  $'Date\tRegion\tUnits\tPrice\tRevenue\tMargin\tNotes' \
  $'2026-07-01\tNorth\t120\t$12.50\t=C{r}*D{r}\t18%\tRenewal' \
  $'2026-07-02\tSouth\t80\t$30.00\t=C{r}*D{r}\t22.5%\tNew customer, signed after a long trial that included two rounds of procurement review' \
  $'2026-07-03\tEast\t2000\t$12.50\t=C{r}*D{r}\t17%\t' \
  $'2026-07-04\tWest\t15\t$120.00\t=C{r}*D{r}\t-3%\tDiscounted' \
  $'2026-07-05\tNorth\t95\t$30.00\t=C{r}*D{r}\t21%\t' \
  $'2026-07-06\tSouth\t310\t$12.50\t=C{r}*D{r}\t19%\t' \
  $'Total\t\t=SUM(C3:C8)\t\t=SUM(E3:E8)\t\t' |
  "$bin" write 01-house-style.xlsx Sales >/dev/null
s fmt 01-house-style.xlsx

# 2. The finance preset, with a link to another sheet.
cp 01-house-style.xlsx 02-finance.xlsx
s set 02-finance.xlsx 'Assumptions!A1' 'Tax rate' 'Assumptions!B1' '21%'
s set 02-finance.xlsx 'Sales!H2' 'After tax' 'Sales!H3:H8' '=E{r}*(1-Assumptions!$B$1)' 'Sales!H9' '=SUM(H3:H8)'
s fmt 02-finance.xlsx -p finance

# 3. Modern functions and dynamic arrays.
printf '%s\n' $'Name\tScore' $'Ada\t9' $'Bob\t4' $'Cy\t7' $'Ada\t8' | "$bin" write 03-modern-formulas.xlsx >/dev/null
s set 03-modern-formulas.xlsx \
  D1 'XLOOKUP Bob' E1 '=XLOOKUP("Bob",A2:A5,B2:B5)' \
  D2 'IFS' E2 '=IFS(B2>8,"high",B2>5,"mid",TRUE,"low")' \
  D3 'TEXTJOIN' E3 '=TEXTJOIN(", ",TRUE,A2:A5)' \
  D4 'LET' E4 '=LET(x,B2,y,B3,x*y)' \
  D5 'LAMBDA' E5 '=LET(twice,LAMBDA(n,n*2),twice(B4))' \
  D6 'MAXIFS' E6 '=MAXIFS(B2:B5,A2:A5,"Ada")' \
  G1 'UNIQUE (spills down)' G2 '=UNIQUE(A2:A5)' \
  H1 'SORT(FILTER) (spills)' H2 '=SORT(FILTER(B2:B5,B2:B5>5))' \
  I1 'SEQUENCE (spills)' I2 '=SEQUENCE(3)'

# 4. Structural edits on a workbook made by another program (openpyxl),
#    with a chart, a validation, a comment, a link and an Excel table.
cp "$root/internal/xlsx/testdata/rich.xlsx" 04-structure.xlsx
s set 04-structure.xlsx 'Summary!A1' 'North units' 'Summary!B1' '=Data!B2' 'Summary!A2' 'All units' 'Summary!B2' '=SUM(Data!B2:B4)'
s rows insert 04-structure.xlsx 'Data!3' -n 2
s cols insert 04-structure.xlsx 'Data!B'
s sheet rename 04-structure.xlsx Data 'Sales Data'
s sheet copy 04-structure.xlsx Summary 'Summary copy'
s rows delete 04-structure.xlsx "Summary copy!1"

# 5. Restyling a workbook formatted by another tool.
cp "$root/internal/xlsx/testdata/styled.xlsx" 05-restyled.xlsx
s fmt 05-restyled.xlsx

# 6. A workbook with an Excel table, formatted.
cp "$root/internal/xlsx/testdata/rich.xlsx" 06-excel-table.xlsx
s fmt 06-excel-table.xlsx

# 7. Charts of each kind, in the house style.
printf '%s\n' $'Region\tQ1\tQ2\tQ3' $'North America\t120\t135\t150' $'South America\t80\t95\t70' \
  $'Europe\t140\t160\t175' $'Asia Pacific\t200\t190\t230' | "$bin" write 07-charts.xlsx Regions >/dev/null
{
  printf 'Month\tVisits\tSignups\n'
  for i in $(seq 1 12); do printf '2026-%02d-01\t%d\t%d\n' "$i" $((900 + i * 70 + (i % 3) * 120)) $((40 + i * 6 + (i % 4) * 9)); done
} | "$bin" write 07-charts.xlsx Traffic >/dev/null
printf '%s\n' $'Channel\tShare' $'Search\t42%' $'Social\t23%' $'Email\t18%' $'Direct\t12%' $'Other\t5%' |
  "$bin" write 07-charts.xlsx Mix >/dev/null
printf '%s\n' $'Spend\tClicks' $'$100\t40' $'$250\t95' $'$400\t170' $'$650\t240' $'$800\t310' $'$1,000\t360' |
  "$bin" write 07-charts.xlsx Ads >/dev/null
s fmt 07-charts.xlsx
s chart add 07-charts.xlsx Regions --at Regions!F2
s chart add 07-charts.xlsx Regions --type bar --stacked --at Regions!F22
s chart add 07-charts.xlsx Traffic --y Visits --at Traffic!E2
s chart add 07-charts.xlsx Traffic --y Signups --type area --at Traffic!E22
s chart add 07-charts.xlsx Mix --type pie --at Mix!D2
s chart add 07-charts.xlsx Ads --x Spend --y Clicks --type scatter --at Ads!D2

# 8. A two-row header, formatted.
printf '%s\n' \
  'Sales by quarter' \
  $'Region\tQ1\t\tQ2\t' \
  $'\tUnits\tRevenue\tUnits\tRevenue' \
  $'North\t5\t$100\t6\t$120' \
  $'South\t7\t$140\t8\t$160' \
  $'Total\t=SUM(B4:B5)\t=SUM(C4:C5)\t=SUM(D4:D5)\t=SUM(E4:E5)' |
  "$bin" write 08-two-row-header.xlsx >/dev/null
s fmt 08-two-row-header.xlsx

# 9. An Excel table with a totals row (made by openpyxl) through edits:
#    rows appended above its totals row, headers renamed (one that
#    formulas name), rows inserted above it.
cp "$root/internal/xlsx/testdata/tables.xlsx" 09-excel-table-edits.xlsx
printf '%s\n' $'West\t4\t1.5' $'Central\t6\t2' | "$bin" write 09-excel-table-edits.xlsx Sales --append >/dev/null
s set 09-excel-table-edits.xlsx Sales!C1 'Unit price' Sales!B1 'Qty'
s rows insert 09-excel-table-edits.xlsx Sales!1 -n 2
s set 09-excel-table-edits.xlsx Sales!A1 'Inserted above the table'

# 10. A query result written back as a table.
cp 01-house-style.xlsx 10-query-into.xlsx
s query 10-query-into.xlsx 'SELECT Region, SUM(Units) AS Units, SUM(Revenue) AS Revenue FROM Sales GROUP BY Region ORDER BY Revenue DESC' --into Summary
s fmt 10-query-into.xlsx

# 11. Targeted styling (sheets style and view), kept through fmt.
"$bin" do 11-styling.xlsx >/dev/null <<'EOF'
write Report!A3 <<END
Item	Budget	Spent	Remaining	Owner	Notes
Rent	1200000	1100000	=B{r}-C{r}	Facilities	Annual lease, renews in March
Staff	800000	850000	=B{r}-C{r}	HR	Two hires delayed
IT	150000	90000.5	=B{r}-C{r}	IT
Travel	60000	72000	=B{r}-C{r}	Finance	Conference season
Total	=SUM(B4:B7)	=SUM(C4:C7)	=SUM(D4:D7)
END
set Report!A1 'Budget review 2026'
fmt
style Report!A1:F1 merge align=center size=16 bold
style 'Report[#Headers]' fill=lightblue
style 'Report[[Budget]:[Remaining]]' format=currency decimals=0
style 'Report[#Totals]' border=top:thin border=bottom:double
style Report!E3 rotate=45
style 'Report[Owner]' italic color=darkgray indent=1
style Report!F:F width=24 wrap
view Report --freeze B4 --tab teal --zoom 120
EOF
"$bin" do 11-styling.xlsx >/dev/null <<'EOF'
write Report --append <<END
Item	Budget	Spent	Remaining	Owner
Training	20000	5000	=B{r}-C{r}	HR
END
fmt
EOF

# 12. Conditional formats of every kind, over fmt's banding.
"$bin" do 12-conditional.xlsx >/dev/null <<'EOF'
write Scores!A1 <<END
Name	Score	Change	Status	Region	Code
Ada	92	0.12	On time	East	A-1
Bob	47	-0.08	Late	West	B-2
Cy	75	0.03	On time	East	A-1
Dee	88	-0.15	Late	North	C-9
Eve	61	0.2	Done	South	D-4
Fay	39	0	Late	East	E-5
Gus	99	0.07	On time	West	F-6
Hal		0.01	Done	North	G-7
END
fmt
cond Scores!A2:F9 '=$D2="Late"' fill=lightyellow
cond 'Scores[Score]' 'top 25%' fill=lightgreen bold
cond 'Scores[Score]' '< 50' color=red bold
cond 'Scores[Score]' blank fill=lightgray
cond 'Scores[Change]' '< 0' color=red 'format=0.0%;(0.0%)'
cond 'Scores[Change]' bars
cond 'Scores[Region]' 'contains st' italic
cond 'Scores[Code]' duplicates color=purple
EOF
"$bin" do 12-conditional.xlsx >/dev/null <<'EOF'
write Scales!A1 <<END
Month	Visits	Growth	Rating
Jan	120	0.05	1
Feb	340	0.12	2
Mar	90	-0.04	3
Apr	410	0.2	4
May	260	0.02	5
END
fmt
cond 'Scales[Visits]' scale
cond 'Scales[Growth]' scale red-white-green
cond 'Scales[Rating]' icons arrows
cond 'Scales[Visits]' 'above average' bold
EOF

# 13. Data validation of every kind.
"$bin" do 13-validation.xlsx >/dev/null <<'EOF'
write Orders!A1 <<END
Region	Product	Qty	Price	Date	Code	Note
East	Pen	2	1.5	2026-03-01	A100
END
write Lists!A1 <<END
Products
Pen
Ink
Paper
Stapler
END
fmt
validate Orders!A2:A50 list East,West,North,South --prompt 'Pick a region'
validate Orders!B2:B50 list =Lists!$A$2:$A$5
validate Orders!C2:C50 whole between 1 and 1000 --error 'Qty is 1 to 1000'
validate Orders!D2:D50 number '>' 0 --warn --error 'A price is usually above 0'
validate Orders!E2:E50 date between 2026-01-01 and 2026-12-31
validate Orders!F2:F50 formula '=COUNTIF($F$2:$F$50,F2)=1' --error 'Codes are unique'
validate Orders!G2:G50 length '<=' 20 --prompt 'Up to 20 characters'
EOF

# 14. 12 after inserting rows above and a column before its data: rules,
#     ranges and frozen panes follow, as Excel moves them.
cp 12-conditional.xlsx 14-moved.xlsx
s view 14-moved.xlsx Scores --freeze B2
s rows insert 14-moved.xlsx Scores!1 -n 2
s cols insert 14-moved.xlsx Scores!A
s set 14-moved.xlsx Scores!A1 'Two rows and a column were inserted'

# 15. Sorting: a report with a running total, a note and a link; and an
#     Excel table (made by openpyxl) sorted within its data rows.
"$bin" do 15-sorted.xlsx >/dev/null <<'EOF'
write Sales!A1 <<END
Region	Units	Price	Revenue	Running
East	5	30	=B{r}*C{r}	=SUM($D$2:D{r})
West	2	40	=B{r}*C{r}	=SUM($D$2:D{r})
North	9	10	=B{r}*C{r}	=SUM($D$2:D{r})
South	1	100	=B{r}*C{r}	=SUM($D$2:D{r})
Total	=SUM(B2:B5)		=SUM(D2:D5)
END
fmt
note Sales!A3 'West: two only'
link Sales!A4 https://example.com/north
sort Sales Revenue:desc
EOF
cp "$root/internal/xlsx/testdata/tables.xlsx" 15b-sorted-table.xlsx
s sort 15b-sorted-table.xlsx Sales Units:desc

# 16. Copy and move: a styled block with a highlight rule, a dropdown, a
#     note and a link moved down, then copied to another sheet; formulas
#     elsewhere follow the move, and one pointing where it lands breaks.
"$bin" do 16-copy-move.xlsx >/dev/null <<'EOF'
write Data!A1 <<END
Name	Score	Status
Ada	92	Late
Bob	47	On time
Cy	75	Late
END
fmt
style 'Data[#Headers]' fill=lightblue
cond Data!A2:C4 '=$C2="Late"' fill=lightyellow
validate Data!C2:C4 list 'Late,On time'
note Data!A2 'Team lead'
link Data!A3 https://example.com/bob
set Summary!A1 'Best score' Summary!B1 '=MAX(Data!B2:B4)' Summary!A2 'Ada' Summary!B2 '=Data!B2' Summary!A3 'Breaks' Summary!B3 '=Data!B10'
EOF
s move 16-copy-move.xlsx Data!A1:C4 Data!A8
s copy 16-copy-move.xlsx Data!A8:C11 Copy!A2
s fmt 16-copy-move.xlsx

# 17. Notes moving with rows in files from other programs: openpyxl's
#     notes, and an Excel 365 comment thread.
cp "$root/internal/xlsx/testdata/rich.xlsx" 17-notes.xlsx
s rows insert 17-notes.xlsx Data!1 -n 2
s note 17-notes.xlsx Data!A1 'Added by sheets'
cp "$root/internal/xlsx/testdata/threads.xlsx" 17b-thread.xlsx
s rows insert 17b-thread.xlsx Data!2
s sort 17b-thread.xlsx Data!A3:B5 B
s note 17b-thread.xlsx Data!C3 'A note next to the thread'

# 18. A CSV with typed values, converted and formatted.
printf '%s\r\n' 'Order,Date,Customer,Amount,Discount,Code,Comment' \
  '1001,2026-03-01,"Acme, Inc.","1,250.00",5%,00042,"said ""rush"""' \
  '1002,2026-03-02,Zoë Café,$99.50,0%,00107,' \
  '1003,2026-03-03,Bob,"12,000",12.5%,01234,"two' 'lines"' > 18-orders.csv
s convert 18-orders.csv 18-from-csv.xlsx
s fmt 18-from-csv.xlsx

# 19. Conditional formats and validations from another program (data bars
#     with Excel 2010's extension, icons, a highlight reading its own row,
#     rules reading another sheet) through row and column edits, one of
#     them streamed.
cp "$root/internal/xlsx/testdata/rules.xlsx" 19-rules.xlsx
s rows delete 19-rules.xlsx Data!3:3
s cols insert 19-rules.xlsx Data!A
s cond 19-rules.xlsx Data!D3:D6 bars
s rows insert 19-rules.xlsx Data!5
s set 19-rules.xlsx Data!B5 Added Data!C5 4 Data!D5 50 Data!E5 1

# 20. Dynamic arrays, saved as Excel 365 saves them: the formula with
#     cm="1" metadata, the cells it spills into holding their values; one
#     blocked (#SPILL!); and on Edits two that grew and shrank after an
#     edit. 20b: the same edits through the streaming editor.
printf '%s\n' $'Name\tScore' $'Ada\t9' $'Bob\t4' $'Cy\t7' $'Ada\t8' $'Dee\t6' | "$bin" write 20-dynamic-arrays.xlsx Data >/dev/null
s set 20-dynamic-arrays.xlsx \
  Data!D1 'Unique' Data!D2 '=UNIQUE(A2:A6)' \
  Data!E1 'Top scores' Data!E2 '=SORT(FILTER(B2:B6,B2:B6>5),,-1)' \
  Data!F1 'Grid' Data!F2 '=SEQUENCE(2,3)' \
  Data!J1 'Blocked' Data!J3 'in the way' Data!J2 '=UNIQUE(A2:A6)'
printf '%s\n' $'Item\tCode' $'a\tx' $'b\ty' $'a\tz' | "$bin" write 20-dynamic-arrays.xlsx Edits >/dev/null
s set 20-dynamic-arrays.xlsx Edits!D1 'Grows' Edits!D2 '=UNIQUE(A2:A4)' Edits!E1 'Shrinks' Edits!E2 '=UNIQUE(B2:B4)'
cp 20-dynamic-arrays.xlsx 20b-dynamic-arrays-streamed.xlsx
s set 20-dynamic-arrays.xlsx Edits!A4 c
s set 20-dynamic-arrays.xlsx Edits!B3 x Edits!B4 x
SHEETS_STREAM=always "$bin" set 20b-dynamic-arrays-streamed.xlsx Edits!A4 c >/dev/null
SHEETS_STREAM=always "$bin" set 20b-dynamic-arrays-streamed.xlsx Edits!B3 x Edits!B4 x >/dev/null

# 20c. A dynamic array on a sheet the same command adds to a workbook.
printf '%s\n' Name d b a c | "$bin" write 20c-spill-on-new-sheet.xlsx Data >/dev/null
s set 20c-spill-on-new-sheet.xlsx New!A1 'Sorted' New!B3 '=SORT(Data!A2:A5)' New!B8 'below'

# 21. The spill operator: C2# (stored as _xlfn.ANCHORARRAY(C2)) read by
#     other formulas, one of them spilling itself.
printf '%s\n' Colour Red Blue Red Green Blue | "$bin" write 21-spill-refs.xlsx Colours >/dev/null
s set 21-spill-refs.xlsx \
  C1 'Sorted unique' C2 '=SORT(UNIQUE(A2:A6))' \
  D1 'How many' D2 '=COUNTA(C2#)' \
  E1 'Shout' E2 '=C2#&"!"' \
  F1 'Times seen' F2 '=COUNTIF(A2:A6,C2#)' \
  G1 'Second' G2 '=INDEX(C2#,2)'

# 22. Sorting rows holding filled-down formulas stored as Excel stores
#     them, shared (one formula, ref="D2:D6", the rows below pointing at
#     it): sheets writes formulas one per cell, so they are shared here as
#     Excel would. The sort unshares every formula on the sheet, and the
#     results must keep their types: numbers, text, TRUE/FALSE. A second
#     block below the sorted range is unshared without being sorted.
#     22b: the same sort through the streaming editor.
share() { # share <file> <ref>...: share the formulas of each column range on the first sheet
  python3 -I - "$@" <<'PY'
import re, sys, zipfile
path, refs = sys.argv[1], sys.argv[2:]
part = "xl/worksheets/sheet1.xml"
with zipfile.ZipFile(path) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
def share(xml):
    for si, ref in enumerate(refs):
        (col, r1), (_, r2) = (re.fullmatch(r"([A-Z]+)(\d+)", a).groups() for a in ref.split(":"))
        for r in range(int(r1), int(r2) + 1):
            m = re.search(r'(<c r="%s%d"[^>]*>)<f>([^<]*)</f>' % (col, r), xml)
            if not m:
                sys.exit("share: no formula in %s%d" % (col, r))
            f = '<f t="shared" ref="%s" si="%d">%s</f>' % (ref, si, m[2]) if r == int(r1) else '<f t="shared" si="%d"/>' % si
            xml = xml[: m.start()] + m[1] + f + xml[m.end() :]
    return xml
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in items:
        z.writestr(info, share(data.decode()).encode() if info.filename == part else data)
PY
}
f22=$'\t=B{r}*C{r}\t=A{r}&" #"&B{r}\t=B{r}>4\t=D{r}/SUM($D$2:$D$6)'
printf '%s\n' $'Region\tUnits\tPrice\tRevenue\tLabel\tBig\tShare' \
  $'East\t5\t30'"$f22" $'West\t2\t40'"$f22" $'North\t9\t10'"$f22" \
  $'South\t1\t100'"$f22" $'Mid\t4\t2.5'"$f22" |
  "$bin" write 22-shared-sort.xlsx Sales >/dev/null
s set 22-shared-sort.xlsx A7 Total B7 '=SUM(B2:B6)' D7 '=SUM(D2:D6)' \
  A9 Item B9 Cost C9 Doubled D9 Tag E9 'Over 50' \
  A10 Rent B10 40 A11 Power B11 55 A12 Ink B12 7 \
  C10:C12 '=B{r}*2' D10:D12 '=A{r}&"!"' E10:E12 '=B{r}>50'
s style 22-shared-sort.xlsx Sales!G2:G6 format=percent decimals=1
share 22-shared-sort.xlsx D2:D6 E2:E6 F2:F6 G2:G6 C10:C12 D10:D12 E10:E12
cp 22-shared-sort.xlsx 22b-shared-sort-streamed.xlsx
s sort 22-shared-sort.xlsx Sales!A1:G6 Revenue:desc
SHEETS_STREAM=always "$bin" sort 22b-shared-sort-streamed.xlsx Sales!A1:G6 Revenue:desc >/dev/null

# 23. A workbook changed by sheets script (Office Scripts): duplicates
#     removed, formulas filled, a total row, number formats, fonts, fills,
#     borders, a sort, frozen panes, a filter; rows added to an Excel table
#     (made by openpyxl) above its total row; a spilling formula; a new
#     sheet with formulas and a date. The script is kept beside the
#     workbook. (A spilling formula on the new sheet is 20c.)
cp "$root/internal/xlsx/testdata/tables.xlsx" 23-script.xlsx
printf '%s\n' $'Region\tUnits\tPrice' $'North\t10\t2.5' $'South\t4\t3' $'North\t10\t2.5' $'East\t7\t1.5' $'West\t12\t2' |
  "$bin" write 23-script.xlsx Orders >/dev/null
cat > 23-script.ts <<'TS'
function main(workbook: ExcelScript.Workbook) {
  const ws = workbook.getWorksheet("Orders")!;
  const removed = ws.getUsedRange()!.removeDuplicates([0, 1, 2], true).getRemoved();
  ws.getRange("D1").setValue("Revenue");
  ws.getRange("D2:D5").setFormula("=B2*C2");
  ws.getRange("A2:D5").getSort().apply([{ key: 3, ascending: false }]);
  ws.getRange("A6:D6").setValues([["Total", "=SUM(B2:B5)", null, "=SUM(D2:D5)"]]);
  ws.getRange("C2:D6").setNumberFormat("$#,##0.00");
  const head = ws.getRange("A1:D1").getFormat();
  head.getFont().setBold(true);
  head.getFont().setColor("#FFFFFF");
  head.getFill().setColor("#1F3864");
  const total = ws.getRange("A6:D6").getFormat();
  total.getFont().setBold(true);
  total.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
  ws.getFreezePanes().freezeRows(1);
  ws.getAutoFilter().apply(ws.getRange("A1:D5"));
  ws.getRange("A:D").getFormat().setColumnWidth(80);

  const t = workbook.getTable("SalesT")!;
  t.addRows(-1, [["West", 4, 1.5], ["Central", 6, 2]]);

  ws.getRange("F1").setValue("A to Z");
  ws.getRange("F2").setFormula("=SORT(A2:A5)");

  const sum = workbook.addWorksheet("Summary");
  sum.getRange("A1:B3").setValues([["Regions", "=COUNTA(Orders!A2:A5)"], ["Best", "=INDEX(Orders!A2:A5,1)"], ["Updated", "2026-07-01"]]);
  return { removed, best: sum.getRange("B2").getValue(), revenue: ws.getRange("D6").getValue() };
}
TS
s script 23-script.xlsx 23-script.ts

# 24. openpyxl stores functions added after Excel 2007 without the _xlfn.
#     prefix Excel needs (AGGREGATE, TEXTJOIN, MAXIFS): 24 as openpyxl left
#     it, 24b after sheets set entered the three formulas again.
cp "$root/internal/xlsx/testdata/unprefixed.xlsx" 24-unprefixed-as-is.xlsx
cp "$root/internal/xlsx/testdata/unprefixed.xlsx" 24b-unprefixed-reentered.xlsx
s set 24b-unprefixed-reentered.xlsx \
  C2 '=IFERROR(INDEX($A$2:$A$4,AGGREGATE(15,6,(ROW($A$2:$A$4)-1)/(LEFT($A$2:$A$4,2)="PK"),1)),"")' \
  D2 '=TEXTJOIN(",",TRUE,A2:A4)' E2 '=MAXIFS(B2:B4,A2:A4,"PK*")'

# 25. Formulas that compute over a range where one value goes, stored as
#     dynamic arrays so Excel computes them as written (older Excel needed
#     Ctrl+Shift+Enter); a whole column (G2) still means the formula's row.
s set 25-array-formulas.xlsx \
  A1 Qty A2 3 A3 1 A4 2 B1 Price B2 10 B3 20 B4 30 \
  D1 'Sum of Qty>1' D2 '=SUM(IF(A2:A4>1,A2:A4))' \
  E1 'Qty doubled' E2 '=A2:A4*2' \
  F1 'Prices>15' F2 '=SUM(--(INDEX(A2:B4,,2)>15))' \
  G1 'Whole column' G2 '=B:B*2' \
  H1 'SUMPRODUCT' H2 '=SUMPRODUCT((A2:A4>1)*B2:B4)' \
  I1 'Max price, Qty<3' I2 '=MAX(IF(A2:A4<3,B2:B4))' \
  J1 'Is first Qty' J2 '=INDEX(A2:B4,,1)=A2'

cat > CHECKLIST.md <<'MD'
# Excel check

Open each file in Excel. First and most important, for every file:
**does Excel open it without "We found a problem with some content"?**
If it asks to repair, note which file and what the repair log says.

## 01-house-style.xlsx
- Title "Q3 sales by region" bold, larger, above the table.
- Header row dark with white bold text, filter buttons, frozen when
  scrolling. Every label left-aligned and fully visible: no filter button
  covers any of it (Units, Price, Revenue, Margin included).
- Data rows alternate white and light gray. Sort Units with the filter
  button: the banding should stay alternating.
- Thin light borders, no gridlines.
- Total row bold with a heavier line above it; the totals add up.
- Price and Revenue show as $ amounts; negative margin as -3.0%;
  dates as 2026-07-01.
- The long note in row 4 wraps, and the row grows tall enough to show it.
  (If it is cut off at one line, that is the thing to report.)

## 02-finance.xlsx
- Typed-in numbers blue; Revenue (formulas) black; After tax (formula
  using the Assumptions sheet) green.

## 03-modern-formulas.xlsx
- Column E: no #NAME? anywhere. Expected: 4, high, "Ada, Bob, Cy, Ada",
  36, 14, 9.
- Click E1: the formula bar shows =XLOOKUP(...) (no _xlfn).
- G2 spills Ada, Bob, Cy downwards; H2 spills 7, 8, 9; I2 spills 1, 2, 3.
- Click G2: the formula bar shows =UNIQUE(A2:A5) without an @ sign, and
  G3/G4 show the formula greyed out (part of the spill).

## 04-structure.xlsx
- Sheets: Sales Data, T, Summary, Summary copy.
- Summary!B1 = 5 and its formula reads ='Sales Data'!C2; B2 = 15.
- The chart on Sales Data still shows the three bars.
- The validation dropdown on the Region cells still works; the comment and
  the link are still there.
- Sheet T keeps its Excel table (Table Design tab appears when clicked).

## 05-restyled.xlsx
- No Comic Sans, yellow, magenta or green borders left.
- "Budget overview" in bold 14 pt; the sheet tab still red.
- Budget and Spent in €; Start as ISO dates; Share as a percentage.

## 06-excel-table.xlsx
- Sheet T: the table keeps its blue table style (no dark sheets header on
  top of it), no extra filter.

## 07-charts.xlsx
- Regions: a clustered column chart (three quarters per region, blue,
  orange, green, with a legend) and a stacked bar chart below it with
  North America at the top and the value axis along the bottom.
- Traffic: a line chart of Visits by Month with dates on the axis, no
  legend (one series), and an area chart of Signups.
- Mix: a pie with five slices in five distinct colours, thin white gaps
  between slices.
- Ads: a scatter of Clicks against Spend, markers only.
- Every chart: thin hairline gridlines, no chart border shadow, titles
  in dark gray, bars and columns thin with small white gaps between
  neighbouring bars.

## 08-two-row-header.xlsx
- Rows 2 and 3 both dark header rows. "Q1" centred across Units and
  Revenue with no line between them; the same for "Q2".
- Units integers, Revenue in $; total row bold.

## 09-excel-table-edits.xlsx
- The table (blue style) now spans rows 3-9: the header, North, South,
  East, West, Central, and the Total row last. Click inside it: Table
  Design shows "Total Row" ticked.
- The headers read Region, Qty, Unit price.
- B9 (the total row) sums all five quantities: 24; its formula reads
  =SUBTOTAL(109,SalesT[Qty]). E4 shows 24 too, with =SUM(SalesT[Qty]).
- Table Design › Total Row still works: untick and tick it again.
- A1 reads "Inserted above the table".

## 10-query-into.xlsx
- A Summary sheet with Region, Units, Revenue, one row per region, sorted
  by Revenue, in the house style; Revenue in $.

## 11-styling.xlsx
- A1 "Budget review 2026": merged over A1:F1, centred, 16 pt bold.
- Header row light blue with black bold text (not white), filter buttons.
- Budget, Spent and Remaining as $ without decimals, the total row and
  Training too (the row appended after the styling, above Total, which
  now sums five rows: 2,230,000 budget).
- The total row has a thin line above and a double line below.
- E3 "Owner" turned 45 degrees; Owner names italic, gray, indented.
- Notes column wide and wrapping.
- Frozen at B4 (rows 1-3 and column A stay put), teal tab, zoom 120%.
- Formulas > Name Manager: no names starting with _sheets (they are
  hidden).

## 12-conditional.xlsx
- Scores: the rows of Late people (Bob, Dee, Fay) light yellow across
  the row; the other rows keep the gray/white banding.
- Score: top 25% (Ada, Gus) light green bold; below 50 (Bob, Fay) red
  bold; Hal's empty score light gray.
- Change: negative values red, shown as (8.0%); solid blue data bars.
- Region: East and West italic (they contain "st").
- Code: the two A-1 codes purple.
- Home > Conditional Formatting > Manage Rules (This Worksheet): the
  rules are listed with the banding (MOD(ROW()-2,2)=1) last.
- Scales: Visits a red-yellow-green scale (410 green, 90 red) and bold
  above average; Growth red-white-green; Rating arrows.

## 13-validation.xlsx
- Orders!A2: a dropdown East, West, North, South, with the prompt
  "Pick a region" on selection; typing "Mars" is refused.
- B2: a dropdown of Pen, Ink, Paper, Stapler (from the Lists sheet).
- C2: 0 or 5000 refused with "Qty is 1 to 1000".
- D2: -1 warns ("A price is usually above 0") but can be kept.
- E2: 2025-12-31 refused. F3: typing A100 (already in F2) refused
  ("Codes are unique"). G2: shows "Up to 20 characters" and refuses
  longer text.

## 14-moved.xlsx
- Scores: A1 reads "Two rows and a column were inserted"; the table now
  starts at B3. Every highlight from 12 is on the same people as before
  (Late rows yellow, Ada and Gus green, Bob and Fay red, ...): the rules
  moved with the data, formulas included.
- Frozen at C4: rows 1-3 and columns A-B stay put, as the header row and
  the Name column did before.

## 15-sorted.xlsx
- Sales sorted by Revenue, highest first: East 150, South 100, North 90,
  West 80; the Total row stays last (17 and 420).
- Running is a running total in the new order: 150, 250, 340, 420.
- The note "West: two only" is on West's name (A5); hovering North's name
  (A4) shows the link to example.com/north, blue and underlined.

## 15b-sorted-table.xlsx
- The Excel table's rows sorted by Units, largest first; its totals row
  stays at the bottom and still adds up; the table style and filter
  buttons are intact.

## 16-copy-move.xlsx
- Data: the block now starts at A8 (header row 8, light blue). Ada and Cy
  (Late) light yellow; C9:C11 have a dropdown (Late, On time); the note
  "Team lead" is on A9, the link on A10.
- Summary: B1 92, B2 92 (=Data!B9: followed the move), B3 #REF! (it read
  B10, where the move pasted).
- Copy: the same block at A2:C5, with the highlight (Ada and Cy), the
  dropdown, the note and the link.

## 17-notes.xlsx
- Data: the note "checked" is on B4 (it was on B2 before two rows were
  inserted), and "Added by sheets" on A1. Hovering shows each, and no
  empty note boxes appear anywhere.

## 17b-thread.xlsx
- Data: the comment thread ("Is 47 right?" with Bob's reply) is on Bob's
  score, wherever the sort put him (B3), and opens as a thread in Excel
  365. The note "Names as entered" stays on A1; a new note on C3.

## 18-from-csv.xlsx
- One sheet "18-orders": Amount as numbers (1,250.00, 99.50, 12,000) in
  the house style, Date as dates, Discount as %, Code keeping its leading
  zeros (00042), "Acme, Inc." and the quoted text intact, and the
  two-line comment wrapping in its cell.

## 19-rules.xlsx
- Data: the title "Orders by region" merged across B1:E1; rows West,
  North, Added, South, East (B3:B7). North, Added and East are light red
  (Flag 1), West and South are not.
- Units (C3:C7) have solid green data bars; Price (D3:D7) blue data bars
  and traffic-light icons. Home › Conditional Formatting › Manage Rules
  shows each rule over rows 3-7, the highlight as `=$E3=1`, and "Cell
  Value > #REF!" (it compared with $C$3, in the deleted row, as Excel's
  would).
- Data › Data Validation on D3: custom, `=D3>$C3`; typing 1 in D4 is
  refused. C3:C7 take whole numbers 0 to 100.
- Other: A1:A3 light red (`=Data!$C$3>1`); B1 has a dropdown listing West,
  North, Added, South, East.
- No repair prompt on opening.

## Round 2 (2026-11-03): what sheets writes since 2026-10-24

Files 20 to 25 need Excel 365 (or 2021): dynamic arrays. sheets marks
every file it saves to be recalculated when Excel opens it, so what Excel
shows is Excel's own result; what matters most here is that each file
opens with no repair prompt, and that Excel reads the formulas as meant.
Close each without saving.

## 20-dynamic-arrays.xlsx
- Data: D2:D5 shows Ada, Bob, Cy, Dee (D1 "Unique"); E2:E5 shows 9, 8, 7,
  6; F2:H3 shows 1, 2, 3 over 4, 5, 6.
- Click D2: the formula bar shows =UNIQUE(A2:A6), with no @ and no { }
  braces; a blue border marks D2:D5. Click D3: its formula shows greyed
  out. The same for E2 (=SORT(FILTER(B2:B6,B2:B6>5),,-1)) and F2
  (=SEQUENCE(2,3)).
- J2 shows #SPILL! (J3 "in the way" blocks it); its warning icon says
  the spill range isn't blank. Delete J3: J2:J5 shows Ada, Bob, Cy, Dee.
- Type 10 in B3: E2 now spills 10, 9, 8, 7, 6, down to E6.
- Edits: D2:D4 shows a, b, c ("Grows": sheets set A4 to c, and the
  result grew from a, b). E2 shows x alone and E3:E4 are empty
  ("Shrinks": B3 and B4 became x, and the result shrank from x, y, z).
- Type a in A4: D's spill shrinks to D2:D3 (a, b). Type y in B4: E's
  spill grows to E2:E3 (x, y).

## 20b-dynamic-arrays-streamed.xlsx
- The same as 20 (the Edits sheet's edits went through sheets' streaming
  editor): D2:D4 a, b, c; E2 x alone; the same actions work.

## 20c-spill-on-new-sheet.xlsx
- A dynamic array set on a sheet the same command added (New). It opens
  with no repair prompt. New shows "Sorted" in A1, a, b, c, d in B3:B6,
  and "below" in B8.
- Click B3: the formula bar shows =SORT(Data!A2:A5), and a blue border
  marks B3:B6. Type e in Data!A2: B3:B6 shows a, b, c, e.

## 21-spill-refs.xlsx
- C2:C4 shows Blue, Green, Red; D2 3; E2:E4 Blue!, Green!, Red!; F2:F4
  2, 1, 2; G2 Green.
- Click D2: the formula bar shows =COUNTA(C2#) (not ANCHORARRAY); E2
  =C2#&"!", F2 =COUNTIF(A2:A6,C2#), G2 =INDEX(C2#,2).
- Type Amber in A3: C2:C5 shows Amber, Blue, Green, Red; D2 4; E's spill
  grows to E5 (Red!); F2:F5 1, 1, 1, 2; G2 Blue.

## 22-shared-sort.xlsx
- Sales sorted by Revenue, highest first: East, South, North, West, Mid.
  Revenue (D2:D6) 150, 100, 90, 80, 10; Label "East #5", "South #1",
  "North #9", "West #2", "Mid #4"; Big TRUE, FALSE, TRUE, FALSE, FALSE;
  Share 34.9%, 23.3%, 20.9%, 18.6%, 2.3%. Total: B7 21, D7 430.
- The rows below (9-12), not sorted: Doubled 80, 110, 14; Tag Rent!,
  Power!, Ink!; Over 50 FALSE, TRUE, FALSE.
- Every number right-aligned and TRUE/FALSE centred: no green triangle
  ("Number Stored as Text") on any cell. Select D2:D6: the status bar
  shows Sum: 430.
- Click D3: =B3*C3; G3: =D3/SUM($D$2:$D$6); C11: =B11*2.

## 22b-shared-sort-streamed.xlsx
- The same as 22 (sorted by the streaming editor).

## 23-script.xlsx
- 23-script.ts beside it is the Office Script that sheets ran.
- Orders: header A1:D1 dark blue with white bold text and filter
  buttons; row 1 frozen; columns A-D wider than the default. Rows (one
  duplicate North removed, sorted by Revenue): North 10 $2.50 $25.00,
  West 12 $2.00 $24.00, South 4 $3.00 $12.00, East 7 $1.50 $10.50. Total
  row bold with a thin line above it: B6 33, D6 $71.50.
- Orders: F2:F5 spills East, North, South, West (=SORT(A2:A5)).
- Sales: the Excel table (blue style) spans A1:C7: North, South, East,
  then the added West 4 1.5 and Central 6 2, and the Total row with B7
  24. Table Design shows "Total Row" ticked. E2 24 (=SUM(SalesT[Units])).
- Summary: B1 4, B2 North; B3 shows ######## (the column is too narrow
  for the date); widen column B: 2026-07-01.

## 24-unprefixed-as-is.xlsx
- As openpyxl wrote it, with no _xlfn. prefixes: D2 and E2 show #NAME?;
  C2 is empty (its IFERROR hides AGGREGATE's #NAME?); F2 shows 60.
- Click D2, press F2 (edit), then Enter: D2 shows PK1,XX2,PK3. (Excel
  reads the function once it is typed again: the reason for 24b.)

## 24b-unprefixed-reentered.xlsx
- After sheets set entered C2:E2 again: C2 PK1, D2 PK1,XX2,PK3, E2 30,
  F2 60. No #NAME? anywhere.
- Click E2: the formula bar shows =MAXIFS(B2:B4,A2:A4,"PK*") (no _xlfn).

## 25-array-formulas.xlsx
- D2 5; E2:E4 spills 6, 2, 4; F2 2; G2 20 (no spill); H2 40; I2 30;
  J2:J4 spills TRUE, FALSE, FALSE.
- Click D2: the formula bar shows =SUM(IF(A2:A4>1,A2:A4)) with no { }
  braces (not a legacy Ctrl+Shift+Enter formula) and no @. The same for
  F2 (=SUM(--(INDEX(A2:B4,,2)>15))) and I2 (=MAX(IF(A2:A4<3,B2:B4))).
- Click G2: the formula bar shows =@B:B*2 (the @ means its own row only).
- Type 5 in A3: D2 becomes 10, E3 10, H2 60.
MD
echo "wrote $(ls "$out" | grep -c xlsx) workbooks to $out; see CHECKLIST.md"
