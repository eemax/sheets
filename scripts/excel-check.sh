#!/usr/bin/env bash
# Builds workbooks to open in real Excel, the one check the test suite
# cannot do. See excel-check/CHECKLIST.md after running:
#
#   scripts/excel-check.sh [output-dir]
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
MD
echo "wrote $(ls "$out" | grep -c xlsx) workbooks to $out; see CHECKLIST.md"
