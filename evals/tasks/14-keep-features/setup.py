# A dashboard as Excel users make them: an Excel table with a formula
# column, sparklines, Excel 2010 data bars, a dropdown, a note, a text box,
# a logo, a link, a chart, frozen panes, a print header and print area.
# Results stored, as Excel saves them.
import json, struct, sys, zlib
import xlsxwriter
from evalkit import excel_saved


def png(w, h, rgb):
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


work, truth = sys.argv[1], sys.argv[2]
logo = f"{work}/../logo.png"
open(logo, "wb").write(png(60, 24, (31, 78, 121)))
path = f"{work}/dashboard.xlsx"
wb = xlsxwriter.Workbook(path)
ws = wb.add_worksheet("Dashboard")
money = wb.add_format({"num_format": "#,##0"})
data = [("North", 1200, 1310, 1405, 1480), ("South", 860, 905, 940, 1010), ("East", 640, 700, 690, 755), ("West", 1500, 1460, 1525, 1610)]
ws.write("A1", "Regional dashboard", wb.add_format({"bold": True, "font_size": 14}))
ws.insert_image("I1", logo)
ws.add_table(2, 0, 2 + len(data), 6, {
    "name": "Sales", "data": [list(r) for r in data],
    "columns": [{"header": "Region"}] + [{"header": q, "format": money} for q in ("Q1", "Q2", "Q3", "Q4")]
               + [{"header": "Total", "formula": "=SUM(Sales[@[Q1]:[Q4]])", "format": money}, {"header": "Status"}]})
for i, r in enumerate(data):
    ws.write_formula(3 + i, 5, "=SUM(Sales[@[Q1]:[Q4]])", money, sum(r[1:]))
    ws.write(3 + i, 6, "On track")
    ws.add_sparkline(3 + i, 7, {"range": f"Dashboard!B{4 + i}:E{4 + i}", "markers": True})
ws.write("H3", "Trend")
ws.conditional_format("F4:F12", {"type": "data_bar", "data_bar_2010": True, "bar_solid": True, "bar_color": "#5B9BD5"})
ws.data_validation("G4:G20", {"validate": "list", "source": ["On track", "At risk", "Behind"]})
ws.write_comment("A4", "North includes the Baltics from Q3.")
ws.insert_textbox("K4", "Figures in $000. Source: ERP export.", {"width": 220, "height": 50})
ws.write_url("A11", "https://example.com/finance", string="Finance wiki")
ws.freeze_panes(3, 1)
ws.set_header("&LRegional dashboard&RPage &P")
ws.print_area("A1:H11")
ch = wb.add_chart({"type": "column"})
for i, q in enumerate(("Q1", "Q2", "Q3", "Q4")):
    col = "BCDE"[i]
    ch.add_series({"name": q, "categories": "=Dashboard!$A$4:$A$7", "values": f"=Dashboard!${col}$4:${col}$7"})
ch.set_title({"name": "Sales by quarter"})
ws.insert_chart("A13", ch)
ws.set_column("A:A", 14)
wb.close()
excel_saved(path)
json.dump({"north": 1200 + 1310 + 1405 + 1520, "nordics": 410 + 455 + 470 + 512}, open(truth, "w"))
