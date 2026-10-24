import csv, io
from evalkit import Checks, args

work, truth, answer = args()
path = f"{work}/artikel.csv"
c = Checks()
raw = open(path, "rb").read()
want_rows = truth["rows"]
c.add("still Windows-1252, as the ERP reads it", "Müller".encode("cp1252") in raw and not raw.startswith(b"\xef\xbb\xbf"), raw[:40])
c.add("CRLF line ends", raw.count(b"\r\n") == len(want_rows) and raw.count(b"\n") == len(want_rows), (raw.count(b"\r\n"), raw.count(b"\n")))
text = raw.decode("cp1252", errors="replace")
rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=";"))
c.add("semicolons, the same columns", rows and rows[0] == want_rows[0] and all(len(r) == len(want_rows[0]) for r in rows), rows[:1])
c.add("the same rows in the same order", [r[0] for r in rows] == [r[0] for r in want_rows], len(rows))
other = all(r[:3] + r[4:] == w[:3] + w[4:] for r, w in zip(rows[1:], want_rows[1:]))
c.add("other fields unchanged (names, suppliers, units, dates)", other,
      next(((r, w) for r, w in zip(rows[1:], want_rows[1:]) if r[:3] + r[4:] != w[:3] + w[4:]), None))
got = [r[3] for r in rows[1:]]


def num(s):
    return float(s.replace(".", "").replace(",", "."))


c.add("prices up 5%, to the cent", len(got) == len(truth["new"]) and all(abs(num(g) - num(w)) < 0.006 for g, w in zip(got, truth["new"])),
      next(((g, w) for g, w in zip(got, truth["new"]) if abs(num(g) - num(w)) >= 0.006), None) if all("," in g for g in got) else got[:3])


def de(x):
    whole, cents = f"{x:.2f}".split(".")
    return f"{int(whole):,}".replace(",", ".") + "," + cents


def written(g):
    try:
        return g == de(num(g))
    except ValueError:
        return False


c.add("prices written as the file writes them (1.249,00)", got and all(written(g) for g in got), [g for g in got if not written(g)][:3])
c.report()
