import re
from evalkit import Checks, args, close

work, truth, answer = args()
c = Checks()
got = {}
for line in answer.splitlines():
    m = re.match(r"\s*[*`-]*\s*([A-Za-z]+)\s*[*`]*\s*:\s*\$?\s*([0-9][0-9,]*(?:\.[0-9]+)?)", line)
    if m:
        got[m.group(1).title()] = float(m.group(2).replace(",", ""))
for region, want in truth["revenue"].items():
    c.add(f"{region} revenue", close(got.get(region), want, 0.0001), f"got {got.get(region)}, want {want}")
c.add("no other regions", set(got) <= set(truth["revenue"]), f"answered {sorted(got)}")
c.report()
