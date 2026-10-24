import re
from evalkit import Checks, args, close

work, truth, answer = args()
c = Checks()
count = re.search(r"count\D{0,5}([0-9][0-9,]*)", answer, re.I)
revenue = re.search(r"revenue\D{0,5}([0-9][0-9,]*(?:\.[0-9]+)?)", answer, re.I)
n = int(count.group(1).replace(",", "")) if count else None
x = float(revenue.group(1).replace(",", "")) if revenue else None
c.add("count of open orders", n == truth["count"], f"got {n}, want {truth['count']}")
c.add("their revenue", close(x, truth["revenue"], 0.0001), f"got {x}, want {truth['revenue']}")
c.report()
