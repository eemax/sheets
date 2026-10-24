import csv, json, random, sys

work, truth = sys.argv[1], sys.argv[2]
random.seed(12)
regions = ["North", "South", "East", "West", "Central"]
revenue = {r: 0.0 for r in regions}
with open(f"{work}/sales.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Region", "Rep", "Units", "Price"])
    for i in range(60):
        r = random.choice(regions)
        units, price = random.randint(1, 50), round(random.uniform(10, 200), 2)
        revenue[r] += units * price
        w.writerow([f"2026-0{1 + i % 6}-{1 + i % 28:02d}", r, random.choice(["Ann", "Bo", "Cy"]), units, price])
targets = {r: round(revenue[r] * random.choice([0.8, 0.9, 1.1, 1.25]), -2) for r in regions}
with open(f"{work}/targets.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Region", "Target"])
    for r in regions:
        w.writerow([r, targets[r]])
json.dump({"revenue": revenue, "targets": targets, "rows": 60}, open(truth, "w"))
