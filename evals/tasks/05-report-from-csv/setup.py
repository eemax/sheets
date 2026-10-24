import csv, json, random, sys

work, truth = sys.argv[1], sys.argv[2]
random.seed(5)
units, revenue = {}, {}
with open(f"{work}/sales.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Date", "Region", "Product", "Units", "Price"])
    for i in range(40):
        region = random.choice(["North", "South", "East", "West"])
        u, p = random.randint(1, 30), random.choice([12.5, 30.0, 7.25, 120.0])
        w.writerow([f"2026-07-{1 + i % 28:02d}", region, random.choice(["Widget", "Gadget"]), u, f"{p:.2f}"])
        units[region] = units.get(region, 0) + u
        revenue[region] = revenue.get(region, 0) + u * p
json.dump({"units": units, "revenue": revenue}, open(truth, "w"))
