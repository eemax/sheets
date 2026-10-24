# A German ERP export: semicolons, decimal commas with thousands points,
# Windows-1252, CRLF line ends, quoted fields with semicolons and quotes,
# dates as DD.MM.YYYY.
import json, random, sys
from decimal import ROUND_HALF_UP, Decimal

work, truth = sys.argv[1], sys.argv[2]
rng = random.Random(18)
items = ["Kaffee; gemahlen 500g", 'Tee "Earl Grey" 100g', "Espressomaschine Barista", "Zucker 1kg", "Milchschäumer",
         "Kaffeebecher (6er)", "Filtertüten Gr. 4", "Kakao 250g", "Wasserkocher 1,7l", "Mühle, elektrisch"]
suppliers = ["Müller GmbH", "Schäfer & Söhne", "Café Technik AG", "Großhandel Nord", "Bäckerei-Bedarf e.K."]


def de(x):
    whole, cents = f"{x:.2f}".split(".")
    return f"{int(whole):,}".replace(",", ".") + "," + cents


def field(f):
    return '"' + f.replace('"', '""') + '"' if any(ch in f for ch in ';"') else f


rows = [["Artikelnr", "Bezeichnung", "Lieferant", "Preis", "Einheit", "Gültig ab"]]
prices = []
for i in range(40):
    price = round(rng.choice([rng.uniform(0.5, 20), rng.uniform(20, 300), rng.uniform(900, 2500)]), 2)
    prices.append(price)
    rows.append([f"A-{1000 + i}", rng.choice(items), rng.choice(suppliers), de(price), rng.choice(["Stk", "Pkg", "kg"]),
                 f"{rng.randint(1, 28):02d}.{rng.randint(1, 6):02d}.2026"])
text = "".join(";".join(field(f) for f in r) + "\r\n" for r in rows)
open(f"{work}/artikel.csv", "wb").write(text.encode("cp1252"))
new = [de(float((Decimal(str(p)) * Decimal("1.05")).quantize(Decimal("0.01"), ROUND_HALF_UP))) for p in prices]
json.dump({"rows": rows, "new": new}, open(truth, "w"))
