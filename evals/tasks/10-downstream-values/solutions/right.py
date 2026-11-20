# The right way: change the input with sheets, which keeps the formulas and
# stores their new results.
import json, os, subprocess, sys

work = sys.argv[1]
path = os.path.join(work, "prices.xlsx")
sheets = os.environ.get("SHEETS_BIN", "sheets")
subprocess.run([sheets, "set", path, "Products!B2", "15"], check=True, capture_output=True)
out = subprocess.run([sheets, "read", path, "Orders!C6", "--json"], check=True, capture_output=True, text=True)
total = json.loads(out.stdout)["rows"][0]["values"][0]
print(f"Changed Widget's price to 15 in prices.xlsx; the order total is now {total}.")
