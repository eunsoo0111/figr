import csv, sys, io, json, urllib.request, collections, gzip
lei, name = sys.argv[1], sys.argv[2]
url = f"https://ffiec.cfpb.gov/v2/data-browser-api/view/csv?years=2025&leis={lei}&actions_taken=1"
r = urllib.request.urlopen(url, timeout=600)
rd = csv.DictReader(io.TextIOWrapper(r, encoding="utf-8"))
agg = collections.defaultdict(lambda: [0, 0.0, 0, 0.0])   # n, amt, n_sold, amt_sold
for row in rd:
    oe = row["open-end_line_of_credit"]; lien = row["lien_status"]
    cat = ("HELOC" if oe == "1" else "CLOSED") + "_L" + lien
    amt = float(row["loan_amount"] or 0)
    sold = row["purchaser_type"] not in ("0", "")
    a = agg[cat]; a[0] += 1; a[1] += amt
    if sold: a[2] += 1; a[3] += amt
out = {k: v for k, v in agg.items()}
json.dump({"lei": lei, "name": name, "agg": out}, open(f"hmda_{name}.json", "w"), indent=1)
for k, v in sorted(out.items()):
    print(name, k, v[0], round(v[1]/1e6,1), "sold%", round(v[3]/v[1]*100,1) if v[1] else None)
