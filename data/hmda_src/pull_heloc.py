"""Stream HMDA loan-level originations for one LEI-year, keep open-end (HELOC) rows.
usage: python3 -I pull_heloc.py LEI YEAR OUTDIR"""
import csv, io, json, os, sys, urllib.request, time
lei, year, out = sys.argv[1], sys.argv[2], sys.argv[3]
url = f"https://ffiec.cfpb.gov/v2/data-browser-api/view/csv?years={year}&leis={lei}&actions_taken=1"
keep = ["activity_year","lei","state_code","loan_purpose","lien_status","loan_amount",
        "combined_loan_to_value_ratio","interest_rate","property_value","income",
        "debt_to_income_ratio","occupancy_type","purchaser_type","loan_term"]
for attempt in range(4):
    try:
        resp = urllib.request.urlopen(url, timeout=600)
        rdr = csv.DictReader(io.TextIOWrapper(resp, encoding="utf-8", newline=""))
        s = {"lei": lei, "year": int(year), "all_orig_n": 0, "all_orig_amt": 0,
             "heloc_n": 0, "heloc_amt": 0, "heloc_1st_n": 0, "heloc_1st_amt": 0,
             "heloc_2nd_n": 0, "heloc_2nd_amt": 0, "closed_2nd_n": 0, "closed_2nd_amt": 0}
        rows = []
        for r in rdr:
            amt = int(float(r["loan_amount"] or 0))
            s["all_orig_n"] += 1; s["all_orig_amt"] += amt
            oe = r["open-end_line_of_credit"] == "1"
            biz = r["business_or_commercial_purpose"] == "1"
            rev = r["reverse_mortgage"] == "1"
            if oe and not biz and not rev:
                s["heloc_n"] += 1; s["heloc_amt"] += amt
                k = "1st" if r["lien_status"] == "1" else "2nd"
                s[f"heloc_{k}_n"] += 1; s[f"heloc_{k}_amt"] += amt
                rows.append({c: r.get(c, "") for c in keep})
            elif (not oe) and r["lien_status"] == "2" and not biz and not rev:
                s["closed_2nd_n"] += 1; s["closed_2nd_amt"] += amt
        break
    except Exception as e:
        err = repr(e); time.sleep(5 * (attempt + 1))
else:
    s = {"lei": lei, "year": int(year), "error": err}; rows = []
os.makedirs(out, exist_ok=True)
json.dump(s, open(f"{out}/sum_{lei}_{year}.json", "w"))
if rows:
    with open(f"{out}/heloc_{lei}_{year}.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=keep); w.writeheader(); w.writerows(rows)
print(lei, year, s.get("heloc_n"), s.get("heloc_amt"), s.get("error", ""))
