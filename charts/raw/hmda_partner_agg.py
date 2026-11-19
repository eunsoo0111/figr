# HMDA LAR(LEI·연도별) 스트리밍 집계: 실행(action_taken=1) 건 중 HELOC(개방형)·폐쇄형 1/2순위 건수·금액·당해 매각
# 사용: python3 -I hmda_partner_agg.py LEI YEAR 이름  → hmda_partners/{이름}_{YEAR}.json
import csv, io, json, sys, os, urllib.request, collections, time
lei, year, name = sys.argv[1], sys.argv[2], sys.argv[3]
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hmda_partners", f"{name}_{year}.json")
if os.path.exists(out):
    sys.exit(0)
url = f"https://ffiec.cfpb.gov/v2/data-browser-api/view/csv?years={year}&leis={lei}&actions_taken=1"
for attempt in range(4):
    try:
        r = urllib.request.urlopen(url, timeout=900)
        rd = csv.DictReader(io.TextIOWrapper(r, encoding="utf-8"))
        agg = collections.defaultdict(lambda: [0, 0, 0, 0])   # 건수, 금액, 매각 건수, 매각 금액
        n = 0
        for row in rd:
            n += 1
            cat = ("HELOC" if row["open-end_line_of_credit"] == "1" else "CLOSED") + "_L" + row["lien_status"]
            amt = int(float(row["loan_amount"] or 0))
            a = agg[cat]; a[0] += 1; a[1] += amt
            if row["purchaser_type"] not in ("0", ""):
                a[2] += 1; a[3] += amt
        json.dump({"lei": lei, "year": int(year), "name": name, "rows": n, "agg": agg}, open(out, "w"), indent=1)
        print(name, year, n, {k: v[1] for k, v in agg.items() if k.startswith("HELOC")})
        break
    except Exception as e:
        print("retry", name, year, e, file=sys.stderr); time.sleep(10 * (attempt + 1))
