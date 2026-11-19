"""은행 지도: 예대율 × HELOC 시스템(FDIC 콜리포트 단면).

투포1 v5 3-4 '누가 먼저 오르나'(도표 3-9)의 숫자를 다시 만든다.

사용법:
    python3 -I research/v5/fdic_bank_map.py                 # 2026-06-30 단면을 API에서 받아 집계
    python3 -I research/v5/fdic_bank_map.py 20260630 x.json # 이미 받은 JSON으로 집계

정의 (FDIC RIS 필드, 단위 $천):
    예대율 = LNLSGR(총대출) / DEP(총예금)
    HELOC 비중 = LNRELOC(1-4가구 주택담보 회전한도 잔액) / LNLSGR
    'HELOC 시스템 없음'의 대용 지표 = HELOC 비중 1% 미만(0 포함)
    1순위 보유 = LNRERSFM(1-4가구 1순위 주택담보) > 0
"""
import csv
import json
import sys
import urllib.parse
import urllib.request

API = "https://api.fdic.gov/banks/financials"
FIELDS = "CERT,NAME,ASSET,DEP,LNLSGR,LNRERSFM,LNRELOC"


def load(repdte, path=None):
    if path:
        with open(path) as f:
            j = json.load(f)
    else:
        q = urllib.parse.urlencode({"filters": f"REPDTE:{repdte}", "fields": FIELDS,
                                    "limit": 10000, "format": "json"})
        with urllib.request.urlopen(f"{API}?{q}", timeout=120) as r:
            j = json.load(r)
    rows = []
    for rec in j["data"]:
        x = rec["data"]
        rows.append({k: (x.get(k) or 0) if k != "NAME" else x.get(k) for k in FIELDS.split(",")})
    return rows


def ldr_bucket(x):
    if not x["DEP"]:
        return "예금 없음"
    r = x["LNLSGR"] / x["DEP"]
    return "예대율 90% 이상" if r >= 0.9 else ("예대율 70% 미만" if r < 0.7 else "예대율 70~90%")


def heloc_bucket(x):
    s = x["LNRELOC"] / x["LNLSGR"] if x["LNLSGR"] else 0
    return "HELOC 거의 없음(<1%)" if s < 0.01 else "HELOC 있음(1% 이상)"


def main():
    repdte = sys.argv[1] if len(sys.argv) > 1 else "20260630"
    rows = load(repdte, sys.argv[2] if len(sys.argv) > 2 else None)
    out = []
    for universe, keep in (("전체 은행", lambda x: True),
                           ("1순위 모기지 보유 은행", lambda x: x["LNRERSFM"] > 0)):
        sub = [x for x in rows if keep(x)]
        cells = {}
        for x in sub:
            k = (ldr_bucket(x), heloc_bucket(x))
            c = cells.setdefault(k, {"n": 0, "dep": 0, "heloc": 0, "first": 0, "zero": 0})
            c["n"] += 1
            c["dep"] += x["DEP"]
            c["heloc"] += x["LNRELOC"]
            c["first"] += x["LNRERSFM"]
            c["zero"] += x["LNRELOC"] == 0
        for (lb, hb), c in sorted(cells.items()):
            out.append({"기준일": repdte, "모집단": universe, "예대율": lb, "HELOC": hb,
                        "은행 수": c["n"], "그중 HELOC 0": c["zero"],
                        "예금 $B": round(c["dep"] / 1e6, 1),
                        "HELOC 잔액 $B": round(c["heloc"] / 1e6, 1),
                        "1순위 주택담보 $B": round(c["first"] / 1e6, 1)})
    w = csv.DictWriter(sys.stdout, fieldnames=list(out[0].keys()))
    w.writeheader()
    w.writerows(out)


if __name__ == "__main__":
    main()
