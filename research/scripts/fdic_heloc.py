"""FDIC 콜리포트로 은행 HELOC 장부를 다시 집계한다.

사용법:
    python3 research/scripts/fdic_heloc.py            # 최신 분기 단면 + 전체 시계열 요약
    python3 research/scripts/fdic_heloc.py 20260930   # 특정 분기 단면

필드 (FDIC RIS, 단위 $천):
    LNRELOC  1-4가구 주택담보 회전한도 잔액 (RC-C 1.c.(1), 은행 HELOC 인출 잔액)
    UCLOC    홈에쿼티 미사용 한도 (RC-L 1.a)
    NTRELOCQ HELOC 순상각액 (분기)
    LNLSGR   총대출, DEP 예금, LNRERSFM 1순위 주택담보, ASSET 총자산
"""
import json
import sys
import urllib.parse
import urllib.request

API = "https://api.fdic.gov/banks/financials"


def get(params):
    url = API + "?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(url, timeout=120) as r:
        return json.load(r)


def series():
    j = get({
        "filters": "REPDTE:[20000101 TO 20991231]",
        "agg_by": "REPDTE",
        "agg_sum_fields": "LNRELOC,UCLOC,NTRELOCQ,LNLSGR",
        "agg_limit": 400, "limit": 0, "format": "json",
    })
    rows = sorted((r["data"] for r in j["data"]), key=lambda x: x["REPDTE"])
    print("분기      HELOC$B  미사용$B  인출률  순상각(연율)  대출대비")
    for x in rows:
        h, u = x["sum_LNRELOC"] / 1e6, x["sum_UCLOC"] / 1e6
        nco = x["sum_NTRELOCQ"] * 4 / x["sum_LNRELOC"] * 100
        sh = x["sum_LNRELOC"] / x["sum_LNLSGR"] * 100
        print(f'{x["REPDTE"]}  {h:7.1f}  {u:7.1f}  {h/(h+u)*100:5.1f}%  {nco:6.2f}%  {sh:5.2f}%')
    return rows[-1]["REPDTE"]


def cross_section(repdte):
    j = get({
        "filters": f"REPDTE:{repdte}",
        "fields": "CERT,NAME,ASSET,DEP,LNLSGR,LNRERSFM,LNRELOC,UCLOC",
        "limit": 10000, "format": "json",
    })
    d = [{k: (v or 0) if k != "NAME" else v for k, v in r["data"].items()} for r in j["data"]]
    n = len(d)
    tot = sum(x["LNRELOC"] for x in d)
    zero = [x for x in d if x["LNRELOC"] == 0]
    low = [x for x in d if x["LNLSGR"] and x["LNRELOC"] / x["LNLSGR"] < 0.01]
    dep = sum(x["DEP"] for x in d)
    print(f"\n{repdte} 기관 수 {n}, HELOC 잔액 ${tot/1e6:.1f}B")
    print(f"HELOC 0인 기관 {len(zero)} (그중 1순위 모기지 보유 {sum(1 for x in zero if x['LNRERSFM']>0)})")
    print(f"HELOC<1% 기관 {len(low)} ({len(low)/n*100:.1f}%), 예금 점유 {sum(x['DEP'] for x in low)/dep*100:.1f}%")
    top = sorted(d, key=lambda x: -x["LNRELOC"])
    for k in (10, 20, 50):
        print(f"상위 {k} 집중도 {sum(x['LNRELOC'] for x in top[:k])/tot*100:.1f}%")


if __name__ == "__main__":
    latest = series()
    cross_section(sys.argv[1] if len(sys.argv) > 1 else latest)
