"""IP1 ②③ + 바퀴 계산: 「왜 선형이 아닌가」와 「은행은 사고팔기를 둘 다 하나」를 숫자로.

입력
  - research/data/ip1_v4_quarterly_kpi.csv (8-K 분기 KPI, 활성 파트너)
  - FDIC BankFind API (콜리포트 2026-06-30 전수, 4,313곳)  ※ 네트워크 필요
  - FRED Z.1 (주택담보 보유 주체별 잔액)                    ※ 네트워크 필요
출력: research/data/ip1_23_*.csv

실행: python3 -I research/scripts/ip1_23_calcs.py research
"""
import csv
import json
import math
import os
import sys
import urllib.parse
import urllib.request

ROOT = sys.argv[1] if len(sys.argv) > 1 else "research"
DATA = os.path.join(ROOT, "data")


def w(name, header, rows):
    with open(os.path.join(DATA, name), "w", encoding="utf-8", newline="") as fh:
        cw = csv.writer(fh)
        cw.writerow(header)
        cw.writerows(rows)
    print(f"\n[{name}]")
    print(" | ".join(header))
    for r in rows:
        print(" | ".join(str(x) for x in r))


def fget(url):
    with urllib.request.urlopen(url, timeout=120) as r:
        return r.read().decode("utf-8")


# ---------------------------------------------------------------------------
# 1. 파트너 동학: 순증이 기존 파트너 수에 비례해 커지나
# ---------------------------------------------------------------------------
kpi = list(csv.DictReader(open(os.path.join(DATA, "ip1_v4_quarterly_kpi.csv"), encoding="utf-8")))
partners = {"3Q25": 250}  # IPO(2025-09) 무렵 「약 250」(회사). 2Q25 행의 「약 250」을 3Q25 기준점으로 씀
for r in kpi:
    p = r["active_partners"].replace("약 ", "").strip()
    if p.isdigit():
        partners[r["quarter"]] = int(p)
vol = {r["quarter"]: float(r["loan_marketplace_vol_$M_standalone"]) for r in kpi}
con = {r["quarter"]: float(r["connect_vol_$M"]) for r in kpi}
qs = ["3Q25", "4Q25", "1Q26", "2Q26", "3Q26"]
rows = []
prev = None
for q in qs:
    n = partners[q]
    add = n - partners[prev] if prev else ""
    g = round((n / partners[prev] - 1) * 100, 1) if prev else ""
    rows.append([q, n, add, g, vol[q], round(vol[q] / n, 1), con[q]])
    prev = q
w("ip1_23_partner_dynamics.csv",
  ["quarter", "active_partners", "net_adds", "net_adds_%_of_prior_base", "volume_$M_standalone", "volume_per_partner_$M", "connect_$M"],
  rows)

# 기존 파트너 vs 새 파트너 (3Q25 → 3Q26). CEO: 1년 전부터 함께한 파트너는 물량 40% 넘게 증가
base, now = vol["3Q25"], vol["3Q26"]
same = 0.40
existing_now = base * (1 + same)
new_now = now - existing_now
grow = now - base
n_old, n_new = partners["3Q25"], partners["3Q26"] - partners["3Q25"]
rows = [
    ["3Q25 거래량 (전부 기존 파트너 + 동사 브랜드)", round(base), "", ""],
    ["3Q26 거래량", round(now), "", ""],
    ["늘어난 거래량", round(grow), 100.0, ""],
    ["  기존 파트너가 더 한 몫 (+40% 가정)", round(existing_now - base), round((existing_now - base) / grow * 100, 1), "CFO: 「올해 성장의 40% 넘게 기존 고객」과 맞음"],
    ["  1년 새 붙은 파트너 몫 (나머지)", round(new_now), round(new_now / grow * 100, 1), ""],
    ["기존 파트너 1곳당 분기 물량 $M (3Q26)", round(existing_now / n_old, 1), "", f"기존 {n_old}곳 (동사 브랜드 물량이 섞여 과대)"],
    ["새 파트너 1곳당 분기 물량 $M (3Q26)", round(new_now / n_new, 1), "", f"새 파트너 {n_new}곳, 아직 램프 중"],
    ["새 파트너가 기존 평균의 1/2에 닿으면 더해지는 분기 물량 $M", round(n_new * existing_now / n_old * 0.5 - new_now), "", "가정: 새 파트너는 작음(SMB·신협 섞임). 램프 여력의 예시일 뿐"],
]
w("ip1_23_existing_vs_new.csv", ["item", "value", "share_of_growth_%", "note"], rows)

# ---------------------------------------------------------------------------
# 2. 은행: 1순위에서는 이미 파는 쪽이자 사는 쪽 (FDIC 콜리포트 2026-06-30 전수)
# ---------------------------------------------------------------------------
F = "CERT,NAME,ASSET,DEP,LNLSGR,LNRERSFM,LNRELOC,LNRERSF2,SCMTGBK,SCGTY,SCCOL,SCRMBPI,LNLSSALE,NETGNSLN"
url = "https://api.fdic.gov/banks/financials?" + urllib.parse.urlencode(
    {"filters": "REPDTE:20260630", "fields": F, "limit": 10000, "format": "json"})
banks = [r["data"] for r in json.loads(fget(url))["data"]]


def g(x, k):
    return x.get(k) or 0


for x in banks:
    x["AGMBS"] = g(x, "SCGTY") + g(x, "SCCOL")  # 에이전시 발행·보증 주거용 MBS(패스스루 + CMO)
B = 1e6  # $천 → $B
tot = {k: sum(g(x, k) for x in banks) for k in ["ASSET", "DEP", "LNRERSFM", "LNRELOC", "LNRERSF2", "SCRMBPI", "AGMBS", "LNLSSALE"]}
sellers = [x for x in banks if g(x, "LNLSSALE") > 0 or g(x, "NETGNSLN") > 0]
hel = [x for x in banks if g(x, "LNRELOC") > 0]
rows = [
    ["은행 수", len(banks), "FDIC 2026-06-30"],
    ["에이전시 주거용 MBS 보유 $B", round(tot["AGMBS"] / B, 1), "SCGTY + SCCOL"],
    ["사설(비에이전시) 주거용 MBS 보유 $B", round(tot["SCRMBPI"] / B, 1), "SCRMBPI"],
    ["에이전시 ÷ 사설 (배)", round(tot["AGMBS"] / tot["SCRMBPI"], 1), "표준화된 주택담보 채권은 사고, 비표준은 거의 안 산다"],
    ["1순위 주택담보 대출 보유 $B", round(tot["LNRERSFM"] / B, 1), ""],
    ["HELOC 인출 잔액 $B", round(tot["LNRELOC"] / B, 1), ""],
    ["대출을 파는 은행 수 (판매대기 대출 또는 대출 매각 이익 > 0)", len(sellers), ""],
    ["  그중 에이전시 MBS도 든 은행 %", round(sum(1 for x in sellers if x["AGMBS"] > 0) / len(sellers) * 100, 1), "파는 은행이 사는 은행"],
    ["  파는 은행이 쥔 은행 에이전시 MBS 몫 %", round(sum(x["AGMBS"] for x in sellers) / tot["AGMBS"] * 100, 1), ""],
    ["HELOC 보유 은행 중 에이전시 MBS도 든 곳 (HELOC 잔액 가중) %", round(sum(g(x, "LNRELOC") for x in hel if x["AGMBS"] > 0) / sum(g(x, "LNRELOC") for x in hel) * 100, 1), ""],
]
top = sorted(banks, key=lambda x: -g(x, "LNRELOC"))[:20]
th, tm = sum(g(x, "LNRELOC") for x in top), sum(x["AGMBS"] for x in top)
rows += [
    ["HELOC 상위 20개 은행 HELOC $B", round(th / B, 1), ""],
    ["HELOC 상위 20개 은행 에이전시 MBS $B", round(tm / B, 1), ""],
    ["  MBS ÷ HELOC (배)", round(tm / th, 1), "HELOC 장부의 9배를 표준화된 남의 모기지 채권으로 들고 있다"],
]
w("ip1_23_bank_two_sided_summary.csv", ["item", "value", "note"], rows)

# HMDA 2025 같은 해 매각 비중(기존 bank-channel 표)과 FDIC 보유를 한 줄로
hmda = {r["lender"]: r for r in csv.DictReader(open(os.path.join(ROOT, "bank-channel", "data", "hmda2025_sold_share_by_lender.csv"), encoding="utf-8"))}
name_map = {"Citizens": "CITIZENS BANK NATIONAL ASSN", "PNC": "PNC BANK NATIONAL ASSN", "Fifth Third": "FIFTH THIRD BANK NA", "Flagstar": "FLAGSTAR BANK NATIONAL ASSN"}
rows = []
for k, nm in name_map.items():
    x = next(b for b in banks if b["NAME"] == nm)
    h = hmda[k]
    rows.append([k, h["first_lien_sold_same_year_%"], h["first_lien_sold_to_GSE_%"], round(x["AGMBS"] / B, 1),
                 h["heloc_sold_same_year_%"], round(g(x, "LNRELOC") / B, 1), round(g(x, "LNLSGR") / max(g(x, "DEP"), 1) * 100)])
w("ip1_23_bank_first_lien_vs_heloc.csv",
  ["bank", "1st_lien_sold_same_year_%(HMDA25)", "1st_lien_sold_to_GSE_%", "agency_MBS_held_$B(FDIC 2Q26)", "HELOC_sold_same_year_%", "HELOC_balance_$B", "loan_to_deposit_%"],
  rows)

rows = []
for x in top:
    rows.append([x["NAME"], round(g(x, "LNRELOC") / B, 1), round(x["AGMBS"] / B, 1), round(g(x, "LNRERSFM") / B, 1),
                 round(g(x, "LNLSSALE") / B, 2), round(g(x, "DEP") / B), round(g(x, "LNLSGR") / max(g(x, "DEP"), 1) * 100)])
w("ip1_23_top20_heloc_banks.csv", ["bank", "HELOC_$B", "agency_MBS_$B", "first_lien_$B", "loans_held_for_sale_$B", "deposits_$B", "LDR_%"], rows)

# ---------------------------------------------------------------------------
# 3. 1순위가 걸은 길: 누가 미국 주택담보를 들고 있나 (FRED Z.1)
# ---------------------------------------------------------------------------
def fred(sid):
    t = fget("https://fred.stlouisfed.org/graph/fredgraph.csv?id=" + sid).splitlines()[1:]
    return {a: float(b) for a, b in (l.split(",") for l in t) if b not in ("", ".")}


T = fred("ASHMA")                 # 전 부문 주택담보(1~4가구) 자산
AG = fred("BOGZ1FL423065105Q")    # GSE + 에이전시·GSE 보증 풀
PL = fred("BOGZ1FL673065105Q")    # 사설 ABS 발행기구
DP = fred("BOGZ1FL703065105Q")    # 민간 예금기관(은행·저축기관·신협)
rows = []
for d in ["1970-01-01", "1975-01-01", "1980-01-01", "1985-01-01", "1990-01-01", "1995-01-01", "2000-01-01", "2007-01-01", "2015-01-01", "2026-04-01"]:
    t = T[d]
    rows.append([d[:7], round(t / 1e6, 2), round(DP[d] / t * 100, 1), round(AG[d] / t * 100, 1), round(PL[d] / t * 100, 1), round((AG[d] + PL[d]) / t * 100, 1)])
w("ip1_23_first_lien_history.csv",
  ["date", "home_mortgages_$T", "held_by_depositories_%", "agency_GSE_%", "private_ABS_%", "securitized_or_GSE_%"], rows)

# HELOC은 지금 어디쯤: NY Fed HELOC 잔액 vs 증권화된 HELOC 잔액(SFA 2026-09)
ny = list(csv.DictReader(open(os.path.join(DATA, "nyfed_hhdc_heloc_2003q1_2026q3.csv"), encoding="utf-8")))[-1]
heloc_bal = float(ny["heloc_tn"]) * 1000
sec_heloc = 21.5
w("ip1_23_heloc_securitized_share.csv", ["item", "value", "source"], [
    ["HELOC 잔액 $B (" + ny["quarter"] + ")", round(heloc_bal, 1), "NY Fed HHDC"],
    ["증권화된 HELOC 잔액 $B", sec_heloc, "SFA Research Corner 2026-09-17 (24년 약 $5B)"],
    ["증권화 비중 %", round(sec_heloc / heloc_bal * 100, 1), "1순위로 치면 1970년 무렵(5%)"],
    ["은행 HELOC 잔액 $B (FDIC 2Q26)", round(tot["LNRELOC"] / B, 1), "FDIC"],
])

# ---------------------------------------------------------------------------
# 4. 경매 깊이: 봉인 입찰에서 입찰자가 늘면 낙찰가가 오르고 흔들림이 준다
#    (대칭 입찰자, 수익 동등성 → 판매자 기대수입 = 둘째로 높은 평가액의 기댓값)
# ---------------------------------------------------------------------------
def phi(z):
    return math.exp(-z * z / 2) / math.sqrt(2 * math.pi)


def Phi(z):
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


def order_stat(n, k, moment=1):
    """표준정규 n개 중 k번째로 큰 값의 moment차 적률 (수치 적분)."""
    # 밀도: n!/((k-1)!(n-k)!) * F^(n-k) * (1-F)^(k-1) * f
    c = math.factorial(n) / (math.factorial(k - 1) * math.factorial(n - k))
    s, h = 0.0, 0.001
    z = -8.0
    while z < 8.0:
        F = Phi(z)
        s += (z ** moment) * c * (F ** (n - k)) * ((1 - F) ** (k - 1)) * phi(z) * h
        z += h
    return s


rows = []
base5 = order_stat(5, 2)
for n in [3, 5, 10, 15, 20, 25, 30, 35, 50]:
    m = order_stat(n, 2)
    sd = math.sqrt(max(order_stat(n, 2, 2) - m * m, 0))
    rows.append([n, round(m, 2), round(m - base5, 2), round(sd, 2)])
w("ip1_23_auction_depth.csv",
  ["bidders", "expected_clearing_price_in_sigma(2nd_highest)", "vs_5_bidders_sigma", "sd_of_clearing_price_sigma"], rows)

# ---------------------------------------------------------------------------
# 5. 은행이 「팔고 사는」 자본의 논리: 같은 HELOC을 대출로 들면 100%, AAA 채권으로 들면 20%
#    미국 표준방법: 후순위 주택담보 100%, 증권화 노출은 SSFA(하한 20%)
# ---------------------------------------------------------------------------
def ssfa(A, D=1.0, KG=0.08, W=0.0, p=0.5):
    KA = (1 - W) * KG + 0.5 * W
    if D <= KA:
        return 12.50
    a = -1 / (p * KA)
    if A >= KA:
        u, l = D - KA, A - KA
        K = (math.exp(a * u) - math.exp(a * l)) / (a * (u - l))
        return max(K * 12.5, 0.20)
    # 걸치는 구간
    u, l = D - KA, 0.0
    Kpart = (math.exp(a * u) - math.exp(a * l)) / (a * (u - l))
    rw = ((KA - A) / (D - A)) * 12.5 + ((D - KA) / (D - A)) * 12.5 * Kpart
    return max(rw, 0.20)


rows = [["HELOC 대출 그대로 (후순위 주택담보)", "", 100, "12 CFR 217.32(g)"]]
for name, A in [("동사 딜 AAA, 완충 35.1% (HE2)", 0.351), ("동사 딜 AAA, 완충 22.85% (HE8)", 0.2285), ("완충 15%로 가정", 0.15)]:
    rows.append([name, A, round(ssfa(A) * 100, 1), "SSFA, KG 8%(기초 위험가중치 100%), W 0, p 0.5, 하한 20%"])
rows.append(["→ 같은 HELOC 위험을 AAA로 들면 필요 자본", "", "약 1/5", "1순위의 lender swap(대출을 넘기고 MBS로 받아 보유)과 같은 논리"])
w("ip1_23_bank_capital_ssfa.csv", ["holding", "attachment", "risk_weight_%", "basis"], rows)

# ---------------------------------------------------------------------------
# 6. 한 창구의 관계 수와 「희소한 쪽」
# ---------------------------------------------------------------------------
P, Bm = 612, 30  # 활성 파트너, Connect 매수자(EC v2 「25, 30, 35곳」 중간)
rows = [
    ["파트너 × 매수자 (일일이 연결할 때 계약·실사 관계)", P * Bm, "612 × 30"],
    ["한 창구일 때 계약 수 (파트너 + 매수자)", P + Bm, ""],
    ["파트너 1곳이 더해질 때 새로 생기는 거래 가능 쌍", Bm, ""],
    ["매수자 1곳이 더해질 때 새로 생기는 거래 가능 쌍", P, "매수자가 희소한 쪽 → 매수자 1곳의 무게가 파트너 1곳의 약 20배"],
    ["은행 20곳이 판매자로만 올 때 더해지는 쌍", 20 * Bm, ""],
    ["은행 20곳이 판매자이자 매수자로 올 때 더해지는 쌍", 20 * Bm + 20 * P + 20 * 19, "은행끼리 사고파는 쌍 포함"],
]
w("ip1_23_network_pairs.csv", ["item", "value", "note"], rows)
