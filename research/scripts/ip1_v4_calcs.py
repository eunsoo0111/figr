"""투자포인트 1 v4 개요(IP1_outline_v4_2026-11-17.md)의 계산을 다시 만드는 스크립트.

실행: python3 -I research/scripts/ip1_v4_calcs.py research
  (인자는 research 폴더 경로. 결과 CSV는 research/data/ 에 씀)

v4의 질문은 「시장(TAM)이 평평해도 Figure가 크는가」예요. 그래서 계산도 그 질문에 맞췄어요.
1. 분기 KPI 표(4Q24~3Q26): Connect vs 비Connect, take rate, 조정 EBITDA, 파트너 수, 에코시스템 수수료 ÷ Connect
2. 2026년 「자연 실험」: 시장 지표 성장률 vs Figure 성장률, 성장 중 점유율이 설명하는 몫
3. 2024→25 성장 분해(시장 크기 / 비은행 몫 / Figure 몫)
4. 점유율 사다리: Figure 플랫폼 HELOC ÷ HMDA 2순위 전체
5. 시장이 2025년 크기에 멈춘다고 둘 때의 거래량 시나리오(비은행 몫 × 예금기관 침투율)
6. 대출사 마진 비교(3Q26): Figure vs Rocket·UWM·loanDepot
7. 대출기관 유형별 지도(HMDA 2025 한도·금액, 기관 수)

입력 숫자는 모두 아래 상수에 출처와 함께 적었어요. 회사 KPI는 8-K Ex.99.1(4Q25·1Q26·2Q26·3Q26) 원문,
시장 지표는 ICE Mortgage Monitor·TransUnion(National Mortgage News 보도)·NY Fed·KBRA·BofA, 경쟁사는 각 사 3Q26 실적 발표.
"""
import csv
import math
import os
import sys

ROOT = sys.argv[1] if len(sys.argv) > 1 else "research"
DATA = os.path.join(ROOT, "data")
HMDA = os.path.join(ROOT, "bank-channel", "data", "hmda_second_lien_by_lender_type.csv")


def write(name, header, rows):
    path = os.path.join(DATA, name)
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"  -> {path}")


def pct(x, d=1):
    return round(100 * x, d)


# ---------------------------------------------------------------- 1. 분기 KPI
# 출처: 8-K Ex.99.1 — 4Q25(2026-02-26, 4Q24 비교열 포함), 1Q26(2026-05-11, 1Q25 비교열), 2Q26(2026-08-13, 2Q25 비교열),
#       3Q26(2026-11-12, 3Q25 비교열). $M. 3Q26은 Kiavi(9월 $677M) 제외 standalone로 맞춤.
#   - 3Q26 standalone 거래량 = 5,796 - 677 = 5,119, standalone Connect = 72% x 5,119 (회사: "72% of standalone")
#   - 조정 EBITDA·조정 순매출 정의가 1Q26·2Q26에 바뀌었고 과거치를 다시 썼어요. 4Q24·4Q25는 옛 정의라 직접 비교는 조심
#   - 파트너 수: 4Q25 307, 1Q26 387(+80), 2Q26 489(+102), 3Q26 612(+123, Kiavi 브로커 포함 894). 2Q25 약 250은 IPO 무렵 회사 표현
Q = ["4Q24", "1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26", "3Q26"]
VOL = [1172, 1365, 1838, 2469, 2705, 2902, 4259, 5796 - 677]
CON = [8, 478, 767, 1131, 1466, 1612, 2773, round(0.72 * (5796 - 677))]
NET_TR = [3.4, 3.6, 4.0, 4.4, 3.8, 3.8, 3.6, 3.5]          # 3Q26 3.5%는 Kiavi 포함(standalone 3.7%, 콜)
GROSS_TR = ["", "", "", 5.2, "", "", "", 3.9]               # gross는 3Q26 보도자료에서 처음 공개(3Q25 비교치 포함)
ADJ_REV = [76.6, 87.0, 111.9, 155.1, 157.6, 166.8, 218.4, 283.3]
ADJ_EBITDA = [15.5, 28.3, 52.9, 85.4, 81.3, 82.7, 119.4, 163.6]
ECO_FEES = [8.5, 15.6, 28.1, 35.7, 41.4, 47.3, 72.9, 109.8]  # 에코시스템·기술 수수료
BOLT_FEES_3Q26 = 11.2                                        # 10-Q: BOLT JV 9월 한 달 에코시스템·기술 수수료
GOS = [24.3, 29.8, 36.3, 62.6, 48.1, 49.4, 57.6, 57.8]       # 대출 매각차익
PARTNERS = ["", "", "약 250", "", 307, 387, 489, 612]

print("[1] 분기 KPI")
rows = []
for i, q in enumerate(Q):
    non = VOL[i] - CON[i]
    eco = ECO_FEES[i] - (BOLT_FEES_3Q26 if q == "3Q26" else 0)
    rows.append([q, VOL[i], CON[i], non, pct(CON[i] / VOL[i], 0), NET_TR[i], GROSS_TR[i],
                 ADJ_REV[i], ADJ_EBITDA[i], pct(ADJ_EBITDA[i] / ADJ_REV[i]), ECO_FEES[i], GOS[i],
                 round(100 * eco / CON[i], 2) if CON[i] > 100 else "", PARTNERS[i]])
write("ip1_v4_quarterly_kpi.csv",
      ["quarter", "loan_marketplace_vol_$M_standalone", "connect_vol_$M", "non_connect_vol_$M", "connect_share_%",
       "net_take_rate_%", "gross_take_rate_%", "adj_net_revenue_$M", "adj_ebitda_$M", "adj_ebitda_margin_%",
       "eco_tech_fees_$M", "gain_on_sale_$M", "eco_fees_per_connect_vol_%_(3Q26_ex_BOLT)", "active_partners"],
      rows)
for r in rows:
    print("  ", r)
d_tot = VOL[-1] - VOL[3]
d_con = CON[-1] - CON[3]
print(f"  3Q25->3Q26 늘어난 거래량 ${d_tot:,}M 중 Connect ${d_con:,}M = {d_con/d_tot:.0%}")
d_tot2 = VOL[-1] - VOL[0]
d_con2 = CON[-1] - CON[0]
print(f"  4Q24->3Q26(7분기) 늘어난 거래량 ${d_tot2:,}M 중 Connect ${d_con2:,}M = {d_con2/d_tot2:.0%}")
print(f"  비Connect 거래량 범위 ${min(VOL[i]-CON[i] for i in range(len(Q))):,}~"
      f"${max(VOL[i]-CON[i] for i in range(len(Q))):,}M / 분기")

# 2026년 연간 standalone 추정: 9개월 실적 + 4Q 가이던스(Kiavi 포함 $6.2~6.7B)에서 Kiavi 한 분기(연 $8B 런레이트 ÷ 4)를 뺌
NINE_M_2026 = VOL[5] + VOL[6] + VOL[7]
KIAVI_Q = 8000 / 4
FY26_LO = NINE_M_2026 + 6200 - KIAVI_Q
FY26_HI = NINE_M_2026 + 6700 - KIAVI_Q
FY26_MID = (FY26_LO + FY26_HI) / 2
print(f"  2026E standalone ${FY26_LO/1000:.1f}~{FY26_HI/1000:.1f}B (중간 ${FY26_MID/1000:.1f}B), "
      f"2025 CLM $8.4B 대비 x{FY26_MID/8377:.2f}")

# ---------------------------------------------------------------- 2. 2026 자연 실험
# 시장 지표(같은 2Q26 기준으로 맞춤)
#   ICE Mortgage Monitor 2026-09: 2Q26 2순위(HELOC+홈에쿼티론) 인출 $29.5B, 2Q25 고점 대비 -3%. 1Q26 $25B(6월호)
#   TransUnion Q3 2026 CIIR(National Mortgage News 보도): 2Q26 홈에쿼티 실행 75.7만 건 +10.1% YoY, 그중 HELOC +20.3%
#   NY Fed HHDC 2026Q3: HELOC 잔액 $481B, 전년 대비 +14% (잔액이라 흐름과 다름)
# Figure: 2Q26 CLM $4,259M vs 2Q25 $1,838M (+132%), Connect $2,773M vs $767M (+262%)
print("[2] 2026 자연 실험 (2Q26 YoY)")
FIG_G = VOL[6] / VOL[2]
CON_G = CON[6] / CON[2]
markets = [("ICE 2순위 인출액($)", 0.97), ("TransUnion 홈에쿼티 실행 건수", 1.101), ("TransUnion HELOC 실행 건수", 1.203)]
rows = []
for name, g in markets:
    share = FIG_G / g
    rows.append([name, pct(g - 1), pct(FIG_G - 1), pct(share - 1), pct(math.log(share) / math.log(FIG_G), 0),
                 round((FIG_G - 1) / (g - 1), 1) if g > 1 else "시장 감소"])
write("ip1_v4_market_vs_figure_2026.csv",
      ["market_measure", "market_yoy_%", "figure_volume_yoy_%", "implied_share_gain_%",
       "share_explains_%_of_figure_growth(log)", "figure_growth_multiple_of_market"], rows)
for r in rows:
    print("  ", r)
print(f"  같은 분기 Connect +{CON_G-1:.0%}")
three_q = [("1Q26", VOL[5] / VOL[1], CON[5] / CON[1]), ("2Q26", FIG_G, CON_G), ("3Q26", VOL[7] / VOL[3], CON[7] / CON[3])]
for q, g, cg in three_q:
    print(f"  {q} standalone 거래량 {g-1:+.0%}, Connect {cg-1:+.0%}")

# ---------------------------------------------------------------- 3. 2024->25 분해
print("[3] 2024->25 성장 분해 (HMDA 2순위, Figure 플랫폼 HELOC)")
hm = list(csv.DictReader(open(HMDA)))


def tot(year, lender=None, product=None):
    return sum(float(r["limit_or_amount_$B"]) for r in hm if r["year"] == year
               and (lender is None or r["lender_type"] == lender)
               and (product is None or r["product"] == product))


FIG = {"2024": 5.1, "2025": 8.3}   # 10-K: 플랫폼 HELOC 2025 $8.3B(+62%)
t24, t25 = tot("2024"), tot("2025")
n24, n25 = tot("2024", "nonbank"), tot("2025", "nonbank")
m_tam = t25 / t24
m_nb = (n25 / t25) / (n24 / t24)
m_fig = (FIG["2025"] / n25) / (FIG["2024"] / n24)
total = FIG["2025"] / FIG["2024"]
L = math.log(total)
rows = [["시장 크기(2순위 전체)", round(m_tam, 3), pct(math.log(m_tam) / L, 0)],
        ["그중 비은행 몫", round(m_nb, 3), pct(math.log(m_nb) / L, 0)],
        ["비은행 안 Figure 몫", round(m_fig, 3), pct(math.log(m_fig) / L, 0)],
        ["곱(= Figure 성장)", round(m_tam * m_nb * m_fig, 3), 100]]
write("ip1_v4_growth_decomposition_2025.csv", ["factor", "multiplier", "share_of_growth_%(log)"], rows)
for r in rows:
    print("  ", r)

# ---------------------------------------------------------------- 4. 점유율 사다리
print("[4] 점유율 사다리 (Figure 플랫폼 ÷ HMDA 2순위 전체)")
rows = [["2024", 5.1, round(t24, 1), pct(5.1 / t24), round(n24, 1), pct(5.1 / n24)],
        ["2025", 8.3, round(t25, 1), pct(8.3 / t25), round(n25, 1), pct(8.3 / n25)]]
for label, g in (("2026E(시장 2025년과 같음)", 1.0), ("2026E(시장 +10%)", 1.10), ("2026E(시장 +20%)", 1.20)):
    rows.append([label, round(FY26_MID / 1000, 1), round(t25 * g, 1), pct(FY26_MID / 1000 / (t25 * g)), "", ""])
write("ip1_v4_share_ladder.csv",
      ["year", "figure_platform_$B", "second_lien_market_$B", "figure_share_%", "nonbank_second_lien_$B",
       "figure_share_of_nonbank_%"], rows)
for r in rows:
    print("  ", r)

# ---------------------------------------------------------------- 5. 시장 정체 시나리오
# 시장을 2025년 HMDA 크기(2순위 $172.5B)에 묶고, 비은행 몫(24.6%)도 묶은 가장 보수적인 틀.
# Figure 2026E 2순위 거래량 = 2026E standalone x 80% (1순위 HELOC 약 20%를 뺌, 2Q 콜 「1Q 수준 약 20%」)
print("[5] 시장이 2025년에 멈춰도: 비은행 몫 x 예금기관 침투율")
NB_POOL = n25
DEP_POOL = t25 - n25
FIRST_LIEN_SHARE = 0.20
fig26_2nd = FY26_MID / 1000 * (1 - FIRST_LIEN_SHARE)
print(f"  비은행 2순위 ${NB_POOL:.1f}B, 예금기관 2순위 ${DEP_POOL:.1f}B, "
      f"Figure 2026E 2순위 약 ${fig26_2nd:.1f}B = 비은행의 {fig26_2nd/NB_POOL:.0%}(전부 비은행이라 둘 때)")
rows = []
for s_nb in (round(fig26_2nd / NB_POOL, 2), 0.40, 0.50, 0.60):
    for p_dep in (0.0, 0.05, 0.10, 0.20):
        second = s_nb * NB_POOL + p_dep * DEP_POOL
        platform = second / (1 - FIRST_LIEN_SHARE)
        rows.append([pct(s_nb, 0), pct(p_dep, 0), round(second, 1), round(platform, 1),
                     round(platform / (FY26_MID / 1000), 2), pct(second / t25)])
write("ip1_v4_flat_tam_scenarios.csv",
      ["figure_share_of_nonbank_%", "depository_penetration_%", "figure_second_lien_$B",
       "figure_platform_$B(1st_lien_20%)", "x_vs_2026E", "share_of_total_second_lien_%"], rows)
for r in rows:
    if r[1] in (0.0, 5.0, 10.0, 20.0) and r[0] in (32.0, 40.0, 50.0):
        print("  ", r)

# 지금 파트너만으로 가능한 물량(v3 2-B 방식): 파트너 1순위 실행 x 2순위 부착률
ATTACH_LO = tot("2025", product="HELOC") / 2226   # HMDA 1순위 실행 $2,226B (연도·범위 확인 필요)
ATTACH_HI = t25 / 2226
for label, base in (("식별 파트너 1순위 $339B(Ottavi)", 339), ("회사 발표 파트너 $500B+", 500)):
    print(f"  {label}: ${base*ATTACH_LO:.0f}~{base*ATTACH_HI:.0f}B (부착률 {ATTACH_LO:.1%}~{ATTACH_HI:.1%})")

# ---------------------------------------------------------------- 6. 대출사 마진 비교
# 3Q26 실적 발표: Figure 조정 EBITDA $163.6M / 조정 순매출 $283.3M
#   Rocket 조정 EBITDA $748M / 조정 매출 $2.57B (서비싱·Mr. Cooper 포함)
#   UWM 조정 EBITDA $178.3M / 총매출 $811.1M
#   loanDepot 조정 EBITDA $17M / 조정 총매출 $302M
print("[6] 대출사 마진 (3Q26)")
rows = [["Figure", 283.3, 163.6], ["Rocket", 2570, 748], ["UWM", 811.1, 178.3], ["loanDepot", 302, 17]]
out = [[n, r, e, pct(e / r)] for n, r, e in rows]
write("ip1_v4_lender_margin_3q26.csv", ["company", "revenue_$M", "adj_ebitda_$M", "adj_ebitda_margin_%"], out)
for r in out:
    print("  ", r)

# ---------------------------------------------------------------- 7. 대출기관 유형별 지도
# HMDA 2025(한도·금액, $B) + 기관 수: FDIC RIS 2026Q2 은행 4,313곳, NCUA 2Q26 연방보험 신협 4,214곳
print("[7] 대출기관 유형별 지도 (HMDA 2025)")
rows = []
for lt, n in (("bank", 4313), ("credit_union", 4214), ("nonbank", "")):
    h = tot("2025", lt, "HELOC")
    c = tot("2025", lt, "closed_end_second")
    rows.append([lt, n, round(h, 1), round(c, 1), round(h + c, 1), pct((h + c) / t25)])
write("ip1_v4_lender_type_map.csv",
      ["lender_type", "institutions", "heloc_$B", "closed_end_second_$B", "second_lien_total_$B",
       "share_of_second_lien_%"], rows)
for r in rows:
    print("  ", r)
