"""투자포인트 1 v3 개요(IP1_outline_v3_2026-11-17.md)의 계산을 다시 만드는 스크립트.

실행: python3 -I research/scripts/ip1_v3_calcs.py research
  (인자는 research 폴더 경로. 결과 CSV는 research/data/ 에 씀)

계산 네 가지
1. take rate 분해: 3Q25·3Q26 8-K 숫자로 채널 단가(Connect / 비Connect)를 역산
2. 현금인출 재융자 vs 기존 모기지 + HELOC 손익분기 금리
3. HELOC 여력 지표: NY Fed HELOC 잔액 / 모기지 잔액 비율
4. TAM · SAM · SOM 사다리: HMDA 2순위 실행(기관 유형별) + Figure 플랫폼 HELOC
"""
import csv
import os
import sys

ROOT = sys.argv[1] if len(sys.argv) > 1 else "research"
DATA = os.path.join(ROOT, "data")
HMDA = os.path.join(ROOT, "bank-channel", "data", "hmda_second_lien_by_lender_type.csv")
NYFED = os.path.join(DATA, "nyfed_hhdc_heloc_2003q1_2026q3.csv")


def write(name, header, rows):
    path = os.path.join(DATA, name)
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"  -> {path}")


# ---------------------------------------------------------------- 1. take rate
# 출처: 3Q26 8-K Ex.99.1 (2026-11-12) Selected Metrics, 3Q26 콜(Kiavi·standalone 설명)
VOL_3Q25, CON_3Q25, GROSS_3Q25 = 2469.0, 1131.0, 0.052
VOL_3Q26, CON_3Q26, GROSS_3Q26 = 5796.0, 4259.0, 0.039
KIAVI_3Q26 = 677.0          # 9월 한 달 Kiavi 거래량
STANDALONE_CONNECT_SHARE = 0.72  # "72% of standalone volume"
NET_3Q25, NET_3Q26 = 0.044, 0.035


def decompose(gross_3q25, kiavi_rate):
    standalone = VOL_3Q26 - KIAVI_3Q26
    con26 = STANDALONE_CONNECT_SHARE * standalone
    non26 = standalone - con26
    non25 = VOL_3Q25 - CON_3Q25
    num25 = gross_3q25 * VOL_3Q25
    num26_total = GROSS_3Q26 * VOL_3Q26
    num26 = num26_total - kiavi_rate * KIAVI_3Q26     # standalone 분자
    # 두 분기 모두 채널 단가가 같다고 두고 2원 1차 연립방정식을 푼다
    # CON_3Q25*c + non25*b = num25 ;  con26*c + non26*b = num26
    det = CON_3Q25 * non26 - non25 * con26
    c = (num25 * non26 - non25 * num26) / det
    b = (CON_3Q25 * num26 - num25 * con26) / det
    return dict(standalone=standalone, con26=con26, non26=non26, non25=non25,
                num25=num25, num26=num26, c=c, b=b)


print("[1] take rate 분해")
rows = []
for g in (0.0515, 0.052, 0.0525):
    for k in (0.015, 0.0175, 0.02):
        d = decompose(g, k)
        rows.append([round(g * 100, 2), round(k * 100, 2), round(d["c"] * 100, 2), round(d["b"] * 100, 2),
                     round(d["num26"] / d["standalone"] * 100, 2)])
write("ip1_v3_take_rate_sensitivity.csv",
      ["gross_3q25_%", "kiavi_rtl_rate_%", "implied_connect_rate_%", "implied_nonconnect_rate_%",
       "standalone_gross_3q26_%"], rows)
base = decompose(0.052, 0.0175)
c, b = base["c"], base["b"]
rev = [
    ["3Q25", CON_3Q25, base["non25"], CON_3Q25 * c, base["non25"] * b],
    ["3Q26 standalone", base["con26"], base["non26"], base["con26"] * c, base["non26"] * b],
]
write("ip1_v3_take_rate_decomposition.csv",
      ["period", "connect_vol_$M", "nonconnect_vol_$M", "connect_rev_$M", "nonconnect_rev_$M"],
      [[r[0]] + [round(x, 1) for x in r[1:]] for r in rev])
d_con = rev[1][3] - rev[0][3]
d_non = rev[1][4] - rev[0][4]
print(f"  Connect 단가 {c*100:.2f}%, 비Connect {b*100:.2f}% (기준 가정)")
print(f"  standalone Connect ${base['con26']:,.0f}M (YoY {base['con26']/CON_3Q25-1:+.0%}), "
      f"Kiavi 몫 Connect ${CON_3Q26-base['con26']:,.0f}M")
print(f"  비Connect ${base['non25']:,.0f}M -> ${base['non26']:,.0f}M ({base['non26']/base['non25']-1:+.1%})")
print(f"  분자 증가분 중 Connect 몫 {d_con/(d_con+d_non):.0%} (+${d_con:.0f}M vs +${d_non:.0f}M)")
mk25 = (GROSS_3Q25 - NET_3Q25) * VOL_3Q25
mk26 = (GROSS_3Q26 - NET_3Q26) * VOL_3Q26
print(f"  gross-net 차이로 본 마케팅비 ${mk25:.0f}M(3Q25) / ${mk26:.0f}M(3Q26), "
      f"비Connect 거래량 대비 {mk25/base['non25']:.1%} / {mk26/(VOL_3Q26-CON_3Q26):.1%}")
# 수렴: Connect 비중이 오를 때 standalone HELOC 혼합 단가
conv = []
for share in (0.46, 0.65, 0.72, 0.80, 0.85, 0.90):
    conv.append([int(share * 100), round((share * c + (1 - share) * b) * 100, 2)])
write("ip1_v3_take_rate_convergence.csv", ["connect_share_%", "blended_gross_take_rate_%"], conv)

# ---------------------------------------------------------------- 2. 손익분기
print("[2] 현금인출 재융자 vs HELOC 손익분기")
rows = []
for r1 in (0.0275, 0.03, 0.035, 0.04):
    for need in (50_000, 90_000):
        for h in (0.075, 0.091):
            bal = 300_000
            be = (bal * r1 + need * h) / (bal + need)
            rows.append([round(r1 * 100, 2), bal, need, round(h * 100, 2), round(be * 100, 2)])
write("ip1_v3_cashout_breakeven.csv",
      ["existing_rate_%", "existing_balance_$", "cash_need_$", "heloc_rate_%", "breakeven_new_mortgage_rate_%"],
      rows)
for r in rows:
    if r[0] == 3.5 and r[2] == 90_000:
        print(f"  기존 3.5% $300k + 필요 ${r[2]:,} @HELOC {r[3]}% -> 재융자 금리가 {r[4]}% 밑이어야 재융자가 유리")

# ---------------------------------------------------------------- 3. 여력 지표
print("[3] HELOC / 모기지 잔액 비율 (NY Fed)")
ny = list(csv.DictReader(open(NYFED)))
by_q = {r["quarter"]: r for r in ny}
last = ny[-1]
mort = float(last["mortgage_tn"])
heloc = float(last["heloc_tn"])
peak = max(ny, key=lambda r: float(r["heloc_tn"]))
rows = []
for q in ("2006Q4", peak["quarter"], "2012Q4", "2019Q4", "2022Q1", last["quarter"]):
    r = by_q[q]
    rows.append([q, r["mortgage_tn"], r["heloc_tn"], r["heloc_pct_of_mortgage"]])
for label, ratio in (("2019Q4 비율", float(by_q["2019Q4"]["heloc_pct_of_mortgage"]) / 100),
                     ("5%", 0.05), ("2009Q1 비율", float(peak["heloc_pct_of_mortgage"]) / 100)):
    implied = ratio * mort * 1000
    rows.append([f"지금 모기지 잔액에 {label} 적용", round(mort, 3), round(implied / 1000, 3),
                 f"{implied:,.0f}B ({implied / (heloc * 1000) - 1:+.0%})"])
write("ip1_v3_heloc_headroom.csv", ["point", "mortgage_tn", "heloc_tn", "heloc_pct_or_note"], rows)
print(f"  지금 {last['quarter']} HELOC ${heloc*1000:.0f}B = 모기지의 {last['heloc_pct_of_mortgage']}%, "
      f"정점 {peak['quarter']} ${float(peak['heloc_tn'])*1000:.0f}B({peak['heloc_pct_of_mortgage']}%)의 "
      f"{heloc/float(peak['heloc_tn']):.0%}")

# ---------------------------------------------------------------- 4. 사다리
print("[4] TAM · SAM · SOM (HMDA 2순위, 한도·금액 기준)")
hm = list(csv.DictReader(open(HMDA)))
FIG_PLATFORM_HELOC = {"2024": 5.1, "2025": 8.3}  # 10-K: 2025 $8.3B(+62%)
rows = []
for y in ("2023", "2024", "2025"):
    tam = sum(float(r["limit_or_amount_$B"]) for r in hm if r["year"] == y)
    heloc_all = sum(float(r["limit_or_amount_$B"]) for r in hm if r["year"] == y and r["product"] == "HELOC")
    sam = sum(float(r["limit_or_amount_$B"]) for r in hm if r["year"] == y and r["lender_type"] == "nonbank")
    sam_heloc = sum(float(r["limit_or_amount_$B"]) for r in hm
                    if r["year"] == y and r["lender_type"] == "nonbank" and r["product"] == "HELOC")
    fig = FIG_PLATFORM_HELOC.get(y)
    rows.append([y, round(tam, 1), round(heloc_all, 1), round(sam, 1), round(sam_heloc, 1),
                 round(100 * sam / tam, 1), fig if fig else "",
                 round(100 * fig / sam, 1) if fig else ""])
write("ip1_v3_tam_sam_som.csv",
      ["year", "tam_2nd_lien_all_$B", "heloc_all_$B", "sam_nonbank_2nd_lien_$B", "nonbank_heloc_$B",
       "sam_share_of_tam_%", "figure_platform_heloc_$B", "figure_share_of_sam_%"], rows)
for r in rows:
    print("  ", r)
