"""IP1 3-1 계산: take rate 하락 = 체질 변화(Connect) 를 8개 분기 숫자로.

입력: research/data/ip1_31_quarterly_revenue_lines.csv (8-K Ex.99.1 손익 줄, 최신 재작성치)
출력: research/data/ip1_31_*.csv

실행: python3 -I research/scripts/ip1_31_calcs.py research
"""
import csv
import os
import sys

ROOT = sys.argv[1] if len(sys.argv) > 1 else "research"
DATA = os.path.join(ROOT, "data")


def f(x):
    return float(x) if x not in ("", None) else 0.0


rows = list(csv.DictReader(open(os.path.join(DATA, "ip1_31_quarterly_revenue_lines.csv"), encoding="utf-8")))

out = []
for r in rows:
    q = r["quarter"]
    vol = f(r["loan_mkt_vol_$M"])
    kv = f(r["kiavi_vol_$M"])
    con = f(r["connect_vol_$M"])
    kcon = f(r["kiavi_connect_vol_$M"])
    eco = f(r["eco_tech_fees_$K"]) / 1000
    orig = f(r["origination_fees_$K"]) / 1000
    gos = f(r["gain_on_sale_$K"]) / 1000
    sva = f(r["gain_on_servicing_asset_$K"]) / 1000
    msr = f(r["msr_valuation_adj_$K"]) / 1000
    rev = f(r["total_net_revenue_$K"]) / 1000
    adj = f(r["adj_net_revenue_$K"]) / 1000
    sm = f(r["sales_marketing_$K"]) / 1000
    tech = f(r["tech_product_$K"]) / 1000
    ops = f(r["ops_processing_$K"]) / 1000
    ebitda = f(r["adj_ebitda_$K"]) / 1000
    lhfs = f(r["loans_held_for_sale_end_$K"]) / 1000
    kfee = f(r["kiavi_eco_fee_$K"]) / 1000
    net_tr = f(r["net_take_rate_reported_%"])

    num = eco + orig + gos + sva + msr  # gross take rate 분자(회사 정의)
    gross_tr = num / vol * 100
    mkt_implied = (gross_tr - net_tr) / 100 * vol  # gross−net 차이로 본 마케팅비(반올림 오차 ±0.05%p)

    # standalone(Kiavi 제외): Kiavi 분자는 BOLT JV 에코시스템 수수료로 근사
    s_vol = vol - kv
    s_con = con - kcon
    s_non = s_vol - s_con
    s_num = num - kfee
    s_gross = s_num / s_vol * 100

    out.append(dict(
        quarter=q, vol_standalone=s_vol, connect_standalone=s_con, nonconnect=s_non,
        connect_share=round(s_con / s_vol * 100, 1),
        gross_numerator=round(s_num, 1), gross_take_rate=round(s_gross, 2),
        net_take_rate_reported=net_tr, marketing_implied=round(mkt_implied, 1),
        eco_share_of_rev=round(eco / rev * 100, 1), gos_share_of_rev=round(gos / rev * 100, 1),
        eco_minus_gos=round(eco - gos, 1),
        lhfs_end=round(lhfs, 0), lhfs_per_vol=round(lhfs / s_vol, 3), lhfs_per_nonconnect=round(lhfs / s_non, 3),
        tech_per_vol_bp=round(tech / vol * 1e4, 1), ops_per_vol_bp=round(ops / vol * 1e4, 1),
        adj_rev=adj, adj_ebitda=ebitda, adj_ebitda_margin=round(ebitda / adj * 100, 1),
        gos_per_nonconnect=round(gos / s_non * 100, 2),
        eco_per_connect=round((eco - kfee) / s_con * 100, 2) if s_con > 50 else "",
    ))

with open(os.path.join(DATA, "ip1_31_quarterly_derived.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(out[0].keys()))
    w.writeheader()
    w.writerows(out)


# ---- 1) 채널 단가 회귀: 분자 = a·Connect + b·비Connect (절편 없음) ----
def ols2(xs1, xs2, ys):
    s11 = sum(a * a for a in xs1)
    s22 = sum(b * b for b in xs2)
    s12 = sum(a * b for a, b in zip(xs1, xs2))
    s1y = sum(a * y for a, y in zip(xs1, ys))
    s2y = sum(b * y for b, y in zip(xs2, ys))
    det = s11 * s22 - s12 * s12
    a = (s1y * s22 - s2y * s12) / det
    b = (s2y * s11 - s1y * s12) / det
    pred = [a * x1 + b * x2 for x1, x2 in zip(xs1, xs2)]
    ybar = sum(ys) / len(ys)
    ss_res = sum((y - p) ** 2 for y, p in zip(ys, pred))
    ss_tot = sum((y - ybar) ** 2 for y in ys)
    return a, b, 1 - ss_res / ss_tot, pred


def fit(label, sel):
    sub = [o for o in out if o["quarter"] in sel]
    a, b, r2, pred = ols2([o["connect_standalone"] for o in sub], [o["nonconnect"] for o in sub],
                          [o["gross_numerator"] for o in sub])
    return dict(sample=label, n=len(sub), connect_rate_pct=round(a * 100, 2),
                nonconnect_rate_pct=round(b * 100, 2), r2=round(r2, 3))


Q = [o["quarter"] for o in out]
fits = [
    fit("8개 분기 전체(4Q24~3Q26)", Q),
    fit("7개 분기(4Q24~2Q26, 3Q26 예측용)", Q[:-1]),
    fit("3Q25 제외 7개 분기(매각차익 이례 분기 제외)", [q for q in Q if q != "3Q25"]),
    fit("최근 5개 분기(3Q25~3Q26)", Q[3:]),
]
# 3Q26 표본 밖 예측
f7 = fits[1]
o3 = out[-1]
pred_num = f7["connect_rate_pct"] / 100 * o3["connect_standalone"] + f7["nonconnect_rate_pct"] / 100 * o3["nonconnect"]
pred_tr = pred_num / o3["vol_standalone"] * 100
fits.append(dict(sample="3Q26 표본 밖 예측(7개 분기 계수)", n=1,
                 connect_rate_pct=f"예측 분자 {pred_num:.1f}", nonconnect_rate_pct=f"실제 분자 {o3['gross_numerator']:.1f}",
                 r2=f"예측 {pred_tr:.2f}% vs 실제 {o3['gross_take_rate']:.2f}%"))

with open(os.path.join(DATA, "ip1_31_channel_rate_regression.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=list(fits[0].keys()))
    w.writeheader()
    w.writerows(fits)

# ---- 2) 3Q25 → 3Q26 take rate 변화 분해(믹스 vs 단가) ----
a = fits[0]["connect_rate_pct"]
b = fits[0]["nonconnect_rate_pct"]
q0 = next(o for o in out if o["quarter"] == "3Q25")
q1 = out[-1]
s0 = q0["connect_share"] / 100
s1 = q1["connect_share"] / 100
tr0 = q0["gross_take_rate"]
tr1 = q1["gross_take_rate"]
mix = (s1 - s0) * (a - b)  # 단가 고정, 비중만 바뀐 효과
total = tr1 - tr0
model0 = a * s0 + b * (1 - s0)  # 3Q25 를 두 단가 식으로 본 값
model1 = a * s1 + b * (1 - s1)
peak = -(tr0 - model0)  # 3Q25 가 식보다 높았던 만큼이 되돌아온 몫
resid = tr1 - model1  # 3Q26 실제 − 식 = 단가 변화
decomp = [
    dict(item="3Q25 gross take rate(standalone)", pp=round(tr0, 2)),
    dict(item="3Q26 gross take rate(standalone)", pp=round(tr1, 2)),
    dict(item="변화 합계", pp=round(total, 2)),
    dict(item="① 꼭지점 되돌림(3Q25 실제 − 식)", pp=round(peak, 2)),
    dict(item="② 믹스 효과(Connect 비중 46%→72%, 8개 분기 회귀 단가 고정)", pp=round(mix, 2)),
    dict(item="③ 단가 변화(3Q26 실제 − 식)", pp=round(resid, 2)),
    dict(item="① 몫 %", pp=round(peak / total * 100, 0)),
    dict(item="② 몫 %", pp=round(mix / total * 100, 0)),
]
with open(os.path.join(DATA, "ip1_31_take_rate_mix_decomposition.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=["item", "pp"])
    w.writeheader()
    w.writerows(decomp)

# ---- 3) 증분으로 보기 ----
first, last = out[0], out[-1]
yago = q0
inc = [
    dict(item="거래량(standalone) 배수 4Q24→3Q26", val=round(last["vol_standalone"] / first["vol_standalone"], 2)),
    dict(item="판매대기 대출(분기말) 배수 4Q24→3Q26", val=round(last["lhfs_end"] / first["lhfs_end"], 2)),
    dict(item="거래량 1달러당 판매대기 대출 4Q24", val=first["lhfs_per_vol"]),
    dict(item="거래량 1달러당 판매대기 대출 3Q26", val=last["lhfs_per_vol"]),
    dict(item="증분 조정 EBITDA 마진 % 3Q25→3Q26", val=round((last["adj_ebitda"] - yago["adj_ebitda"]) / (last["adj_rev"] - yago["adj_rev"]) * 100, 1)),
    dict(item="증분 조정 EBITDA 마진 % 4Q24→3Q26", val=round((last["adj_ebitda"] - first["adj_ebitda"]) / (last["adj_rev"] - first["adj_rev"]) * 100, 1)),
    dict(item="1년 증가 거래량 중 Connect 몫 %", val=round((last["connect_standalone"] - yago["connect_standalone"]) / (last["vol_standalone"] - yago["vol_standalone"]) * 100, 1)),
    dict(item="1년 증가 take rate 분자 중 에코시스템·기술 수수료 증가 몫 %(Kiavi 제외)",
         val=round(((109.783 - 11.2) - 35.691) / (last["gross_numerator"] - yago["gross_numerator"]) * 100, 1)),
    dict(item="기술·제품 비용 ÷ 거래량 bp 4Q24", val=first["tech_per_vol_bp"]),
    dict(item="기술·제품 비용 ÷ 거래량 bp 3Q26", val=last["tech_per_vol_bp"]),
    dict(item="운영·처리 비용 ÷ 거래량 bp 4Q24", val=first["ops_per_vol_bp"]),
    dict(item="운영·처리 비용 ÷ 거래량 bp 3Q26", val=last["ops_per_vol_bp"]),
]
with open(os.path.join(DATA, "ip1_31_incremental.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.DictWriter(fh, fieldnames=["item", "val"])
    w.writeheader()
    w.writerows(inc)

for name in ["ip1_31_quarterly_derived.csv", "ip1_31_channel_rate_regression.csv",
             "ip1_31_take_rate_mix_decomposition.csv", "ip1_31_incremental.csv"]:
    print("====", name)
    print(open(os.path.join(DATA, name), encoding="utf-8").read())
