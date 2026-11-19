import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from smic_xlsx import SourceBook, load
from smic_charts_lib import *

HM = sys.argv[1]           # hmda/out dir
OUTX, OUTP = sys.argv[2], sys.argv[3]
M = 1_000_000

from smic_charts_lib import _plot_area, axis_unit, qn
def lift_units(p, top, left_unit, right_unit=""):
    """범례가 두 줄일 때: 플롯 위 여백을 top(cm)으로 늘리고 단위를 범례 아래·플롯 바로 위에 둔다"""
    l, t, r, b = p.inner
    p.units = ("", "")
    p.layout((l, top, r, b))
    yf = (top - 0.56) / p.h
    axs = _plot_area(p.ch).findall(qn("c:valAx"))
    axis_unit(axs[0], left_unit, 0.0, yf)
    if right_unit:
        from smic_charts_lib import tw
        axis_unit(axs[-1], right_unit, max(0.0, 1 - (tw(right_unit) + 0.05) / p.w), yf)
    p.units = (left_unit, right_unit)

def heloc(lei, y):
    return json.load(open(f"{HM}/sum_{lei}_{y}.json"))["heloc_amt"]

wb = SourceBook()
# ---------- 도표 1-1: 동사 1순위 거래액·비중 (회사 공시) ----------
src_q3 = "Figure 3Q25 실적 발표(SEC 8-K, 2025-11-13, Ex.99.1)"
src_q4 = "Figure 4Q25 실적 발표(SEC 8-K, 2026-02-26, Ex.99.1)"
src_q1 = "Figure 1Q26 실적 발표(SEC 8-K, 2026-05-12, Ex.99.1)"
wb.add("도표 1-1", "동사 1순위 HELOC 거래액 및 비중 추이", ["3Q24", "3Q25", "4Q25", "1Q26"],
       cols=[("Consumer Loan Marketplace 거래액", [1450 * M, 2469 * M, 2705 * M, 2902 * M], FMT["백만"], "hide"),
             ("1순위 비중(RHS)", [0.105, 0.17, 0.19, 0.20], FMT["%"]),
             ("1순위(LHS)", "={Consumer Loan Marketplace 거래액}*{1순위 비중(RHS)}", FMT["백만"]),
             ("2순위·기타(LHS)", "={Consumer Loan Marketplace 거래액}-{1순위(LHS)}", FMT["백만"]),
             ("1순위 배수(vs 3Q24)", "={1순위(LHS)}/{1순위(LHS)@-3}", FMT["배"], "hide")],
       row_sources=[src_q3 + ": 거래액 1,450백만 달러; 비중 = 3Q25 '650bps 증가해 17%'에서 역산(17%-6.5%p)",
                    src_q3 + ": 거래액 2,469백만 달러; 'First-lien volume as a percentage of Consumer Loan Marketplace volume increased 650 bps to 17%'",
                    src_q4 + ": 거래액 2,705백만 달러; 'First-lien volume as a percentage of Consumer Loan Marketplace volume increased to 19%'",
                    src_q1 + ": 거래액 2,902백만 달러; 'growing first lien volume to 20% of total production'"],
       source="Figure 실적 발표(SEC 8-K)", unit="(백만 달러) / (%)",
       note="1순위 거래액 = 거래액 x 공시 비중(반올림 %) -> 근사. 2순위·기타 = 2순위 HELOC + DSCR·개인대출 등. 4Q24~2Q25·2Q26 이후 분기 비중 미공시",
       check="1Q26 비중은 'total production' 대비로 표현 — 같은 기준으로 간주")

# ---------- 도표 1-2: 플랫폼 합류 파트너 HELOC 실행액 (HMDA) ----------
YRS = [2021, 2022, 2023, 2024, 2025]
L = {"HB": "5493001WHVQBGRSWEU75", "NAF": "549300E2UX99HKDBR481", "TLS": "254900DTLHVWQ7NP7R34",
     "CCM": "549300VZVN841I2ILS84", "GR": "549300U3721PJGQZYY68", "GRA": "54930001NSTOD85LT125",
     "MOV": "549300DD4R4SYK5RAQ92", "UHM": "549300RPOGWJRH63HS39", "FW": "549300MGPZBLQDIL7538",
     "SYN": "2549004FXPNKCLIN8250", "HP": "549300BRJZYHYKT4BJ84"}
v = {k: [heloc(l, y) for y in YRS] for k, l in L.items()}
hm_src = "FFIEC HMDA loan-level(실행분, open-end HELOC, 사업목적·역모기지 제외), 파트너 LEI별 합산 — data/figr_partner_heloc_hmda.xlsx"
wb.add("도표 1-2", "동사 플랫폼 합류 파트너 10곳의 HELOC 실행액 추이", [str(y) for y in YRS],
       cols=[("Homebridge", v["HB"], FMT["백만"]),
             ("New American", v["NAF"], FMT["백만"]),
             ("The Loan Store", v["TLS"], FMT["백만"]),
             ("CrossCountry", v["CCM"], FMT["백만"]),
             ("Guaranteed Rate", v["GR"], FMT["백만"], "hide"),
             ("Guaranteed Rate Affinity", v["GRA"], FMT["백만"], "hide"),
             ("Movement", v["MOV"], FMT["백만"], "hide"),
             ("Union Home", v["UHM"], FMT["백만"], "hide"),
             ("Fairway", v["FW"], FMT["백만"], "hide"),
             ("Synergy One", v["SYN"], FMT["백만"], "hide"),
             ("Home Point", v["HP"], FMT["백만"], "hide"),
             ("기타 6곳", "={Guaranteed Rate}+{Guaranteed Rate Affinity}+{Movement}+{Union Home}+{Fairway}+{Synergy One}+{Home Point}", FMT["백만"]),
             ("합계", "={Homebridge}+{New American}+{The Loan Store}+{CrossCountry}+{기타 6곳}", FMT["백만"], "hide"),
             ("배수(vs 3년 전)", "={합계}/{합계@-3}", FMT["배"], "hide")],
       source=hm_src, unit="(백만 달러)",
       note="파트너 = Figure 보도자료·DBRS 보고서로 확인된 HELOC 화이트라벨 파트너 중 2024년까지 합류, CMG(2018년부터 자체 HELOC) 제외. 기타 6곳 = Guaranteed Rate(+Affinity)·Movement·Union Home·Fairway·Synergy One·Home Point",
       check="HMDA 보고기준(직전 2년 각 200건) 미달 시 0 — 합류 초기 연도 과소집계(예: CrossCountry 2023~24)")
wb.save(OUTX)
bk = load(OUTX)

prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)",
            ["도표 1-1. 동사 1순위 HELOC 거래액 및 비중 추이 (기본)",
             "도표 1-2. 동사 플랫폼 합류 파트너 10곳의 HELOC 실행액 추이 (기본)"],
            checks=["도표 번호는 임시(1-1, 1-2) — 보고서 순서에 맞게 변경",
                    "도표 1-1 1순위 거래액은 공시 비중(반올림)으로 산출한 근사치"])

# ---- 도표 1-1 ----
sh = bk["도표 1-1"]
cv = chart_slide(prs, "도표 1-1", "동사 1순위 HELOC 거래액 및 비중 추이", "기본",
                 "Figure, SMIC 3팀",
                 notes="거래액(Consumer Loan Marketplace, 백만 달러): 3Q24 1,450 / 3Q25 2,469 / 4Q25 2,705 / 1Q26 2,902. "
                       "1순위 비중: 3Q24 10.5%(역산) / 17% / 19% / 20%. 1순위 = 거래액 x 비중. 출처: SEC 8-K Ex.99.1 각 분기 실적 발표",
                 note_lines=["1순위 비중은 회사 공시(Consumer Loan Marketplace 거래액 대비, 3Q24는 3Q25 발표의 '+650bp, 17%'로 역산), 1순위 거래액 = 거래액 × 비중(근사)",
                             "2순위·기타 = 2순위 HELOC 및 DSCR·개인대출 등. 4Q24~2Q25, 2Q26 이후는 분기 비중 미공시"])
p = add_chart(cv, "stack", sh.cats, sh.series("1순위(LHS)", "2순위·기타(LHS)", "1순위 비중(RHS)"), sheet=sh.name,
              fmt=sh.fmt("1순위(LHS)"), units=("(백만 달러)", ""), colors=[RED, LG, DG], head=0.15, gap=80)
p.rhs(["1순위 비중(RHS)"], fmt="0%", unit="(%)", colors=[DG])
lift_units(p, 1.42, "(백만 달러)", "(%)")
p.labels("1순위 비중(RHS)", points=[0, 1, 2, 3], fmt="0.0%", pos="t")
p.labels("1순위(LHS)", points=[1, 2, 3], fmt=FMT["백만"], color=WHT, pos="ctr")
cv.val(p.px0 + 0.15, p.y(3_500 * M), sh, "1순위 배수(vs 3Q24)", 3, pre="1순위 거래액 3Q24 대비 ", fmt='0.0"배"', color=RED)

# ---- 도표 1-2 ----
sh2 = bk["도표 1-2"]
cv2 = chart_slide(prs, "도표 1-2", "동사 플랫폼 합류 파트너 10곳의 HELOC 실행액 추이", "기본",
                  "FFIEC HMDA, Figure, DBRS, SMIC 3팀",
                  notes="HMDA loan-level 실행분 중 open-end HELOC(사업목적·역모기지 제외) 한도액 합산. 파트너별 원값은 원자료 엑셀 숨김 열 및 data/figr_partner_heloc_hmda.xlsx",
                  note_lines=["파트너 명의로 실행돼 HMDA에 보고된 HELOC 한도액. 공개 확인 파트너 중 '24년까지 합류한 10곳(자체 HELOC 보유 CMG 제외), 빨강 계열 = '22년 실적 0이던 8곳",
                              "HMDA는 직전 2년 각 200건 이상 기관만 보고 의무 → 합류 초기 연도는 0으로 잡힐 수 있음(과소집계)"])
names = ["Homebridge", "New American", "The Loan Store", "CrossCountry", "기타 6곳"]
p2 = add_chart(cv2, "stack", sh2.cats, sh2.series(*names), sheet=sh2.name, fmt=sh2.fmt("Homebridge"),
               units=("(백만 달러)", ""), colors=[DG, MG, RED, PINK2, PINK], head=0.28, gap=70)
lift_units(p2, 1.55, "(백만 달러)")
tot = sh2.cols["합계"]
for i in range(len(sh2.cats)):
    cv2.val(p2.x(i), p2.y(tot[i]) - 0.5, sh2, "합계", i, align="c", color=RED if i == 4 else DG)
x22, x25 = p2.x(1), p2.x(4)
cv2.arrow(x22, p2.y(tot[1]) - 0.5, x25 - 0.55, p2.y(tot[4]) - 0.2, color=RED)
cv2.val((x22 + x25) / 2 - 0.2, p2.y(tot[4]) - 0.05, sh2, "배수(vs 3년 전)", 4, pre="3년 만에 ", fmt='0.0"배"', color=RED, align="r")
prs.save(OUTP)
print("built")
