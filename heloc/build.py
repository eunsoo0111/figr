import json, sys
sys.path.insert(0, ".")
from smic_xlsx import SourceBook, load
from smic_charts_lib import *

d = json.load(open("raw/ddp/sel.json"))
M = 1_000_000
dates = [f"{y}-12-31" for y in range(1966, 2026)] + ["2026-06-30"]
cats = [str(y) for y in range(1966, 2026)] + ["26.2Q"]
axis = [("'" + str(y)[2:]) if y % 10 == 0 else "" for y in range(1966, 2026)] + ["26.2Q"]
def col(code):
    out = []
    for t in dates:
        v = d[code].get(t)
        out.append(None if v in (None, "", "ND") else int(round(float(v) * M)))
    return out
SRC = "Federal Reserve Z.1(2026.09.11, 26.2Q) L.218·L.127·L.124·L.125"
S = SourceBook()
cols = [
    ("1~4가구 모기지 전체", col("893065105"), FMT["십억"], "hide"),
    ("홈에퀴티론 전체", col("893065125"), FMT["십억"], "hide"),
    ("HELOC 전체", col("893065215"), FMT["십억"], "hide"),
    ("Agency·GSE 풀", col("413065105"), FMT["십억"], "hide"),
    ("GSE 보유 증권화분", col("403065195"), FMT["십억"], "hide"),
    ("ABS 1~4가구", col("673065105"), FMT["십억"], "hide"),
    ("ABS 홈에퀴티론", col("673065123"), FMT["십억"], "hide"),
    ("ABS HELOC", col("673065203"), FMT["십억"], "hide"),
    ("1순위(LHS)",
     "=({Agency·GSE 풀}+{GSE 보유 증권화분}+{ABS 1~4가구}-{ABS 홈에퀴티론})/({1~4가구 모기지 전체}-{홈에퀴티론 전체})", FMT["%"]),
    ("HELOC(RHS)", "={ABS HELOC}/{HELOC 전체}", FMT["%"]),
    ("2순위 전체 증권화율", "={ABS 홈에퀴티론}/{홈에퀴티론 전체}", FMT["%"], "hide"),
]
ws = S.add("도표 1-1", "1순위 모기지 vs HELOC 증권화율 추이", cats, cols,
           labels=[("축 레이블", axis)], source=SRC + " — 연말 잔액(26.2Q는 6월 말)",
           unit="(%) / (%)",
           note="1순위 = (Agency·GSE 풀 + GSE 보유 증권화분 + ABS 1~4가구 − ABS 홈에퀴티론) ÷ (1~4가구 전체 − 홈에퀴티론); "
                "HELOC = ABS 발행기구 보유 HELOC ÷ HELOC 전체. 홈에퀴티론은 1990년부터 집계(이전은 0으로 처리)",
           check="Z.1 계열코드: FL893065105·125·215, FL413065105, FL403065195, FL673065105·123·203; "
                 "기존 도표의 'HELOC 26.3Q 4.5%'와 출처 상이 — Z.1 기준 26.2Q 1.5%")
# 홈에퀴티·HELOC 집계 전(1966~1989)은 HELOC·2순위 비율을 비운다(0으로 나누기 방지)
hdr = [c.value for c in ws[1]]
for name in ("HELOC(RHS)", "2순위 전체 증권화율"):
    j = hdr.index(name) + 1
    for r in range(2, 2 + cats.index("1990")):
        ws.cell(r, j).value = None
S.save("SMIC_도표_HELOC_원자료.xlsx")

bk = load("SMIC_도표_HELOC_원자료.xlsx")
sh = bk["도표 1-1"]
n = len(cats)
i70, i71, i06, last = cats.index("1970"), cats.index("1971"), cats.index("2006"), n - 1

prs = new_deck()
intro_slide(prs, "HELOC 증권화", ["도표 1-1. 1순위 모기지 vs HELOC 증권화율 추이"],
            checks=["Z.1 26.3Q(2026.12 발표) 나오면 갱신", "기존 'HELOC 26.3Q 4.5%'와 출처·정의 확인"])
cv = chart_slide(prs, "도표 1-1", "1순위 모기지 vs HELOC 증권화율 추이", "기본",
                 "Federal Reserve Z.1, SMIC 3팀",
                 notes="1순위 = (Agency·GSE 풀+GSE 보유 증권화분+ABS 1~4가구−ABS 홈에퀴티론)/(1~4가구 전체−홈에퀴티론)\n"
                       "HELOC = ABS 발행기구 보유 HELOC(FL673065203)/HELOC 전체(FL893065215)\n"
                       "연말 잔액 기준, 26.2Q는 6월 말. Z.1 2026.09.11 발표",
                 note_lines=["1순위 = 증권화된 1~4가구 1순위 모기지(Agency·GSE·민간 ABS) ÷ 1~4가구 모기지(홈에퀴티론 제외), 연말 기준",
                             "HELOC = ABS 발행기구 보유 HELOC ÷ HELOC 잔액. HELOC은 1990년부터 집계"])
p = add_chart(cv, "line", sh.axis, sh.series("1순위(LHS)", "HELOC(RHS)"), sheet=sh.name,
              fmt="0%", units=("(%)", ""), colors=[MG, RED], zero=True, yscale=(0, 1.0, 0.2))
p.rhs(["HELOC(RHS)"], fmt="0.0%", unit="(%)", colors=[RED], yscale=(0, 0.025, 0.005))
# 표식 없는 선(61개 점)
for ser in p.ch._chartSpace.iter(qn("c:ser")):
    mk = ser.find(qn("c:marker"))
    if mk is not None:
        for e in list(mk):
            mk.remove(e)
        mk.append(etree.fromstring(f'<c:symbol xmlns:c="{C}" val="none"/>'))
# 끝값 레이블
L, R = "1순위(LHS)", "HELOC(RHS)"
cv.val(p.x(last) - 0.05, p.y(sh.cols[L][last]) - 0.42, sh, L, last, fmt="0.0%", color=MG, align="r")
cv.val(p.x(last) - 0.05, p.y(sh.cols[R][last], rhs=True) + 0.04, sh, R, last, fmt="0.0%", color=RED, align="r")
cv.val(p.x(i06) + 0.12, p.y(sh.cols[R][i06], rhs=True) - 0.12, sh, R, i06, pre="'06 ", fmt="0.0%", color=RED)
# 1순위 1970~71년 위치 표시 + 설명
cv.diamond((p.x(i70) + p.x(i71)) / 2, p.y((sh.cols[L][i70] + sh.cols[L][i71]) / 2), 0.16)
tx = p.x(i70) - 0.15
cv.val(tx, p.top + 0.05, sh, R, last, pre="HELOC 증권화율(26.2Q) ", fmt="0.0%", color=RED)
cv.val(tx, p.top + 0.38, sh, L, i70, pre="= 1순위의 1970~71년 수준(", fmt="0.0%", color=RED)
cv.val(tx + tw("= 1순위의 1970~71년 수준(" + fmt_value(sh.cols[L][i70], "0.0%")) + 0.02, p.top + 0.38,
       sh, L, i71, pre="~", post=")", fmt="0.0%", color=RED)
cv.line((p.x(i70) + p.x(i71)) / 2, p.top + 0.75, (p.x(i70) + p.x(i71)) / 2,
        p.y((sh.cols[L][i70] + sh.cols[L][i71]) / 2) - 0.12, RED, 0.75, dash=True)
prs.save("SMIC_도표_HELOC.pptx")
print("rates:", {c: (round(sh.cols[L][cats.index(c)], 4), sh.cols[R][cats.index(c)]) for c in ("1970", "1971", "2006", "2022", "2025", "26.2Q")})
