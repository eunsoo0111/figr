# 도표 3-1 사라졌다 돌아온 HELOC 유통시장 — 원자료 엑셀 → PPT
import sys
sys.path.insert(0, ".")
from smic_xlsx import SourceBook, load
from smic_charts_lib import *
from data_3_1 import YEARS, MKT, FIG, MKT_SRC, NOTE, EVENTS, FIG_SRC, ROW_FORMULA, ROW, EVENT_SRC

NO, TITLE = "도표 3-1", "사라졌다 돌아온 HELOC 유통시장: 미국 2순위 증권화 발행액"
XLSX, PPTX = "SMIC_도표3-1_원자료.xlsx", "SMIC_도표3-1.pptx"

# ── 1) 원자료 엑셀 ──
wb = SourceBook()
n = len(YEARS)
ev = [EVENTS.get(y, "") for y in YEARS]
ws = wb.add(NO, TITLE, YEARS,
            cols=[("시장 합계", MKT, FMT["십억"], "hide"),
                  ("동사(Figure)", FIG, FMT["십억"]),
                  ("기타 발행사", '=IF({시장 합계}="","",{시장 합계}-{동사(Figure)})', FMT["십억"])],
            labels=[("사건", ev), ("사건 출처", [EVENT_SRC.get(y, "") for y in YEARS]),
                    ("동사 출처", [FIG_SRC.get(y, "") for y in YEARS])],
            row_sources=[MKT_SRC[y] for y in YEARS],
            source="IMF, S&P, HEL News, J.P. Morgan, BofA, DBRS, Figure 공시(S-1·10-K·10-Q) — 행별 출처 열 참고",
            unit="(십억 달러)", size="기본", note=NOTE.replace(" / ", " "),
            check="06·07년 시장 연간 발행액 미확보(n/a) — IMF 'Mortgage & Asset Securities Issuance' 등 원자료 필요; "
                  "19~21년은 확인된 딜 합계(하한); 동사 25년·26년은 10-K·10-Q 증가액으로 SMIC 계산 — 딜별 합계로 확인 권장")
for (col, y), f in ROW_FORMULA.items():     # 역산·합산·누적 차이는 수식으로 남긴다
    c = ws[f"{col}{ROW[y]}"]
    c.value = f
wb.save(XLSX)
bk = load(XLSX)
sh = bk[NO]

# ── 2) PPT ──
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)", [f"{NO}. {TITLE}"],
            checks=["06·07년 시장 발행액 원자료 미확보(n/a) — 받으면 엑셀·차트 데이터에 입력", "연도별 집계 기관이 다름(주 참고)", "26E 시장은 BofA 전망, 26E 동사는 26.3Q 누적 실적"])
cv = chart_slide(prs, NO, TITLE, "기본", "Inside Mortgage Finance, S&P, HEL News, J.P. Morgan, BofA, DBRS, Figure, SMIC 3팀",
                 notes="원 수치·출처는 엑셀 '도표 3-1' 시트 참고", note_lines=NOTE.split(" / "))
left = tw("80") + 0.18
p = add_chart(cv, "stack", sh.cats, sh.series("동사(Figure)", "기타 발행사"), sheet=sh.name,
              fmt=sh.fmt("동사(Figure)"), units=("(십억 달러)", ""), colors=[RED, LG], gap=35,
              yscale=(0, 80_000_000_000, 20_000_000_000), inner=(left, 0.62, 0.12, 0.5))
# 가로축: 23개 연도가 막대 폭(0.37cm)보다 넓어 PowerPoint·LibreOffice가 회전·줄바꿈한다(45° 금지).
# → 축 글자는 숨기고 2년마다 같은 8pt 글자로 직접 쓴다(위치는 고정 플롯 좌표 p.x).
from smic_charts_lib import _plot_area
_cat = _plot_area(p.ch).find(qn("c:catAx"))
_cat.find(qn("c:tickLblPos")).set("val", "none")
for i, y in enumerate(sh.cats):
    if i % 2 == 0:
        w = tw(y) + 0.1
        cv.label(min(p.x(i), cv.w - w / 2), p.bottom + 0.11, y, DG, align="c")

# 전망 구간 음영 + 라벨
i26 = n - 1
x0 = (p.x(i26 - 1) + p.x(i26)) / 2
cv.band(x0, p.px0 + p.pw, p.top, p.bottom, BG, 0.6)
cv.label(p.px0 + p.pw, p.top + 0.02, "전망", MG, align="r")

# 합계 레이블(핵심 연도만) — 25·26E는 빨강
tot = sh.cols["시장 합계"]
for y in ("'05", "'22", "'25", "'26E"):
    i = YEARS.index(y)
    w = tw(fmt_value(tot[i], sh.fmt("시장 합계"))) + 0.1
    xc = min(p.x(i), cv.w - w / 2 - 0.02)          # 가이드 밖으로 나가지 않게
    cv.val(xc, p.y(tot[i]) - 0.36, sh, "시장 합계", i, color=RED if y in ("'25", "'26E") else DG, align="c")

# 사건 표시(글자는 엑셀 '사건' 열에서) — 점선 지시선은 막대 꼭대기까지
def event(y, ty, align, color, x_line=None, y_end=None):
    i = YEARS.index(y)
    xl = p.x(i) if x_line is None else x_line
    tx = xl + (0.08 if align == "l" else -0.08 if align == "r" else 0)
    cv.val(tx, ty, sh, "사건", i, color=color, align=align)
    ye = (p.y(tot[i]) if tot[i] else p.bottom) - 0.04 if y_end is None else y_end
    cv.line(xl, ty + 0.36, xl, ye, color, 0.75, dash=True)

i06, i07 = YEARS.index("'06"), YEARS.index("'07")
known = [v for v in tot[i06:i07 + 1] if v is not None]
xm = (p.x(i07) + p.x(i07 + 1)) / 2
# 07/08 경계에 세로 점선(붕괴 시점) + 오른쪽 가운데 높이에 글자 — 19·20·24 글자와 다른 줄
y_line0 = (p.y(max(known)) if known else p.top + 0.40)
cv.line(xm, y_line0, xm, p.bottom - 0.02, DG, 0.75, dash=True)
cv.val(xm + 0.08, 2.95, sh, "사건", i07, color=DG, align="l")
if any(tot[i] is None for i in (i06, i07)):  # 원자료 미확보 연도 표시(06·07 한 번만)
    cv.label((p.x(i06) + p.x(i07)) / 2, p.bottom - 0.40, "n/a", MG, align="c")
# 08~18 공백: 바닥 위 괄호선 + 글자
i08, i18 = YEARS.index("'08"), YEARS.index("'18")
yb = p.bottom - 0.22
xa, xb = p.x(i08) - 0.12, p.x(i18) + 0.12
cv.line(xa, yb, xb, yb, MG, 0.75)
cv.line(xa, yb, xa, yb + 0.12, MG, 0.75)
cv.line(xb, yb, xb, yb + 0.12, MG, 0.75)
i13 = YEARS.index("'13")
cv.val((xa + xb) / 2, yb - 0.38, sh, "사건", i13, color=MG, align="c")
# 19·20·24: 오른쪽 끝을 지시선에 맞춘 글자, 높이를 엇갈림
# 오른쪽 사건일수록 높게(글자가 왼쪽 사건의 지시선을 가로지르지 않게)
event("'24", 0.98, "r", RED)
event("'20", 1.50, "r", RED)
event("'19", 2.02, "r", DG)

prs.save(PPTX)
print("saved")
