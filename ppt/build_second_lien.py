# 도표: 미국 2순위(HELOC+CES) 증권화 발행액 — 금융위기 전 vs 재개 이후, 중간 휴면기 이벤트
# 와이드 18.50 x 5.10 cm. 숫자는 원자료 엑셀에서만 읽는다.
#   python build_second_lien.py  →  SMIC_도표_2순위증권화.pptx + SMIC_도표_2순위증권화_원자료.xlsx
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx.enum.dml import MSO_PATTERN
from pptx.dml.color import RGBColor
import smic_charts_lib as L
from smic_charts_lib import *
from smic_xlsx import SourceBook, load

L.SIZES["와이드"] = (18.50, 5.10)
B = 1_000_000_000
M = 1_000_000
NO = "도표 1-1"
TITLE = "미국 2순위(HELOC·CES) 증권화 발행액 추이: 사라졌다 돌아온 HELOC 유통시장"
XLSX = "SMIC_도표_2순위증권화_원자료.xlsx"
PPTX = "SMIC_도표_2순위증권화.pptx"
SRC = "Inside Mortgage Finance, S&P, KBRA, HEL News, J.P. Morgan(SFA), Figure 공시, SMIC 3팀"

# ───────── 1. 원자료 엑셀 ─────────
wb = SourceBook()
SL, SR, SG = NO + " 위기전", NO + " 재개후", NO + " 동사근거"

ws = wb.add(SL, "2순위 전용 비에이전시 MBS 발행액(2004~2007)", ["'04", "'05", "'06", "'07"],
            cols=[("2순위 RMBS 발행액", [None, 60_740 * M, None, None], FMT["십억"]),
                  ("연간 단편 합(하한)", [None, None, 0, 0], FMT["십억"]),
                  ("상반기 HELOC", [None, None, 0, None], FMT["십억"], "hide"),
                  ("분기 HELOC(S&P)", [None, None, 3_900 * M, 7_280 * M], FMT["십억"], "hide"),
                  ("분기 CES(S&P)", [None, None, 13_540 * M, 11_120 * M], FMT["십억"], "hide")],
            row_sources=["Inside Mortgage Finance(2006): 2005년 +23.6% → 2004 = 2005/1.236 역산",
                         "Inside Mortgage Finance(2006): 2순위 전용 MBS USD 60.74bn",
                         "IMF(2006) 상반기 HELOC = Countrywide 9.00bn ÷ 2/3 + S&P 4Q06 HELOC 3.90bn·CES 13.54bn",
                         "S&P via HousingWire(2007): 1Q07 HELOC 7.28bn·CES 11.12bn(이후 시장 붕괴)"],
            unit="(십억 달러)", size="와이드",
            note="'06·'07 연간 총액은 유료 원문(IMF 연감)이라 미확보 → 공개된 반기·분기 단편 합을 하한으로 표시",
            check="'06·'07 연간 총액(Inside Mortgage Finance 연감) 확보 시 교체")
ws["B2"] = "=B3/1.236"                                   # 2004 역산
ws["D4"] = "=9000000000/(2/3)"                           # 2006 상반기 HELOC(Countrywide 9.00bn = 2/3)
ws["C4"] = "=D4+E4+F4"                                   # 2006 하한 = 상반기 HELOC + 4Q HELOC + 4Q CES
ws["C5"] = "=E5+F5"                                      # 2007 하한 = 1Q HELOC + 1Q CES

wg = wb.add(SG, "동사(Figure) HELOC 증권화 발행액 역산 근거",
            ["'24 연간", "'25.06말 누적", "'25.09말 누적", "'25말 누적", "'25 증가분(vs '24)", "9M26 증가분(vs 9M25)"],
            cols=[("값", [2_300 * M, 4_500 * M, 5_800 * M, 7_000 * M, 1_200 * M, 4_100 * M], FMT["십억"]),
                  ("건수", [7, 15, 19, 23, None, None], "#,##0")],
            row_sources=["GlobalCapital(Finsight 집계): 2024년 HELOC RMBS 7건 USD 2.3bn+",
                         "Figure 424B4(2025-09): 15건 USD 4.5bn(2023년 4월 이후 등급 딜)",
                         "Figure 424B4(2026): 19건 USD 5.8bn",
                         "Figure 10-K FY2025: 23건 USD 7bn",
                         "Figure 10-K FY2025: 프로그램 수수료 대상 증권화 물량 +USD 1.2bn",
                         "Figure 10-Q 3Q26: 9M 프로그램 수수료 대상 증권화 물량 +USD 4.1bn"],
            unit="(십억 달러)", size="—", first="항목",
            note="도표에는 직접 쓰지 않음. 재개후 시트의 동사 발행액 수식이 이 값을 참조",
            check="근거 시트(차트 없음)")

g = lambda r: f"'{SG}'!B{r}"                              # 근거 시트 행: 2=24연간 3=25.06 4=25.09 5=25말 6=25증가 7=9M26증가
wr = wb.add(SR, "2순위 RMBS 발행액(2022~2026) 중 동사 비중", ["'22", "'23", "'24", "'25", "'26"],
            cols=[("동사(Figure)", [0, 0, 0, 0, 0], FMT["십억"]),
                  ("기타 발행사", "={전체 발행액}-{동사(Figure)}", FMT["십억"]),
                  ("2026 연간 전망(S&P ~40)", [None, None, None, None, 0], FMT["십억"]),
                  ("전체 발행액", [626 * M, 4_800 * M, 15_600 * M, 27_500 * M, 29_829 * M], FMT["십억"], "hide"),
                  ("S&P 연간 전망", [None, None, None, None, 40_000 * M], FMT["십억"], "hide"),
                  ("동사 비중", "={동사(Figure)}/{전체 발행액}", FMT["%0"], "hide")],
            row_sources=["KBRA(2023): 3건 약 USD 626mn / 동사 등급 딜 없음(첫 등급 딜 2023.04)",
                         "HEL News(2024): HELOC 2.9 + CES 1.9 / 동사 = '25말 누적 − '24 − '25 (SMIC 역산)",
                         "HEL News(2024): HELOC 5.4 + CES 10.2 / 동사 = Finsight 7건 2.3bn+",
                         "KBRA(2026): 75건 / 동사 = '24 + 10-K 증가분 1.2bn (SMIC 역산)",
                         "J.P. Morgan(SFA 2026-09-17) 9월 중순 YTD, S&P 연간 ~40 / 동사 1~9월 = 9M25 + 10-Q 증가분 4.1bn (SMIC 역산)"],
            unit="(십억 달러)", size="와이드",
            note="집계처 혼재(KBRA·HEL News·J.P. Morgan). 동사 발행액은 공시 누적치·수수료 대상 물량으로 SMIC 역산",
            check="2026 동사 값은 1~9월, 시장 값은 9월 중순까지 — 연간 확정치 나오면 교체")
wr["B3"] = f"={g(5)}-{g(2)}-({g(2)}+{g(6)})"              # 2023 = 7.0 − 2024 − 2025
wr["B4"] = f"={g(2)}"                                     # 2024 = 2.3
wr["B5"] = f"={g(2)}+{g(6)}"                              # 2025 = 2.3 + 1.2
wr["B6"] = f"=({g(4)}-B3-B4)+{g(7)}"                     # 9M26 = (9M25 = 5.8 − '23 − '24) + 4.1
wr["D6"] = "=F6-E6"                                       # 전망 잔여 = 40 − YTD
wb.save(XLSX)
bk = load(XLSX)
L_, R_ = bk[SL], bk[SR]

# ───────── 2. PPT 도표 ─────────
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)", [f"{NO}. {TITLE} (와이드 18.50 × 5.10 cm)"],
            checks=["'06·'07 연간 총액 미확보(하한 표시)", "동사 발행액 SMIC 역산 — 원자료 '동사근거' 시트",
                    "도표 번호는 보고서 순서에 맞게 수정"])
cv = chart_slide(prs, NO, TITLE, "와이드", SRC,
                 notes=("막대: 2순위(HELOC+CES) RMBS 발행액. '04 = '05/1.236, '05 IMF 60.74bn, '06·'07 = 공개 단편 합(하한). "
                        "'22 KBRA 0.626, '23·'24 HEL News, '25 KBRA 27.5, '26 J.P. Morgan 9월 중순 YTD 29.8 + S&P 연간 전망 ~40. "
                        "동사: '24 Finsight 2.3bn(7건), '25 = '24 + 10-K 수수료 대상 물량 증가 1.2bn, '23 = '25말 누적 7.0 − '24 − '25, "
                        "'26 1~9월 = ('25.09말 누적 5.8 − '23 − '24) + 10-Q 9M 증가 4.1bn."),
                 note_lines=["'04는 '05 증가율(+23.6%)로 역산, '06·'07은 공개된 반기·분기 단편 합(하한). 2022년 이후 집계처 혼재(KBRA·HEL News·J.P. Morgan)",
                             "동사 발행액은 공시 누적치와 프로그램 수수료 대상 증권화 물량으로 SMIC 역산. 2026년은 동사 1~9월, 시장 9월 중순 누적"])
W, H = cv.w, cv.h
YS = (0, 80 * B, 20 * B)
TOP, XL = 0.62, 0.36                    # 플롯 위 여백(단위·범례 줄), 아래 여백(연도 글자)
CH = H - 0.36                           # 차트 높이(아래 0.36 = 동사 발행액 줄)
YW = tw("80") + 0.22                    # 왼쪽 눈금 글자 폭
LX1, RX0 = 4.75, 10.75                  # 왼쪽 차트 끝 / 오른쪽 차트 시작
# 눈금 글자·연도는 차트가 아니라 도형으로 그린다 → PowerPoint·LibreOffice 모두 플롯 위치가 계산값과 같다
IN = (0.05, TOP, 0.05, XL)


def bare(p):
    """막대만 남긴 차트: 값축 숨김, 가로축은 선·눈금만(글자 없음)"""
    pa = p.ch._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    for ax in pa.findall(qn("c:catAx")):
        ax.find(qn("c:tickLblPos")).set("val", "none")


pl = add_chart(cv, "stack", L_.cats, L_.series("2순위 RMBS 발행액", "연간 단편 합(하한)"), sheet=SL,
               fmt=L_.fmt("2순위 RMBS 발행액"), x=YW, y=0, w=LX1 - YW, h=CH, colors=[MG, LG], yscale=YS,
               legend=False, gap=55, inner=IN, axes=False)
pr = add_chart(cv, "stack", R_.cats, R_.series("동사(Figure)", "기타 발행사", "2026 연간 전망(S&P ~40)"), sheet=SR,
               fmt=R_.fmt("기타 발행사"), x=RX0, y=0, w=W - RX0, h=CH, colors=[RED, MG, PINK], yscale=YS,
               legend=False, gap=95, inner=IN, axes=False)
bare(pl); bare(pr)
s26 = pr.ch.plots[0].series[2]                                   # 2026 전망: 연분홍 + 점선 테두리
s26.format.line.color.rgb = RGBColor.from_string(PINK2)
s26.format.line.width = Pt(0.75)
from lxml import etree
etree.SubElement(s26.format.line._get_or_add_ln(), qn("a:prstDash")).set("val", "sysDash")

y0, y1 = pl.top, pl.bottom
# 세로축(왼쪽): 선·바깥 눈금·글자
ax_x = pl.px0
cv.line(ax_x, y0, ax_x, y1, DG, 0.75)
for v in range(0, 81, 20):
    yy = pl.y(v * B)
    cv.line(ax_x - 0.08, yy, ax_x, yy, DG, 0.75)
    cv.label(ax_x - 0.12, yy - 0.18, str(v), DG, align="r")
cv.label(0.0, 0.02, "(십억 달러)", DG)
# 연도
for p_, S_ in ((pl, L_), (pr, R_)):
    for i, c in enumerate(S_.cats):
        cv.label(p_.x(i), y1 + 0.03, c, DG, align="c")

# 가운데: 휴면기 빗금 + 끊긴 가로축
mx0, mx1 = LX1 + 0.12, RX0 - 0.12
hz = cv.box(mx0, y0, mx1 - mx0, y1 - y0, shape="rect")
hz.fill.patterned(); hz.fill.pattern = MSO_PATTERN.WIDE_UPWARD_DIAGONAL
hz.fill.fore_color.rgb = RGBColor.from_string(LG); hz.fill.back_color.rgb = RGBColor.from_string(WHT)
cv.back(hz)
cv.line(pl.px0 + pl.pw, y1, pr.px0, y1, DG, 0.75)
for xb in (mx0 - 0.03, mx1 + 0.03):                              # 축 끊김 //
    cv.box(xb - 0.09, y1 - 0.1, 0.18, 0.2, fill=WHT, shape="rect")
    for dx in (-0.045, 0.045):
        cv.line(xb + dx - 0.05, y1 + 0.09, xb + dx + 0.05, y1 - 0.09, DG, 0.75)
mid = (mx0 + mx1) / 2
cv.label(mid, y1 + 0.03, "'08 ~ '21", DG, align="c")
cv.box(mid - 2.3, y1 - 0.92, 4.6, 0.64, fill=WHT, shape="rect",
       paras=[[("2순위 증권화 휴면기 14년", DG, None, None, False)],
              [("HELOC 유통시장 침체·소멸", RED, None, None, False)]], align="c")

# 막대 값
LB = 0.40
for i in (0, 1):
    cv.val(pl.x(i), pl.y(L_.cols["2순위 RMBS 발행액"][i]) - LB, L_, "2순위 RMBS 발행액", i, align="c")
for i in (2, 3):
    cv.val(pl.x(i), pl.y(L_.cols["연간 단편 합(하한)"][i]) - LB, L_, "연간 단편 합(하한)", i, pre="≥", align="c", color=MG)
tot = R_.cols["전체 발행액"]
for i in range(4):
    cv.val(pr.x(i), pr.y(tot[i]) - LB, R_, "전체 발행액", i, align="c")
cv.val(pr.x(4), pr.y(tot[4]) + 0.04, R_, "전체 발행액", 4, align="c", color=WHT)       # YTD는 막대 안 흰 글자
cv.val(pr.x(4), pr.y(R_.cols["S&P 연간 전망"][4]) - LB, R_, "S&P 연간 전망", 4, pre="~", post="E",
       fmt="#,##0,,,", align="c", color=PINK2)

# 동사 발행액 줄(연도 아래)
ry = CH + 0.02
cv.label(pr.px0 - 0.15, ry, "동사(Figure) 발행액", RED, align="r")
cv.label(pr.x(0), ry, "—", RED, align="c")
for i in (1, 2, 3):
    cv.val(pr.x(i), ry, R_, "동사(Figure)", i, align="c", color=RED)
cv.val(pr.x(4), ry, R_, "동사(Figure)", 4, post="(9M)", align="c", color=RED)

# 사건 타임라인: 휴면기 → 재개 구간을 한 줄로 관통. 회색 = 시장, 빨강 = 동사
TY = 1.30
cv.line(mx0 + 0.1, TY, W - 0.15, TY, MG, 1.0)
EV = [  # (x, 날짜, 설명, 위/아래, 동사 여부)
    (mx0 + 0.55, "'08.09", "리먼 파산·모노라인 붕괴", "up", False),
    (mx0 + 3.0, "'19.06", "위기 후 첫 HELOC 딜", "dn", False),
    (mx1 - 1.05, "'20", "동사 첫 증권화", "up", True),
    (pr.x(1), "'23.04", "동사 첫 등급 딜", "up", True),
    (pr.x(2), "'24.06", "Figure Connect 출범", "dn", True),
    (pr.x(3), "'25.09", "동사 나스닥 상장", "up", True),
]
for x, d, t, pos, hot in EV:
    c = RED if hot else DG
    w_ = max(tw(d), tw(t)) + 0.12
    x0 = min(max(mx0 + 0.02, x - w_ / 2), W - w_)
    yy = TY - 0.80 if pos == "up" else TY + 0.16
    cv.box(x0, yy, w_, 0.64, fill=WHT, shape="rect", margin=0,
           paras=[[(d, c, None, None, False)], [(t, DG, None, None, False)]], align="c", anchor="m")
    cv.box(x - 0.1, TY - 0.1, 0.2, 0.2, fill=WHT, line=c, lw=1.5, shape="oval")
for i in (1, 2, 3):                                              # 사건 → 해당 연도 막대 점선
    top_lab = pr.y(tot[i]) - LB - 0.02
    start = TY + (0.82 if i == 2 else 0.12)
    if top_lab - start > 0.1:
        cv.line(pr.x(i), start, pr.x(i), top_lab, PINK2, 0.5, dash=True)

# 범례(단위 오른쪽 한 줄)
lx, ly = tw("(십억 달러)") + 0.45, 0.02
for name, col, line in [("2순위 RMBS 발행액", MG, None), ("동사(Figure)", RED, None),
                        ("'06·'07 공개 단편 합(하한)", LG, None), ("2026 연간 전망(S&P)", PINK, PINK2)]:
    cv.box(lx, ly + 0.07, 0.2, 0.2, fill=col, line=line, lw=0.5, shape="rect")
    cv.label(lx + 0.27, ly, name, DG)
    lx += 0.27 + tw(name) + 0.45

prs.save(PPTX)
print("saved", PPTX, XLSX)
