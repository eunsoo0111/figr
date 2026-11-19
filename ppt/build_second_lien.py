# -*- coding: utf-8 -*-
"""도표: 2순위(HELOC+CES) RMBS 발행액 + HELOC 잔액 + 주요 사건 — SMIC 양식 PPT(18.50 × 5.10 cm).
원자료 엑셀(유일한 원본)을 먼저 쓰고, 그 엑셀을 load()로 읽어 PPT 네이티브 차트를 만든다.
입력: ../data/C_second_lien_rmbs_issuance_annual_2004_2026.csv · ../data/HELOC_2004_2026_rawdata.xlsx
출력: SMIC_도표_2순위증권화.pptx · SMIC_도표_2순위증권화_원자료.xlsx
"""
import csv, os
import openpyxl
from smic_xlsx import SourceBook, load
import smic_charts_lib as L
from smic_charts_lib import *

L.SIZES["가로"] = (18.50, 5.10)                     # 요청 규격(가로 두 단 18.5 × 5.1)
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
NO, TITLE = "도표 1-1", "2순위(HELOC+CES) RMBS 발행액·HELOC 잔액 추이와 주요 사건"
XLSX = os.path.join(HERE, "SMIC_도표_2순위증권화_원자료.xlsx")
PPTX = os.path.join(HERE, "SMIC_도표_2순위증권화.pptx")
BN = 1_000_000_000                                  # 원자료가 십억 달러 단위 → 달러 원값(1의 자리)

# ── 원천 읽기 ───────────────────────────────────────────────
C = {int(r["year"]): r for r in csv.DictReader(open(os.path.join(DATA, "C_second_lien_rmbs_issuance_annual_2004_2026.csv"), encoding="utf-8-sig"))}
raw = openpyxl.load_workbook(os.path.join(DATA, "HELOC_2004_2026_rawdata.xlsx"), data_only=True)
B = {int(r[0]): r for r in raw["B_잔액_연간"].iter_rows(min_row=2, values_only=True) if isinstance(r[0], int)}
FC = next(float(str(r[2]).lstrip("~>")) for r in raw["E_2026_YTD"].iter_rows(min_row=2, values_only=True) if r[1] == "S&P Global Ratings")
FIGRE = next(r[1] for r in raw["F_증권화잔액·발행사"].iter_rows(min_row=2, values_only=True) if str(r[0]).startswith("FIGRE"))
SRC = {r[0]: r[1] for r in raw["SOURCES"].iter_rows(min_row=2, values_only=True) if r[0]}

YEARS = list(range(2004, 2027))
iss = [round(float(C[y]["chart_value_bn"]) * BN) if C[y]["chart_value_bn"] else None for y in YEARS]
bal = [round(float(B[y][2]) * BN) for y in YEARS]
fc = [round(FC * BN) if y == 2026 else None for y in YEARS]

EVENT = {  # 연도 → 사건(① 순서). Figure 관련은 빨강
    2005: "4Q05 2순위 신규취급 정점",
    2007: "서브프라임 붕괴·증권화 중단",
    2008: "리먼·모노라인 붕괴",
    2019: "FirstKey 위기 후 첫 HELOC 딜",
    2020: "Figure HELOC 첫 증권화(GRADE)",
    2022: "연준 인상 → 1순위 금리 잠김",
    2023: "FIGRE 첫 등급딜·2.0 재개",
    2025: "Figure IPO(25.9)",
    2026: f"FIGRE, HELOC RMBS 잔액 점유 {FIGRE}%",
}
FIG_EV = {2020, 2023, 2025, 2026}
EV_SRC = {2005: "NY Fed SR569", 2007: "NY Fed SR569; HousingWire(S&P, 2007)", 2008: "HousingWire(2008, MBIA·Ambac 2순위 보증손실)",
          2019: "Fitch·Bloomberg(Towd Point HE 2019-HE1)", 2020: "Saluda Grade(GRADE 2020-FIG1)", 2022: "Fed",
          2023: "Figure 보도자료(FIGRE 2023-HE1); KBRA2023", 2025: "Figure 424B4", 2026: "SFA2026(Deutsche Bank, Intex) 2026-09"}


def row_src(y):
    keys = [k.strip() for k in (C[y]["source_key"] or "").split(";") if k.strip()]
    iss_src = ", ".join(keys) if C[y]["chart_value_bn"] else f"발행액 공개치 없음({C[y]['grade']})"
    ev = f" / 사건: {EV_SRC[y]}" if y in EVENT else ""
    return f"발행: {iss_src} / 잔액: NYFED {B[y][1]}{ev}"


# ── 1) 원자료 엑셀 ───────────────────────────────────────────
wb = SourceBook()
wb.add(NO, TITLE, [f"'{str(y)[2:]}" for y in YEARS],
       cols=[("2순위 RMBS 발행(LHS)", iss, FMT["십억"]),
             ("S&P 26년 전망(LHS)", fc, FMT["십억"]),
             ("HELOC 잔액(RHS)", bal, FMT["십억"])],
       labels=[("사건", [EVENT.get(y, "") for y in YEARS]),
               ("발행 등급", [C[y]["grade"] for y in YEARS])],
       row_sources=[row_src(y) for y in YEARS], size="가로", unit="(십억 달러) / (십억 달러)",
       note="발행액은 집계처 혼재(04~05 IMF, 22 KBRA, 23~24 HEL News, 25 KBRA, 26 J.P. Morgan YTD). 2006~07 연간 총액 미확보, "
            "2008~21 휴면(연도별 공개치 없음). 2026 잔액은 3Q, 발행은 9월 중순 YTD.",
       check="2006~07 연간 총액(IMF 유료) 보강 시 갱신; 2025는 KBRA 27.5 vs JPM 29.0·S&P ~30")
wb.save(XLSX)
bk = load(XLSX)
sh = bk[NO]
SRCS = sorted({k.strip() for y in YEARS for k in (C[y]["source_key"] or "").split(";") if k.strip()} | {"NYFED"})

# ── 2) PPT 도표 ─────────────────────────────────────────────
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)", [f"{NO}. {TITLE} — 가로 18.50 × 5.10 cm"],
            checks=["도표 번호(1-1)는 보고서 순서에 맞게 바꾸기",
                    "2006~07 연간 발행액(Inside Mortgage Finance 연감) 확보 시 엑셀 두 칸만 채우고 다시 빌드",
                    "2025 발행액은 집계처별로 다름(KBRA 27.5 / J.P. Morgan 29.0 / S&P ~30) — 각주 유지"])
notes = "원 수치·출처: 엑셀 '" + NO + "' 시트(행별 출처). 출처 키 → 원문:\n" + "\n".join(f"{k}: {SRC.get(k, '')}" for k in SRCS)
cv = chart_slide(prs, NO, TITLE, "가로",
                 "Inside Mortgage Finance, KBRA, HEL News, J.P. Morgan·S&P(SFA), Deutsche Bank/Intex, NY Fed, SMIC 3팀",
                 notes=notes,
                 note_lines=["발행액 집계처 혼재(04~05 IMF, 22·25 KBRA, 23~24 HEL News, 26 J.P. Morgan 9월 중순 YTD). 2006~07 연간 총액 미확보, 2008~21 연도별 공개치 없음",
                             "2026 HELOC 잔액은 3Q 기준. 빨강 번호 = Figure 관련 사건"])
L1, R1 = (0, 200 * BN, 50 * BN), (0, 1200 * BN, 300 * BN)          # 두 축 4칸 — 잔액 선이 막대 라벨 위로 지나가게
S0, S1, S2 = "S&P 26년 전망(LHS)", "2순위 RMBS 발행(LHS)", "HELOC 잔액(RHS)"
p = add_chart(cv, "bar", sh.cats, sh.series(S0, S1, S2), sheet=NO, fmt=sh.fmt(S1), units=("(십억 달러)", ""),
              colors=[PINK, MG, DG], gap=45, overlap=100, yscale=L1)     # 전망(뒤) 위에 실적(앞)을 겹침
color_points(p.ch.plots[0].series[1], MG, {len(YEARS) - 1: RED})
p.rhs([S2], fmt=sh.fmt(S2), yscale=R1, unit="(십억 달러)", colors=[DG])

ix = {y: i for i, y in enumerate(YEARS)}
iss_pts = [ix[y] for y in YEARS if sh.cols[S1][ix[y]] is not None]
b = sh.cols[S2]
pk, lo = b.index(max(b)), b.index(min(b))
# 값 라벨은 차트 데이터레이블(렌더러와 무관하게 막대·점에 붙는다)
p.labels(S1, points=iss_pts, color=DG, pos="outEnd")
p.labels(S0, points=[ix[2026]], fmt='"~"#,##0,,,"E"', color=DG, pos="outEnd")
p.labels(S2, points=[pk, lo, len(b) - 1], fmt="#,##0,,,", color=DG, pos="t")
# 2026 YTD 라벨만 막대 안 흰색(위는 전망 라벨 자리)
ser = next(e for e in L._series_xml(p.ch) if L._ser_name(e) == S1)
d = next(e for e in ser.find(L.qn("c:dLbls")).findall(L.qn("c:dLbl")) if e.find(L.qn("c:idx")).get("val") == str(ix[2026]))
d.find(L.qn("c:dLblPos")).set("val", "inEnd")
d.replace(d.find(L.qn("c:txPr")), L._txpr(CT, WHT))

slot = p.pw / p.n
# 휴면기 음영 + 이름
d0, d1 = p.x(ix[2008]) - slot / 2, p.x(ix[2021]) + slot / 2
cv.band(d0, d1, p.top, p.bottom, BG, 0.7)
cv.label((d0 + d1) / 2, p.bottom - 0.55, "휴면기(2008~21): 연도별 공개 발행치 없음", MG, align="c")
for y in (2006, 2007):
    cv.label(p.x(ix[y]), p.bottom - 0.55, "n.a.", MG, align="c")

# 사건 번호(Figure 관련 빨강): 선·라벨 위
for k, y in enumerate(EVENT, 1):
    i = ix[y]
    v = max(sh.cols[S1][i] or 0, sh.cols[S0][i] or 0)
    tops = [p.y(b[i], rhs=True) - (0.50 if i in (pk, lo, len(b) - 1) else 0.05)]
    if v:
        tops.append(p.y(v) - 0.40)
    cv.num(k, p.x(i), min(tops) - 0.42, RED if y in FIG_EV else DG)

# 사건 목록: 휴면기 음영 위쪽 2단
items = list(EVENT)
cols = [(p.x(ix[2010]) - slot * 0.2, items[:4]), (p.x(ix[2016]) + slot * 0.45, items[4:])]   # 잔액 정점(09) 오른쪽 빈 곳
for x0, ys in cols:
    for r, y in enumerate(ys):
        k = items.index(y) + 1
        cv.val(x0, p.top + 0.14 + r * 0.33, sh, "사건", ix[y], pre=f"{circled(k)} ", color=RED if y in FIG_EV else DG)

prs.save(PPTX)
print("saved", PPTX, XLSX)
