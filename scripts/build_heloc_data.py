"""HELOC 시장 2004~2026 원자료 빌더.

입력: data/heloc/source_raw/HHD_C_Report_2026Q3.xlsx (NY Fed Household Debt and Credit, 2026Q3, 2026-11 공표)
출력: data/heloc/HELOC_2004_2026_rawdata.xlsx + CSV 4종

증권화(2순위 RMBS) 발행액은 공개 원문에서 직접 확인한 값만 손으로 넣었고, 셀마다 출처·등급을 단다.
"""
from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "heloc"
RAW = OUT / "source_raw" / "HHD_C_Report_2026Q3.xlsx"

NYFED_SRC = ("NY Fed Quarterly Report on Household Debt and Credit 2026Q3 (Nov 2026), "
             "NY Fed Consumer Credit Panel/Equifax. "
             "https://www.newyorkfed.org/microeconomics/hhdc")

# ---------------------------------------------------------------- NY Fed

def qkey(lbl):
    if isinstance(lbl, dt.datetime):
        return f"{lbl.year}Q{(lbl.month - 1) // 3 + 1}"
    if isinstance(lbl, str) and ":Q" in lbl:
        y, q = lbl.split(":Q")
        return f"20{y.strip()}Q{q.strip()}"
    return None


def simple_table(ws, cols):
    """첫 열이 분기 라벨인 표에서 {분기: {열이름: 값}}."""
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if cols[0] in [str(c).strip() for c in r if c is not None])
    hdr = [str(h).strip() if h else None for h in rows[hdr_i]]
    out = {}
    for r in rows[hdr_i + 1:]:
        k = qkey(r[0])
        if not k:
            continue
        out[k] = {c: r[hdr.index(c)] for c in cols}
    return out


def load_nyfed():
    wb = openpyxl.load_workbook(RAW, read_only=True, data_only=True)
    p3 = simple_table(wb["Page 3 Data"], ["Mortgage", "HE Revolving", "Total"])
    p4 = simple_table(wb["Page 4 Data"], ["HE Revolving"])
    p12 = simple_table(wb["Page 12 Data"], ["HELOC"])
    p13 = simple_table(wb["Page 13 Data"], ["HELOC"])
    p14 = simple_table(wb["Page 14 Data"], ["HELOC"])
    # Page 10: 분기 라벨 행 = 카드, 바로 다음 행 = HE Revolving (Balance, Available, Limit)
    rows = list(wb["Page 10 Data"].iter_rows(values_only=True))
    p10 = {}
    for i, r in enumerate(rows):
        k = qkey(r[0])
        if k and i + 1 < len(rows):
            n = rows[i + 1]
            p10[k] = {"bal": n[4], "avail": n[5], "limit": n[6]}
    out = []
    for k in p3:
        bal = p3[k]["HE Revolving"]
        lim = p10[k]["limit"]
        # Page 10 잔액은 일부 분기가 소수 2자리로 반올림돼 있음 → Page 3 값을 잔액으로 쓰고 차이만 점검
        assert abs(p10[k]["bal"] - bal) < 0.006, (k, p10[k]["bal"], bal)
        out.append({
            "quarter": k,
            "year": int(k[:4]),
            "heloc_bal_bn": round(bal * 1000, 1),
            "heloc_limit_bn": round(lim * 1000, 1),
            "heloc_avail_bn": round(p10[k]["avail"] * 1000, 1),
            "util_pct": round(bal / lim * 100, 1),
            "heloc_accts_mn": p4[k]["HE Revolving"],
            "mortgage_bal_bn": round(p3[k]["Mortgage"] * 1000, 1),
            "total_hh_debt_bn": round(p3[k]["Total"] * 1000, 1),
            "heloc_share_hh_pct": round(bal / p3[k]["Total"] * 100, 2),
            "dq90_pct": round(p12[k]["HELOC"], 2),
            "flow30_pct": round(p13[k]["HELOC"], 2),
            "flow90_pct": round(p14[k]["HELOC"], 2),
        })
    return out


# ------------------------------------------------- 2순위 RMBS 증권화 (손 입력, 출처 필수)

S = {
    "IMF2005": ("Inside Mortgage Finance, 'Second-Lien MBS Issuance Up Sharply in 2005 Despite Slowdown in HELOC "
                "Securitization' (2006). https://www.insidemortgagefinance.com/products/216118"),
    "IMF2006": ("Inside MBS & ABS, 'Countrywide Leads in Securitization of Second Mortgages as Sector Sagged in "
                "Second Quarter' (2006). https://www.insidemortgagefinance.com/products/217095"),
    "SP2007": ("HousingWire, 'S&P: CES Securitizations Flat, HELOCs Soar in First Quarter of 2007' (2007, S&P 평가 딜 기준). "
               "https://www.housingwire.com/articles/sp-ces-securitizations-flat-helocs-soar-first-quarter-2007/"),
    "KBRA2023": ("HousingWire, 'Home-equity lending blossomed in 2023' (2024, KBRA 인용). "
                 "https://www.housingwire.com/articles/home-equity-lending-blossomed-in-2023/"),
    "HELN2024": ("HEL News, 'CEO Letter to Stakeholders' (2024-12). "
                 "https://www.hel.news/articles/bank-home-equity-news/2024-stakeholders-letter/"),
    "KBRA2026": ("KBRA, 'Second Lien Securitizations: HELOC and CES Deals Catch a 2.0 Wind' (2026-09-08). "
                 "https://www.kbra.com/publications/NvPkXxcx/"),
    "SFA2026": ("SFA Research Corner, 'From Revival to Scale: Second-Lien RMBS Moves Into a New Phase' (2026-09-17); "
                "S&P Global Ratings·J.P. Morgan·Deutsche Bank/Intex 인용. source_raw/ 에 PDF 보관"),
    "BOFA2026": ("National Mortgage News, 'Second-lien issuance expected to reach $41 billion this year' (2026-07, BofA). "
                 "https://www.nationalmortgagenews.com/news/second-lien-issuance-expected-to-reach-41-billion-this-year"),
    "HELN2026": ("HEL News Capital Markets 2026-11-17/18 (자체 집계 'home-equity issuance', HEI 포함 가능). "
                 "https://www.hel.news/category/mbs/"),
}

# chart_bn: 도표에 한 줄로 쓸 값(연간 총 2순위 RMBS 발행, $bn). None = 원문 미확보 → 막대 비움.
SEC_ANNUAL = [
    dict(year=2004, heloc=None, ces=None, total=49.1, deals=None, chart_bn=49.1, grade="파생",
         src="IMF2005", note="2005년 $60.74B가 전년 대비 +23.6% → 2004 = 60.74/1.236. 2순위 전용 비에이전시 MBS(HELOC+CES)."),
    dict(year=2005, heloc=None, ces=None, total=60.74, deals=None, chart_bn=60.74, grade="확정",
         src="IMF2005", note="2순위 전용 MBS. 서브프라임 딜에 섞여 들어간 2순위 $10.73B 별도(미포함). HELOC 증권화는 둔화."),
    dict(year=2006, heloc=None, ces=None, total=None, deals=None, chart_bn=None, grade="부분",
         src="IMF2006; SP2007", note="연간 총액 원문 미확보(유료 IMF). H1 HELOC ≈ $13.5B(파생: Countrywide $9.00B = '거의 2/3'). "
                                    "4Q06 S&P 평가분 HELOC $3.90B(5건)·CES $13.54B(24건). 분기 시트 참조."),
    dict(year=2007, heloc=None, ces=None, total=None, deals=None, chart_bn=None, grade="부분",
         src="SP2007", note="연간 총액 원문 미확보. 1Q07 S&P 평가분 HELOC $7.28B(10건)·CES $11.12B(23건). 하반기 시장 붕괴."),
]
for y in range(2008, 2022):
    SEC_ANNUAL.append(dict(year=y, heloc=None, ces=None, total=None, deals=None, chart_bn=None, grade="미확보(휴면)",
                           src="KBRA2026", note="KBRA: 금융위기 후 '수년간 제한적 활동'. 연도별 정량 원문 없음 → 도표에선 0 근처/공백 처리 권장."))
SEC_ANNUAL += [
    dict(year=2022, heloc=None, ces=None, total=0.626, deals=3, chart_bn=0.626, grade="확정",
         src="KBRA2023; SFA2026", note="KBRA: HELOC/CES 딜 3건 약 $626M. S&P는 '약 $1B'로 표기."),
    dict(year=2023, heloc=2.9, ces=1.9, total=4.8, deals=21, chart_bn=4.8, grade="확정",
         src="HELN2024; KBRA2023", note="HEL News: HELOC 13건 $2.9B + CES 8건 $1.9B (HEI 포함 시 $5.4B). KBRA 추정은 약 $4B(약 16건)."),
    dict(year=2024, heloc=5.4, ces=10.2, total=15.6, deals=59, chart_bn=15.6, grade="확정",
         src="HELN2024", note="HEL News: HELOC 18건 $5.4B + CES 41건 $10.2B (HEI 포함 $17B+). Figure 단독 HELOC 7건 $2.3B+(GlobalCapital)."),
    dict(year=2025, heloc=None, ces=None, total=27.5, deals=75, chart_bn=27.5, grade="확정",
         src="KBRA2026; SFA2026; BOFA2026", note="KBRA: 75건(CES 43·HELOC 30·혼합 2). J.P. Morgan $29.0B(28,997mn), BofA $29B, S&P 약 $30B."),
    dict(year=2026, heloc=None, ces=None, total=29.829, deals=None, chart_bn=29.829, grade="YTD(9월 중순)",
         src="SFA2026; KBRA2026; BOFA2026; HELN2026",
         note="J.P. Morgan YTD $29.8B(SFA 9/17 게재; 전년 동기 $16.6B). KBRA 2Q말 >$23B(은행 매각 혼합 풀 $6B 포함), "
              "BofA 7/24 $24B. 연간 전망 S&P ~$40B·BofA $41B. HEL News 11/18 자체 집계 '$40B 돌파'."),
]

PRE_CRISIS_Q = [
    ("2006H1", "HELOC", 13.5, None, "파생", "IMF2006", "Countrywide HELOC $9.00B = 상반기 HELOC 증권화의 '거의 2/3' → 9.00/(2/3)"),
    ("2006H1", "Countrywide 2순위 전체", 12.27, None, "확정", "IMF2006", "Countrywide 단일 발행사 2순위 MBS"),
    ("2006Q4", "HELOC", 3.90, 5, "확정", "SP2007", "S&P 평가 딜 기준"),
    ("2006Q4", "CES", 13.54, 24, "확정", "SP2007", "S&P 평가 딜 기준"),
    ("2007Q1", "HELOC", 7.28, 10, "확정", "SP2007", "S&P 평가 딜 기준"),
    ("2007Q1", "CES", 11.12, 23, "확정", "SP2007", "S&P 평가 딜 기준"),
]

YTD_2026 = [
    ("2026-06-30", "KBRA", ">23.0", "실적", "은행 매각 혼합(CES/HELOC) 풀 $6B 1건 포함", "KBRA2026"),
    ("2026-07-24", "BofA Securities", "24", "실적", "GFC 이후 같은 시점 최대", "BOFA2026"),
    ("2026-09 중순", "J.P. Morgan (HELOC/CES)", "29.829", "실적", "전년 동기 16.634", "SFA2026"),
    ("2026-11-18", "HEL News 자체", ">40", "실적", "'home-equity issuance' 정의, HEI 포함 가능 → 타 집계와 직접 비교 주의", "HELN2026"),
    ("2026 연간", "S&P Global Ratings", "~40", "전망", "2026-02 리포트", "SFA2026"),
    ("2026 연간", "BofA Securities", "41", "전망", "2026-07", "BOFA2026"),
]

OUTSTANDING_SEC = [
    ("HELOC RMBS 잔액", "~$0.5B (2022) → ~$5B (2024) → $21.5~22B (2026)", "SFA2026"),
    ("CES RMBS 잔액", "$1.4B (2022년 이전) → ~$39B (2026)", "SFA2026"),
    ("KBRA HELOC/CES 인덱스", "2023-04: 4건 $375.2M → 2026-07: 140건 $34.2B", "KBRA2026"),
    ("90일+ 연체(2026-07, KBRA)", "CES 1.03% / HELOC 1.69%", "KBRA2026"),
]

SHELF_HELOC = [("FIGRE (Figure)", 40), ("GRADE (Saluda Grade)", 13), ("JPMMT", 11), ("TPMT", 10), ("GSMBS", 8),
               ("ACHM", 5), ("PRMI", 5), ("OBX", 3), ("AOMT", 2), ("CMLTI", 1), ("BRAVO", 1), ("AVHO", 1)]
SHELF_CES = [("RCKT (Rocket)", 32), ("JPMMT", 24), ("TPMT", 12), ("VSTA", 5), ("BRAVO", 4), ("EFMT", 4),
             ("COOPR", 4), ("GSMBS", 3), ("SAN", 2), ("BVCES", 2), ("기타", 8)]

PASTED_CHECK = [
    ("2020~2021 연간 $1B 미만(휴면)", "방향 맞음·수치 근거 없음",
     "KBRA는 GFC 후 '제한적 활동'이라고만 함. 2020~21 연간 정량 원문 미확보.", "KBRA2026"),
    ("2022 연간 약 $4~5B, 10여 건", "틀림",
     "KBRA: 3건 약 $626M. S&P: 약 $1B. 2023년 4월 KBRA 인덱스도 4건 $375M뿐.", "KBRA2023; SFA2026"),
    ("2023 연간 약 $7~8B", "과대",
     "HEL News $4.8B(HELOC $2.9B+CES $1.9B, HEI 포함 $5.4B). KBRA 약 $4B.", "HELN2024; KBRA2023"),
    ("2023-04 4개 딜($375.2M)", "맞음(단, 발행액 아님)",
     "KBRA 인덱스 '잔액' 기준 시작점이지 2023년 4월 발행액이 아님.", "KBRA2026"),
    ("2024 연간 약 $14~16B, 40~45건", "금액 맞음·건수 틀림",
     "HEL News $15.6B, 59건(CES 41 + HELOC 18).", "HELN2024"),
    ("2025 $27.5B, 75건(CES 43·HELOC 30·혼합 2)", "맞음",
     "KBRA 2026-09-08. 단 'KBRA 2026년 9월 섹터 리포트' = 위 리서치. JPM $29.0B, S&P ~$30B로 집계처별 차이.", "KBRA2026"),
    ("2026 H1 $23.0B 초과, 은행 매각 $6B 혼합 풀", "맞음", "KBRA 2Q말 기준.", "KBRA2026"),
    ("H1만으로 2025 연간의 83% 돌파", "계산 맞음", "23.0/27.5 = 83.6%. 다만 '초과'이므로 하한.", "KBRA2026"),
    ("2026-07 누적 잔액 $34.2B, 140건", "맞음", "KBRA HELOC/CES 인덱스 잔액.", "KBRA2026"),
    ("SEC ABS-15G로 144A 딜 전수 추출 가능", "주의",
     "ABS-15G는 실사 결과(Rule 15Ga-2) 공시용이라 딜별 발행액·풀 잔액이 표로 다 들어있다는 보장 없음. "
     "efts.sec.gov는 이 환경에서 403(접근 불가)이라 미검증.", "-"),
]


# ---------------------------------------------------------------- 쓰기

HDR_FILL = PatternFill("solid", fgColor="1F2937")
HDR_FONT = Font(bold=True, color="FFFFFF")
NOTE_FONT = Font(italic=True, color="555555", size=9)


def write_table(ws, headers, rows, start_row=1, widths=None):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=start_row, column=j, value=h)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows, start_row + 1):
        for j, v in enumerate(r, 1):
            ws.cell(row=i, column=j, value=v).alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 30)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=2)
    for j, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def write_csv(path, headers, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


def main():
    q = load_nyfed()
    q04 = [r for r in q if r["year"] >= 2004]
    peak = max(q, key=lambda r: r["heloc_bal_bn"])
    trough = min([r for r in q if r["year"] >= 2010], key=lambda r: r["heloc_bal_bn"])
    last = q[-1]

    wb = openpyxl.Workbook()

    # README
    ws = wb.active
    ws.title = "README"
    lines = [
        "HELOC 시장 2004~2026 원자료 (금융위기 전후)",
        f"빌드: scripts/build_heloc_data.py | 최신 분기: {last['quarter']} | 작성일 2026-11-19",
        "",
        "[시트 구성]",
        "A_잔액_분기  : 미국 HELOC 잔액·한도·이용률·계좌수·연체 (NY Fed, 분기, 2003Q1~2026Q3) ← 메인 '시장 규모' 시계열",
        "B_잔액_연간  : A의 연말(4Q) 값 + YoY. 2026은 3Q 최신치",
        "C_증권화_연간: 2순위(HELOC+CES) RMBS 발행액 = '유동화 시장' 플로우. 셀별 출처·등급",
        "D_위기전_분기: 2006~2007 분기 단편 (연간 총액 원문은 유료라 미확보)",
        "E_2026_YTD  : 2026 누적 발행 집계처별 비교 + 연간 전망",
        "F_증권화잔액·발행사: HELOC/CES RMBS 잔액, 발행 선반(shelf)별 점유 — FIGRE 40%",
        "G_붙여넣은표_검증: 받으신 표 항목별 원문 대조 결과",
        "SOURCES      : 출처 키 → 원문",
        "",
        "[핵심 숫자]",
        f"HELOC 잔액 정점: {peak['quarter']} ${peak['heloc_bal_bn']:.0f}B (금융위기 직후)",
        f"HELOC 잔액 저점: {trough['quarter']} ${trough['heloc_bal_bn']:.0f}B",
        f"최신: {last['quarter']} ${last['heloc_bal_bn']:.0f}B, 한도 ${last['heloc_limit_bn']:.0f}B, 이용률 {last['util_pct']}%",
        "",
        "[정의·주의]",
        "1) 'HE Revolving' = HELOC 잔액(신용평가사 Equifax 패널 기반). 폐쇄형 2순위(CES, home equity loan)는 NY Fed에선 'Mortgage'에 포함되어 분리 불가.",
        "2) 잔액(stock)과 증권화 발행(flow)은 단위·개념이 다름. 한 도표에 넣을 땐 이중축 또는 2패널 권장.",
        "3) 2004~2007 증권화는 2순위 전용 비에이전시 MBS(HELOC+CES) 기준이며 서브프라임 딜 내 2순위는 제외. 2022~ 은 '2순위 RMBS(Second lien 2.0)'.",
        "4) 2008~2021 증권화 연도별 정량치는 공개 원문이 없음. 0으로 채우지 말고 '휴면(n/a)'로 표기 권장.",
        "5) 집계처마다 2025년 값이 다름: KBRA $27.5B / J.P. Morgan $29.0B / BofA $29B / S&P ~$30B.",
        "6) 2004~2007 연간 HELOC 단독 증권화액(Inside Mortgage Finance 연감)·Bloomberg/Finsight 딜 단위 덤프는 유료라 미포함. 학교 터미널로 보강 가능.",
        "7) 미검증 2차 인용(본문에 미반영): 'KBRA 애널리스트: 2004~2007 2순위·HELOC 약 $200B 증권화', '2007년 초 증권화 2순위 잔액 정점 $87B' — 원문(nationalmortgageprofessional.com) 403으로 확인 불가.",
    ]
    for i, t in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=t)
        if i == 1:
            c.font = Font(bold=True, size=14)
        elif t.startswith("["):
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 140

    # A 분기
    ws = wb.create_sheet("A_잔액_분기")
    a_hdr = ["분기", "HELOC 잔액($bn)", "HELOC 한도($bn)", "미사용 한도($bn)", "이용률(잔액/한도,%)", "HELOC 계좌수(백만)",
             "모기지 잔액($bn)", "가계부채 총액($bn)", "HELOC/가계부채(%)", "90일+ 연체 잔액비중(%)",
             "30일+ 신규연체 전이율(%,연율)", "90일+ 신규연체 전이율(%,연율)"]
    a_rows = [[r["quarter"], r["heloc_bal_bn"], r["heloc_limit_bn"], r["heloc_avail_bn"], r["util_pct"], r["heloc_accts_mn"],
               r["mortgage_bal_bn"], r["total_hh_debt_bn"], r["heloc_share_hh_pct"], r["dq90_pct"], r["flow30_pct"], r["flow90_pct"]]
              for r in q]
    write_table(ws, a_hdr, a_rows, widths=[9, 12, 12, 12, 12, 12, 12, 12, 12, 12, 14, 14])
    ws.cell(row=len(a_rows) + 3, column=1, value="출처: " + NYFED_SRC).font = NOTE_FONT
    ws.cell(row=len(a_rows) + 4, column=1,
            value="원표: Page 3(잔액), Page 4(계좌수, 공동명의 중복계상), Page 10(한도), Page 12·13·14(연체). 원 단위 $조 → $bn 환산.").font = NOTE_FONT
    write_csv(OUT / "A_heloc_balance_quarterly_2003Q1_2026Q3.csv",
              ["quarter", "heloc_balance_bn", "heloc_limit_bn", "heloc_available_bn", "utilization_pct", "heloc_accounts_mn",
               "mortgage_balance_bn", "total_household_debt_bn", "heloc_share_of_hh_debt_pct", "heloc_90plus_dq_pct",
               "heloc_flow_into_30plus_pct", "heloc_flow_into_90plus_pct"], a_rows)

    # B 연간
    ws = wb.create_sheet("B_잔액_연간")
    events = {2005: "2순위 신규취급 금액 정점 4Q05 (NY Fed SR569)",
              2007: "2순위(HELOC+CES) 잔액 $1.0T+ 정점(SR569)·서브프라임 붕괴·2순위 증권화 중단",
              2008: "리먼 파산(9월)", 2009: f"잔액 정점 {peak['quarter']}", 2021: "초저금리·현금인출 리파이 우위",
              2022: "연준 급격한 인상 → 1순위 금리 잠김", 2023: "2순위 RMBS 재개(Second lien 2.0)",
              2025: "Figure IPO(9월)", 2026: f"{last['quarter']} 최신치"}
    yrows = []
    prev = None
    for y in range(2004, last["year"] + 1):
        cand = [r for r in q if r["year"] == y]
        r = next((c for c in cand if c["quarter"].endswith("Q4")), cand[-1])
        yoy = round((r["heloc_bal_bn"] / prev - 1) * 100, 1) if prev else None
        prev = r["heloc_bal_bn"]
        yrows.append([y, r["quarter"], r["heloc_bal_bn"], yoy, r["heloc_limit_bn"], r["util_pct"], r["heloc_accts_mn"],
                      r["dq90_pct"], events.get(y, "")])
    # 2004 YoY: 2003Q4 대비
    q03 = next(r for r in q if r["quarter"] == "2003Q4")
    yrows[0][3] = round((yrows[0][2] / q03["heloc_bal_bn"] - 1) * 100, 1)
    b_hdr = ["연도", "기준분기", "HELOC 잔액($bn)", "YoY(%)", "HELOC 한도($bn)", "이용률(%)", "계좌수(백만)", "90일+ 연체(%)", "이벤트 주석"]
    write_table(ws, b_hdr, yrows, widths=[7, 9, 13, 9, 13, 10, 11, 11, 44])
    ws.cell(row=len(yrows) + 3, column=1, value="출처: " + NYFED_SRC + " 2026은 3Q(연말 아님) — YoY도 2025Q4 대비 3분기 변화.").font = NOTE_FONT
    write_csv(OUT / "B_heloc_balance_annual_2004_2026.csv",
              ["year", "ref_quarter", "heloc_balance_bn", "yoy_pct", "heloc_limit_bn", "utilization_pct", "accounts_mn",
               "dq90_pct", "event"], yrows)

    # C 증권화 연간
    ws = wb.create_sheet("C_증권화_연간")
    c_hdr = ["연도", "도표용 값($bn)", "HELOC($bn)", "CES($bn)", "합계($bn)", "딜 수", "등급", "출처키", "비고"]
    c_rows = [[d["year"], d["chart_bn"], d["heloc"], d["ces"], d["total"], d["deals"], d["grade"], d["src"], d["note"]]
              for d in SEC_ANNUAL]
    write_table(ws, c_hdr, c_rows, widths=[7, 12, 10, 10, 10, 7, 12, 22, 110])
    n = len(c_rows) + 3
    for t in ["등급: 확정=원문 수치 그대로 / 파생=원문 수치로 계산 / 부분=분기·반기 단편만 / 미확보(휴면)=공개 정량치 없음 / YTD=연중 누적",
              "도표용 값: 2004~2005 Inside Mortgage Finance, 2022 KBRA, 2023~2024 HEL News, 2025 KBRA, 2026 J.P. Morgan YTD. 집계처가 섞여 있으니 각주 필수."]:
        ws.cell(row=n, column=1, value=t).font = NOTE_FONT
        n += 1
    write_csv(OUT / "C_second_lien_rmbs_issuance_annual_2004_2026.csv",
              ["year", "chart_value_bn", "heloc_bn", "ces_bn", "total_bn", "deals", "grade", "source_key", "note"], c_rows)

    # D 위기 전 분기
    ws = wb.create_sheet("D_위기전_분기")
    write_table(ws, ["기간", "구분", "발행액($bn)", "딜 수", "등급", "출처키", "비고"], PRE_CRISIS_Q,
                widths=[9, 22, 11, 7, 8, 10, 70])

    # E 2026 YTD
    ws = wb.create_sheet("E_2026_YTD")
    write_table(ws, ["기준일", "집계처", "누적/전망($bn)", "구분", "비고", "출처키"], YTD_2026, widths=[13, 24, 14, 7, 60, 12])

    # F 잔액·발행사
    ws = wb.create_sheet("F_증권화잔액·발행사")
    write_table(ws, ["항목", "값", "출처키"], OUTSTANDING_SEC, widths=[28, 60, 12])
    r0 = len(OUTSTANDING_SEC) + 3
    ws.cell(row=r0, column=1, value="HELOC RMBS 잔액 shelf별 점유(%)").font = Font(bold=True)
    for i, (k, v) in enumerate(SHELF_HELOC, r0 + 1):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
    r1 = r0 + len(SHELF_HELOC) + 2
    ws.cell(row=r1, column=1, value="CES RMBS 잔액 shelf별 점유(%)").font = Font(bold=True)
    for i, (k, v) in enumerate(SHELF_CES, r1 + 1):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
    ws.cell(row=r1 + len(SHELF_CES) + 2, column=1,
            value="출처: SFA Research Corner 2026-09-17 (Deutsche Bank, Intex). 시점은 2026년 9월 게재 기준.").font = NOTE_FONT

    # G 검증
    ws = wb.create_sheet("G_붙여넣은표_검증")
    write_table(ws, ["받은 표의 주장", "판정", "원문 확인 내용", "출처키"], PASTED_CHECK, widths=[40, 20, 90, 18])

    # SOURCES
    ws = wb.create_sheet("SOURCES")
    write_table(ws, ["키", "원문"], [["NYFED", NYFED_SRC]] + [[k, v] for k, v in S.items()], widths=[12, 160])

    xlsx = OUT / "HELOC_2004_2026_rawdata.xlsx"
    wb.save(xlsx)
    print("saved", xlsx)
    print("peak", peak["quarter"], peak["heloc_bal_bn"], "trough", trough["quarter"], trough["heloc_bal_bn"],
          "last", last["quarter"], last["heloc_bal_bn"], "util", last["util_pct"])


if __name__ == "__main__":
    main()
