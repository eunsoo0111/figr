import json, csv, glob, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

BASE = sys.argv[1]; OUT = sys.argv[2]; CSV_OUT = sys.argv[3]
YEARS = list(range(2018, 2026))

# partner group, display, LEIs, join, source, set
P = [
 # group, LEIs, join, evidence, url, role
 ("Guaranteed Rate (Rate + Affinity)", ["549300U3721PJGQZYY68","54930001NSTOD85LT125"], "2022-07",
  "2022.7 Figure 보도자료(Guaranteed Rate·Homebridge·Synergy One 화이트라벨 제휴); DBRS FIGRE 2024-HE1 화이트라벨 실행사",
  "https://www.prnewswire.com/news-releases/figure-expands-reach-of-heloc-products-to-consumers-through-partnerships-with-leading-mortgage-companies-301595835.html", "partner"),
 ("Homebridge Financial Services", ["5493001WHVQBGRSWEU75"], "2022-07",
  "2022.7 Figure 보도자료; DBRS FIGRE 2023-HE1·2024-HE1 화이트라벨 실행사",
  "https://dbrs.morningstar.com/research/429315/morningstar-dbrs-assigns-provisional-credit-ratings-to-figre-trust-2024-he1", "partner"),
 ("Synergy One Lending", ["2549004FXPNKCLIN8250"], "2022-07",
  "2022.7 Figure 보도자료; DBRS FIGRE 2024-HE1 화이트라벨 실행사",
  "https://dbrs.morningstar.com/research/429315/morningstar-dbrs-assigns-provisional-credit-ratings-to-figre-trust-2024-he1", "partner"),
 ("Movement Mortgage", ["549300DD4R4SYK5RAQ92"], "2022-11",
  "2022.11 Movement-Figure 제휴 보도자료; DBRS FIGRE 2024-HE1 화이트라벨 실행사",
  "https://prnewswire.com/news-releases/movement-mortgage-partners-with-figure-to-offer-improved-heloc-delivery-301671054.html", "partner"),
 ("CMG Financial", ["549300GKFNPRWNS0GF29"], "2023-07",
  "2023.7 Figure 보도자료(CMG·CrossCountry·Fairway·The Loan Store). 단 2018년부터 자체 HELOC 실적 보유",
  "https://figure.com/newsroom/figure-expands-reach-of-heloc-products-through-four-new-partnerships/", "partner_preexisting"),
 ("CrossCountry Mortgage", ["549300VZVN841I2ILS84"], "2023-07",
  "2023.7 Figure 보도자료",
  "https://www.inman.com/2023/07/12/figure-powering-helocs-for-4-more-independent-mortgage-banks/", "partner"),
 ("Fairway Independent Mortgage", ["549300MGPZBLQDIL7538"], "2023-07",
  "2023.7 Figure 보도자료",
  "https://www.inman.com/2023/07/12/figure-powering-helocs-for-4-more-independent-mortgage-banks/", "partner"),
 ("The Loan Store", ["254900DTLHVWQ7NP7R34"], "2023-07",
  "2023.7 Figure 보도자료; 2024.10 피기백 HELOC 첫 도입사",
  "https://www.inman.com/2023/07/12/figure-powering-helocs-for-4-more-independent-mortgage-banks/", "partner"),
 ("New American Funding", ["549300E2UX99HKDBR481"], "<=2024-03",
  "DBRS FIGRE 2024-HE1(2024.3) 화이트라벨 실행사로 처음 확인",
  "https://dbrs.morningstar.com/research/429315/morningstar-dbrs-assigns-provisional-credit-ratings-to-figre-trust-2024-he1", "partner"),
 ("Home Point Financial", ["549300BRJZYHYKT4BJ84"], "<=2022",
  "DBRS FIGRE 2024-HE1 화이트라벨 실행사. 2023년 영업 종료(HMDA 2023~ 미제출) - 생존편향 방지 위해 포함",
  "https://dbrs.morningstar.com/research/429315/morningstar-dbrs-assigns-provisional-credit-ratings-to-figre-trust-2024-he1", "partner"),
 ("Union Home Mortgage", ["549300RPOGWJRH63HS39"], "<=2024-10",
  "2024.10 Figure 피기백 HELOC 출시 보도자료(첫 도입사)",
  "https://www.businesswire.com/news/home/20241028579758/en/", "partner"),
 # reference rows (not in partner aggregate)
 ("[참고] loanDepot", ["549300AG64NHILB7ZP05"], "2026-04",
  "사용자 문단 기준 2026.4 도입. HMDA 2024~25 HELOC은 Figure 도입 전 자체 실적 -> 합산 제외, '25년 연간 실적' 기준치로만 사용", "", "ref"),
 ("[참고] Lakeview Loan Servicing", ["549300NOCASXPA34X033"], "2025~26",
  "사용자 문단 기준 '실적 없던' 신규 파트너. HMDA 2018~25 HELOC(open-end) 0건. 2025년분은 보고기준 미달로 미보고 가능성", "", "ref"),
 ("[참고] Figure Lending (자사 브랜드)", ["254900UL88QFG0E40516"], "-",
  "Figure 자사 명의 실행분(파트너 아님). 비교용", "", "ref"),
]

summ = {}
for f in glob.glob(f"{BASE}/out/sum_*.json"):
    s = json.load(open(f)); summ[(s["lei"], s["year"])] = s
lei_name = {}
for l in open(f"{BASE}/candidates.tsv"):
    a, b = l.rstrip("\n").split("\t"); lei_name[a] = b

thin = Side(style="thin", color="BFBFBF")
HDR = PatternFill("solid", fgColor="1F3864"); HFONT = Font(name="Arial", bold=True, color="FFFFFF", size=9)
F = Font(name="Arial", size=9); FB = Font(name="Arial", size=9, bold=True)
FIN = Font(name="Arial", size=9, color="0000FF")  # hardcoded inputs
FLINK = Font(name="Arial", size=9, color="008000")
SUB = PatternFill("solid", fgColor="D9E1F2"); KEY = PatternFill("solid", fgColor="FFF2CC")

def hdr(ws, row, vals, col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=col + i, value=v); c.font = HFONT; c.fill = HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()
# ---------------- README ----------------
ws = wb.active; ws.title = "README"
lines = [
 ("Figure(FIGR) 주요 파트너사 합산 HELOC 실행액 추이 - raw data", FB),
 ("작성: 2026-11-19 / 출처: FFIEC·CFPB HMDA Data Browser API (loan-level, 2018~2025 최종·스냅샷)", F),
 ("", F),
 ("[무엇을 셌나]", FB),
 ("· 대상: 아래 'Partners' 탭의 대출기관 LEI별, 실행 완료(action_taken=1) 건 중 HELOC", F),
 ("· HELOC 정의: open-end_line_of_credit=1 이고 사업목적(business_or_commercial_purpose=1)·역모기지(reverse_mortgage=1) 제외. 1순위·2순위 담보 모두 포함(열로 분리)", F),
 ("· 금액: HMDA loan_amount = 승인된 한도액. 공개 데이터는 $10,000 구간의 중간값으로 반올림되어 있어 합계는 근사치", F),
 ("· 파트너 명의(화이트라벨)로 실행된 HELOC은 HMDA에 파트너 LEI로 보고됨 -> 파트너별 실행액을 공시 없이 추적 가능 (Figure의 'Figure Factor' 산출도 HMDA 기반)", F),
 ("", F),
 ("[파트너 선정 규칙]", FB),
 ("· Figure 보도자료·DBRS FIGRE 신용평가 보고서에 Figure HELOC 화이트라벨 파트너로 이름이 나온 곳 중 HMDA 보고기관, 2024년까지 합류한 곳(2025년이 합류 후 온전한 1년이 되도록)", F),
 ("· Set A = 확인된 파트너 11개 그룹 전부 / Set B = Set A에서 CMG 제외(CMG는 Figure 도입 전인 2018년부터 자체 HELOC 실적 보유 -> '없던 물량 창출' 논지의 비교군으로 부적합)", F),
 ("· loanDepot(2026.4 도입)·Lakeview는 2025년까지 HMDA에 Figure 효과가 잡히지 않아 합산 제외, 참고행으로만 표시", F),
 ("· 사용자 문단의 '14곳' 원명단은 공개 자료로 확인 불가 -> 공개 확인 가능한 범위로 재구성", F),
 ("", F),
 ("[해석 시 주의]", FB),
 ("· HMDA 보고 기준: 직전 2개년 각각 open-end 200건 이상 실행 기관만 open-end 보고 의무(2022~). 0 = '실적 없음 또는 200건 미만으로 미보고'", F),
 ("  -> 2022년 0은 '없었거나 극히 미미'로 읽는 것이 정확. 합류 직후 1~2년도 미보고로 0일 수 있음(예: CrossCountry 2023.7 합류, HMDA 2023~24 0 -> 2025 4,466건)", F),
 ("  -> 따라서 2023~24년 값은 과소집계, 시작점(2022)과 끝점(2025) 비교가 가장 깨끗함", F),
 ("· HMDA는 연간 데이터만 존재(월·분기 없음). 2026년분은 2027년 공개 예정", F),
 ("· Synergy One은 2025년 HMDA에 open-end 0건(총 실행 8,115건은 보고) - 원자료 그대로 둠", F),
 ("", F),
 ("[탭 구성]", FB),
 ("· Partners: 파트너 목록·LEI·합류시점·근거 출처", F),
 ("· Data_long: LEI×연도 원자료(건수·금액·담보순위별) - 모든 집계의 원천", F),
 ("· Agg: 파트너별 연도별 HELOC 실행액·건수(Data_long에서 SUMIFS) + Set A/B 합산·배수·2022년 0인 파트너 수", F),
 ("· 별도 CSV(hmda_heloc_loanlevel_partners.csv): 위 기관들의 HELOC 건별 원장(주요 컬럼만)", F),
 ("", F),
 ("API 예시: https://ffiec.cfpb.gov/v2/data-browser-api/view/csv?years=2025&leis=549300VZVN841I2ILS84&actions_taken=1", F),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = f
ws.column_dimensions["A"].width = 150

# ---------------- Partners ----------------
wp = wb.create_sheet("Partners")
hdr(wp, 1, ["#", "파트너(그룹)", "HMDA LEI", "HMDA 등록명", "Figure 합류(공개 확인 시점)", "Set A", "Set B", "근거", "출처 URL"])
r = 2; n = 0
for grp, leis, join, ev, url, role in P:
    for lei in leis:
        n += 1
        vals = [n, grp, lei, lei_name.get(lei, ""), join,
                "Y" if role.startswith("partner") else "N",
                "Y" if role == "partner" else "N", ev, url]
        for j, v in enumerate(vals, 1):
            c = wp.cell(row=r, column=j, value=v); c.font = F; c.alignment = Alignment(vertical="top", wrap_text=(j == 8))
        r += 1
for col, w in zip("ABCDEFGHI", [4, 32, 24, 34, 16, 7, 7, 70, 60]):
    wp.column_dimensions[col].width = w
wp.freeze_panes = "C2"

# ---------------- Data_long ----------------
wd = wb.create_sheet("Data_long")
cols = ["파트너(그룹)", "LEI", "연도", "Set A", "Set B", "HELOC 건수", "HELOC 실행액($)",
        "1순위 HELOC 건수", "1순위 HELOC 실행액($)", "2순위 HELOC 건수", "2순위 HELOC 실행액($)",
        "전체 실행 건수(전 상품)", "전체 실행액($, 전 상품)", "[참고] 폐쇄형 2순위 건수", "[참고] 폐쇄형 2순위 실행액($)"]
hdr(wd, 1, cols)
r = 2
for grp, leis, join, ev, url, role in P:
    for lei in leis:
        for y in YEARS:
            s = summ[(lei, y)]
            vals = [grp, lei, str(y), "Y" if role.startswith("partner") else "N", "Y" if role == "partner" else "N",
                    s["heloc_n"], s["heloc_amt"], s["heloc_1st_n"], s["heloc_1st_amt"], s["heloc_2nd_n"], s["heloc_2nd_amt"],
                    s["all_orig_n"], s["all_orig_amt"], s["closed_2nd_n"], s["closed_2nd_amt"]]
            for j, v in enumerate(vals, 1):
                c = wd.cell(row=r, column=j, value=v); c.font = FIN if j >= 6 else F
                if j >= 6: c.number_format = '#,##0;(#,##0);-'
            r += 1
DL_LAST = r - 1
for j, w in enumerate([32, 24, 6, 6, 6, 10, 16, 10, 16, 10, 16, 12, 18, 12, 16], 1):
    wd.column_dimensions[get_column_letter(j)].width = w
wd.freeze_panes = "D2"
wd.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{DL_LAST}"

# ---------------- Agg ----------------
wa = wb.create_sheet("Agg")
YC = {y: 3 + i for i, y in enumerate(YEARS)}  # year -> column index (C..J)
def rng(col): return f"Data_long!${col}$2:${col}${DL_LAST}"
A_GRP, A_YR, A_SETA, A_SETB = rng("A"), rng("C"), rng("D"), rng("E")

def block(top, title, valcol, unit_div, fmt, unit_label):
    wa.cell(row=top, column=1, value=title).font = Font(name="Arial", size=10, bold=True)
    hdr(wa, top + 1, ["파트너(그룹)", "Set"] + [str(y) for y in YEARS] + ["25/22 배수"])
    r = top + 2; first = r
    groups = [(g, role) for g, _, _, _, _, role in P]
    part_rows = []
    for g, role in groups:
        if role == "ref" and g.startswith("[참고]"):
            continue
        setlab = "A,B" if role == "partner" else "A only"
        wa.cell(row=r, column=1, value=g).font = F
        wa.cell(row=r, column=2, value=setlab).font = F
        for y, c in YC.items():
            f = f'=SUMIFS({rng(valcol)},{A_GRP},$A{r},{A_YR},"{y}")/{unit_div}'
            cell = wa.cell(row=r, column=c, value=f); cell.font = FLINK; cell.number_format = fmt
        m = wa.cell(row=r, column=11, value=f'=IF({get_column_letter(YC[2022])}{r}=0,"n/a (22년 0)",{get_column_letter(YC[2025])}{r}/{get_column_letter(YC[2022])}{r})')
        m.font = F; m.number_format = '0.0"x"'
        part_rows.append(r); r += 1
    last = r - 1
    # aggregates
    for lab, setcol in (("합산 Set A (확인 파트너 전부)", A_SETA), ("합산 Set B (CMG 제외)", A_SETB)):
        wa.cell(row=r, column=1, value=lab).font = FB
        for y, c in YC.items():
            f = f'=SUMIFS({rng(valcol)},{setcol},"Y",{A_YR},"{y}")/{unit_div}'
            cell = wa.cell(row=r, column=c, value=f); cell.font = FB; cell.number_format = fmt; cell.fill = KEY
        cl22, cl25 = get_column_letter(YC[2022]), get_column_letter(YC[2025])
        m = wa.cell(row=r, column=11, value=f'=IF({cl22}{r}=0,"n/a",{cl25}{r}/{cl22}{r})'); m.font = FB; m.number_format = '0.0"x"'; m.fill = KEY
        r += 1
    # reference rows
    for g, _, _, _, _, role in P:
        if not g.startswith("[참고]"): continue
        wa.cell(row=r, column=1, value=g).font = Font(name="Arial", size=9, italic=True, color="7F7F7F")
        for y, c in YC.items():
            f = f'=SUMIFS({rng(valcol)},{A_GRP},$A{r},{A_YR},"{y}")/{unit_div}'
            cell = wa.cell(row=r, column=c, value=f); cell.font = Font(name="Arial", size=9, italic=True, color="7F7F7F"); cell.number_format = fmt
        r += 1
    wa.cell(row=r, column=1, value=f"단위: {unit_label}. 파트너 행·합산 행 모두 Data_long에서 SUMIFS로 집계").font = Font(name="Arial", size=8, italic=True)
    return first, last, r + 2

amt_first, amt_last, nxt = block(1, "① 파트너별 HELOC 실행액 (HMDA, $M)", "G", 1000000, '#,##0;(#,##0);-', "$M (백만 달러)")
cnt_first, cnt_last, nxt2 = block(nxt, "② 파트너별 HELOC 실행 건수 (HMDA, 건)", "F", 1, '#,##0;(#,##0);-', "건")

# key stats
r = nxt2
wa.cell(row=r, column=1, value="③ 문단용 핵심 수치").font = Font(name="Arial", size=10, bold=True); r += 1
hdr(wa, r, ["지표", "Set A", "Set B", "비고"]); r += 1
cl22, cl25 = get_column_letter(YC[2022]), get_column_letter(YC[2025])
setA_amt = amt_last + 1; setB_amt = amt_last + 2
setA_cnt = cnt_last + 1; setB_cnt = cnt_last + 2
rngA_amt22 = f"{cl22}{amt_first}:{cl22}{amt_last}"
stats = [
 ("파트너 수(그룹)", f'=COUNTA(A{amt_first}:A{amt_last})', f'=COUNTIF(B{amt_first}:B{amt_last},"A,B")', "Guaranteed Rate+Affinity는 1개 그룹"),
 ("2022년 HELOC 실적 0인 파트너 수", f'=COUNTIF({rngA_amt22},0)', f'=COUNTIFS({rngA_amt22},0,B{amt_first}:B{amt_last},"A,B")', "0 = 실적 없음 또는 보고기준(200건) 미달"),
 ("2022년 합산 실행액($M)", f"={cl22}{setA_amt}", f"={cl22}{setB_amt}", ""),
 ("2025년 합산 실행액($M)", f"={cl25}{setA_amt}", f"={cl25}{setB_amt}", ""),
 ("2025/2022 배수(금액)", f"={cl25}{setA_amt}/{cl22}{setA_amt}", f"={cl25}{setB_amt}/{cl22}{setB_amt}", "3년간"),
 ("2025/2022 배수(건수)", f"={cl25}{setA_cnt}/{cl22}{setA_cnt}", f"={cl25}{setB_cnt}/{cl22}{setB_cnt}", ""),
 ("[참고] loanDepot 2025년 HELOC 실행액($M)", f"={cl25}{amt_last+3}", "", "'26.4 도입 후 7개월간 이 값의 3배라는 문단 주장의 기준치(HMDA 2025)"),
]
for lab, fa, fb, note in stats:
    wa.cell(row=r, column=1, value=lab).font = F
    for j, f in ((2, fa), (3, fb)):
        c = wa.cell(row=r, column=j, value=f if f else None); c.font = FB; c.fill = KEY
        c.number_format = '0.0"x"' if "배수" in lab else '#,##0'
    wa.cell(row=r, column=4, value=note).font = F
    r += 1
wa.column_dimensions["A"].width = 40; wa.column_dimensions["B"].width = 9
for c in range(3, 12): wa.column_dimensions[get_column_letter(c)].width = 10
wa.column_dimensions["D"].width = 10
wa.freeze_panes = "C3"
wb.move_sheet("Agg", offset=-2)
wb.save(OUT)

# loan-level CSV
keep_leis = {lei for g, leis, *_ in P for lei in leis if "Figure Lending" not in g}
grp_of = {lei: g for g, leis, *_ in P for lei in leis}
with open(CSV_OUT, "w", newline="") as fo:
    w = None
    for lei in sorted(keep_leis):
        for y in YEARS:
            fn = f"{BASE}/out/heloc_{lei}_{y}.csv"
            if not os.path.exists(fn): continue
            for row in csv.DictReader(open(fn)):
                row = {"partner_group": grp_of[lei], **row}
                if w is None:
                    w = csv.DictWriter(fo, fieldnames=list(row)); w.writeheader()
                w.writerow(row)
print("ok", DL_LAST)
