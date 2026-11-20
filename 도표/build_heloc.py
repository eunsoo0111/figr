# build_heloc.py — 도표 1-1: 동사 합류 파트너사 합산 HELOC 실행액 추이 (누적 영역 + 원그래프 삽입)
# 1) hmda_raw/heloc_by_lei_<연도>.csv(HMDA 공개 LAR 집계) + hmda/partners.csv → 원자료 엑셀(유일한 원본)
# 2) 엑셀을 load()로 읽어 SMIC 양식 네이티브 PPT 도표
#   cd /home/user/figr/도표 && python3 build_heloc.py && python3 smic_check.py SMIC_도표_HELOC.pptx SMIC_도표_원자료_HELOC.xlsx --render render
import csv, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from smic_xlsx import SourceBook, load  # noqa: E402
from smic_charts_lib import *  # noqa: E402,F401

NO, TITLE = "도표 1-1", "동사 합류 파트너사 합산 HELOC 실행액 추이"
PIE = NO + " 삽입"
RAW = os.path.join(HERE, "hmda_raw")
YEARS_ALL = list(range(2018, 2026))                # 원자료(참고 시트)
YEARS = list(range(2020, 2026))                    # 도표: 첫 합류('22.7) 2년 전부터
BASE = 2022                                        # 배수 기준 연도(첫 합류 연도)
TOP = 5                                            # 개별 계열 수(+기타 = 6색)
XLSX = os.path.join(HERE, "SMIC_도표_원자료_HELOC.xlsx")
PPTX = os.path.join(HERE, "SMIC_도표_HELOC.pptx")
API = "https://ffiec.cfpb.gov/v2/data-browser-api/view/nationwide/csv?years={y}&actions_taken=1"
FIGURE_LEI = "254900UL88QFG0E40516"                # Figure Lending LLC(동사 자체 명의, 참고)

# ── 원자료 읽기: 후순위(lien_status=2) 한도형 여신 실행 약정액 ──
parts = list(csv.DictReader(open(os.path.join(HERE, "hmda", "partners.csv"), encoding="utf-8-sig")))
used = [p for p in parts if p["도표"] == "포함"]
lar = {y: {r["LEI"]: r for r in csv.DictReader(open(os.path.join(RAW, f"heloc_by_lei_{y}.csv"), encoding="utf-8"))}
       for y in YEARS_ALL}


def sub(leis, y, col="후순위실행액"):
    return sum(int(lar[y][l][col]) for l in leis.split(";") if l in lar[y])


market = {y: sum(int(r["후순위실행액"]) for r in lar[y].values()) for y in YEARS_ALL}
vol = {p["약칭"]: [sub(p["LEI"], y) for y in YEARS] for p in used}
order = sorted(vol, key=lambda k: -vol[k][-1])     # 2025 실행액 큰 순 = 아래부터 쌓는 순서·원그래프 순서
top, rest = order[:TOP], order[TOP:]
ETC = f"기타 {len(rest)}개사"
SUM = "파트너 합계"
MKT = "후순위 HELOC 시장 전체"
k = YEARS.index(2025) - YEARS.index(BASE)          # 2025행에서 기준 연도까지 행 수
MUL, MKG = f"파트너 '{BASE % 100}→'25 배수", f"시장 '{BASE % 100}→'25 증감"

# ── 1) 원자료 엑셀 ──
wb = SourceBook()
cols = [(n, vol[n], FMT["백만"]) for n in top]
cols += [(n, vol[n], FMT["백만"], "hide") for n in rest]
cols += [(ETC, "=" + "+".join("{%s}" % n for n in rest), FMT["백만"])]
cols += [(SUM, "=" + "+".join("{%s}" % n for n in top + [ETC]), FMT["백만"], "hide"),
         (MKT, [market[y] for y in YEARS], FMT["백만"], "hide"),
         ("파트너 비중", "={%s}/{%s}" % (SUM, MKT), FMT["%"], "hide"),
         (MUL, "={%s}/{%s@-%d}" % (SUM, SUM, k), FMT["배"].replace('"x"', '"배"'), "hide"),
         (MKG, "={%s}/{%s@-%d}-1" % (MKT, MKT, k), '+0.0%;-0.0%', "hide")]
src = [f"HMDA 공개 LAR {y}(FFIEC) 실행·한도형 여신·후순위 loan_amount 합 — {API.format(y=y)}" for y in YEARS]
ws = wb.add(NO, TITLE, [str(y) for y in YEARS], cols, row_sources=src,
            source="HMDA 공개 LAR(FFIEC Data Browser), 파트너·합류 시점은 '파트너 목록' 시트",
            unit="(백만 달러)",
            note=("HELOC = open-end_line_of_credit=1 & lien_status=2(후순위), 가계 목적, 역모기지 제외 실행 건의 약정 한도 합(파트너 명의 실행분). "
                  "한도형 여신 HMDA 보고 의무는 직전 2년 연속 연 500건('21년까지)·200건('22~) 이상 실행 기관 — 합류 초기 연도는 과소 집계. "
                  "loan_amount는 1만 달러 구간 중간값."),
            check="합류 시점 미공개: GoodLeap. 보고 의무 전이라 0: Lakeview('25.9 합류, 도표 제외)")
# 원그래프: 2025년 값은 도표 시트를 수식으로 참조(숫자를 두 번 치지 않는다)
letter = {c[0]: chr(ord("B") + i) for i, c in enumerate(cols)}
last = len(YEARS) + 1
for n in (MUL, MKG):                                 # 배수·증감은 2025행만(다른 행은 열 이름과 다른 연도 쌍이 된다)
    for r in range(2, last):
        ws[f"{letter[n]}{r}"].value = None
pie_names = top + [ETC]
ws2 = wb.add(PIE, "2025년 파트너사별 후순위 HELOC 실행액 비중", pie_names,
             [("2025 실행액", [0] * len(pie_names), FMT["백만"]),
              ("비중", "={2025 실행액}/SUM($B$2:$B$%d)" % (len(pie_names) + 1), FMT["%"], "hide")],
             first="구분", source=f"'{NO}' 시트 2025년 행 참조", unit="(백만 달러)", size="삽입")
for i, n in enumerate(pie_names):
    ws2.cell(i + 2, 2).value = f"='{NO}'!{letter[n]}{last}"
wb.save(XLSX)

# 참고 시트 2개(마지막 열이 '출처'가 아니므로 load()·대조 대상이 아니다)
import openpyxl  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402
xb = openpyxl.load_workbook(XLSX)


def head(ws_, names, widths):
    ws_.append(names)
    for j, c in enumerate(ws_[1], 0):
        c.font = Font(name="AppleSDGothicNeoB00", color="FFFFFF"); c.fill = PatternFill("solid", fgColor="595959")
        ws_.column_dimensions[chr(65 + j)].width = widths[j]
    ws_.freeze_panes = "A2"


pw = xb.create_sheet("파트너 목록")
head(pw, ["약칭", "파트너", "LEI", "합류", "근거", "URL", "도표 계열"], (14, 44, 44, 12, 50, 60, 40))
for p in parts:
    pw.append([p["약칭"], p["파트너"], p["LEI"], p["합류"], p["근거"], p["URL"],
               (p["약칭"] if p["약칭"] in top else ETC) if p["도표"] == "포함" else p["도표"]])
hw = xb.create_sheet("HMDA 집계(참고)")
head(hw, ["구분", "LEI", "연도", "HELOC 건수", "HELOC 실행액(전체)", "1순위 실행액", "후순위 건수", "후순위 실행액", "API"],
     (30, 44, 8, 12, 20, 18, 12, 18, 90))
rows_ref = [(p["약칭"], p["LEI"]) for p in parts] + [("(참고) Figure Lending LLC — 동사 자체 명의", FIGURE_LEI)]
for name, leis in rows_ref:
    for y in YEARS_ALL:
        hw.append([name, leis, y, sub(leis, y, "HELOC건수"), sub(leis, y, "HELOC실행액"), sub(leis, y, "1순위실행액"),
                   sub(leis, y, "후순위건수"), sub(leis, y, "후순위실행액"), API.format(y=y)])
for y in YEARS_ALL:
    hw.append(["(시장) HMDA 보고 전 기관 합계", "", y, sum(int(r["HELOC건수"]) for r in lar[y].values()),
               sum(int(r["HELOC실행액"]) for r in lar[y].values()), sum(int(r["1순위실행액"]) for r in lar[y].values()),
               sum(int(r["후순위건수"]) for r in lar[y].values()), market[y], API.format(y=y)])
for r in hw.iter_rows(min_row=2):
    for c in r[3:8]:
        c.number_format = "#,##0"; c.font = Font(name="Arial")
xb.calculation.fullCalcOnLoad = True
xb.save(XLSX)

# ── 2) PPT 도표: 엑셀에서 읽은 값으로만 ──
bk = load(XLSX)
sh, pi = bk[NO], bk[PIE]
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)", [f"{NO}. {TITLE}"],
            checks=["파트너 목록·합류 시점(엑셀 '파트너 목록' 시트)", "HMDA는 연간 공시 — 2025년이 최신(2026.6 공개)",
                    "합류 초기 연도는 HMDA 보고 의무 기준 때문에 과소 집계 → 배수는 상한으로 읽기"])
cv = chart_slide(prs, NO, TITLE, "기본", "HMDA(FFIEC), 각사, SMIC 3팀",
                 notes=("원자료: SMIC_도표_원자료_HELOC.xlsx / 산출: hmda/hmda_heloc.py — HMDA 공개 LAR 연도별 전국 실행분 → LEI별 집계. "
                        "후순위 한도형 여신(lien_status=2) 약정 한도 합. 파트너 명의(LEI)로 보고된 실행분만 잡힌다."),
                 note_lines=[f"동사 파트너 {len(used)}개사의 HMDA 후순위 한도형 여신(HELOC) 실행 약정액 합(파트너 명의 실행분). Lakeview('25.9 합류)는 '25 보고 의무 전이라 제외",
                             "HMDA 한도형 여신 보고 의무는 직전 2년 연속 연 500건('21년까지)·200건('22~) 이상 기관 → 합류 초기 연도는 과소 집계"])
names = top + [ETC]
tot = [sum(sh.cols[n][i] for n in names) for i in range(len(YEARS))]          # 누적 합계로 세로축
p = add_chart(cv, "area", sh.cats, sh.series(*names), sheet=NO, fmt=sh.fmt(top[0]),
              units=("(백만 달러)", ""), on_tick=True, yscale=nice_scale(tot, head=0.12, steps=(1, 2, 5, 10)))
# 원그래프 삽입: 영역이 0에 가까운 왼쪽 위
d = 1.95
px, py = p.px0 + 0.1, p.top + 0.42
pp = add_chart(cv, "pie", pi.cats, pi.series("2025 실행액"), sheet=PIE, fmt=pi.fmt("2025 실행액"),
               x=px, y=py, w=d, h=d, legend=False)
plot_inner(pp.ch, 0.0, 0.0, 1.0, 1.0)                                        # 원을 칸 가득
cv.text(px - 0.3, py + d + 0.02, d + 0.6, 0.62, ["2025년 기준", "파트너사별 비중"], CT, DG, "c")
# 추세 화살표 + 배수(엑셀 수식 값) — 기준 연도 → 2025
i0, i1 = YEARS.index(BASE), YEARS.index(2025)
y0, y1 = p.y(tot[i0]) - 0.22, p.y(tot[i1]) - 0.18
cv.arrow(p.x(i0) + 0.05, y0, p.x(i1) - 0.12, y1, color=PINK2, lw=1.0)
tx = p.x(i0) - 0.35                                                          # 원그래프 오른쪽, 화살표 왼쪽 위
cv.val(tx, p.top + 0.32, sh, MUL, i1, pre=f"'{BASE % 100}→'25 ", color=RED, bold=True)
cv.val(tx, p.top + 0.68, sh, MKG, i1, pre="(같은 기간 후순위|HELOC 시장 ", post=")", color=DG, wrap_at="|")
prs.save(PPTX)
print("saved", XLSX, PPTX)
print("순서:", order)
for n in order:
    print(f"  {n:<16}", " ".join(f"{v/1e9:6.3f}" for v in vol[n]))
print("  시장           ", " ".join(f"{market[y]/1e9:6.1f}" for y in YEARS))
