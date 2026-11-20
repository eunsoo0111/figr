# build_heloc.py — 도표: 동사 합류 파트너사 합산 HELOC 실행액 추이 (누적 영역 + 원그래프 삽입)
# 1) hmda_raw/heloc_by_lei_<연도>.csv(HMDA 공개 LAR 집계) + hmda/partners.csv → 원자료 엑셀(유일한 원본)
# 2) 엑셀을 load()로 읽어 SMIC 양식 네이티브 PPT 도표
#   python3 build_heloc.py            (작업 폴더: /home/user/figr/도표)
import csv, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from smic_xlsx import SourceBook, load  # noqa: E402
from smic_charts_lib import *  # noqa: E402,F401

NO, TITLE = "도표 1-1", "동사 합류 파트너사 합산 HELOC 실행액 추이"
PIE = NO + " 삽입"
RAW = os.path.join(HERE, "hmda_raw")
YEARS = list(range(2018, 2026))
TOP = 5                                            # 개별 계열 수(+기타 = 6색)
XLSX = os.path.join(HERE, "SMIC_도표_원자료_HELOC.xlsx")
PPTX = os.path.join(HERE, "SMIC_도표_HELOC.pptx")
API = "https://ffiec.cfpb.gov/v2/data-browser-api/view/nationwide/csv?years={y}&actions_taken=1"

# ── 원자료 읽기 ──
parts = list(csv.DictReader(open(os.path.join(HERE, "hmda", "partners.csv"), encoding="utf-8-sig")))
by_lei, market = {}, {}
for y in YEARS:
    for r in csv.DictReader(open(os.path.join(RAW, f"heloc_by_lei_{y}.csv"), encoding="utf-8")):
        a = int(r["HELOC실행액"])
        by_lei[(r["LEI"], y)] = a
        market[y] = market.get(y, 0) + a
vol = {p["약칭"]: [sum(by_lei.get((l, y), 0) for l in p["LEI"].split(";")) for y in YEARS] for p in parts}
order = sorted(vol, key=lambda k: -vol[k][-1])     # 2025 실행액 큰 순 = 아래부터 쌓는 순서·원그래프 순서
top, rest = order[:TOP], order[TOP:]
ETC = f"기타 {len(rest)}개사"
SUM = "파트너 합계"

# ── 1) 원자료 엑셀 ──
wb = SourceBook()
cols = [(k, vol[k], FMT["십억"]) for k in top]
cols += [(k, vol[k], FMT["십억"], "hide") for k in rest]
cols += [(ETC, "=" + "+".join("{%s}" % k for k in rest), FMT["십억"])]
cols += [(SUM, "=" + "+".join("{%s}" % k for k in top + [ETC]), FMT["십억"], "hide"),
         ("HELOC 시장 전체", [market[y] for y in YEARS], FMT["십억"], "hide"),
         ("파트너 비중", "={%s}/{HELOC 시장 전체}" % SUM, FMT["%"], "hide"),
         ("파트너 합계 YoY", "={%s}/{%s@-1}-1" % (SUM, SUM), FMT["%"], "hide"),
         ("시장 YoY", "={HELOC 시장 전체}/{HELOC 시장 전체@-1}-1", FMT["%"], "hide")]
src = [f"HMDA 공개 LAR {y}(FFIEC Data Browser) 실행·한도형 여신 loan_amount 합 — {API.format(y=y)}" for y in YEARS]
ws = wb.add(NO, TITLE, [str(y) for y in YEARS], cols, row_sources=src,
            source="HMDA 공개 LAR(FFIEC), 파트너·합류 시점은 hmda/partners.csv 근거",
            unit="(십억 달러)",
            note=("HELOC = open-end_line_of_credit=1, 가계 목적, 역모기지 제외 실행 건의 약정 한도 합. "
                  "loan_amount는 1만 달러 구간 중간값. 파트너: " +
                  "; ".join(f"{p['약칭']}({p['합류']}, LEI {p['LEI']})" for p in parts)))
# 원그래프: 2025년 값은 도표 시트를 수식으로 참조(숫자를 두 번 치지 않는다)
letter = {c[0]: chr(ord("B") + i) for i, c in enumerate(cols)}
last = len(YEARS) + 1
pie_names = top + [ETC]
ws2 = wb.add(PIE, "2025년 파트너사별 HELOC 실행액 비중", pie_names,
             [("2025 실행액", [0] * len(pie_names), FMT["십억"]),
              ("비중", "={2025 실행액}/SUM($B$2:$B$%d)" % (len(pie_names) + 1), FMT["%"], "hide")],
             first="구분", source="도표 1-1 시트 2025년 행 참조", unit="(십억 달러)", size="삽입")
for i, n in enumerate(pie_names):
    ws2.cell(i + 2, 2).value = f"='{NO}'!{letter[n]}{last}"
wb.save(XLSX)

# 파트너 목록 시트(근거) — 마지막 열이 '출처'가 아니므로 load()·대조 대상이 아니다
import openpyxl  # noqa: E402
xb = openpyxl.load_workbook(XLSX)
pw = xb.create_sheet("파트너 목록")
pw.append(["약칭", "파트너", "LEI", "합류", "근거", "URL", "도표 계열"])
for p in parts:
    pw.append([p["약칭"], p["파트너"], p["LEI"], p["합류"], p["근거"], p["URL"],
               p["약칭"] if p["약칭"] in top else ETC])
for c, w in zip("ABCDEFG", (14, 30, 24, 9, 44, 60, 14)):
    pw.column_dimensions[c].width = w
xb.calculation.fullCalcOnLoad = True
xb.save(XLSX)

# ── 2) PPT 도표: 엑셀에서 읽은 값으로만 ──
bk = load(XLSX)
sh, pi = bk[NO], bk[PIE]
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR)", [f"{NO}. {TITLE}"],
            checks=["파트너 목록·합류 시점(엑셀 '파트너 목록' 시트)", "HMDA는 연간 공시 — 2025년이 최신(2026.6 공개)"])
cv = chart_slide(prs, NO, TITLE, "기본", "HMDA(FFIEC), 각사, SMIC 3팀",
                 notes="원자료: " + XLSX + " / 산출: hmda/hmda_heloc.py(HMDA 공개 LAR 연도별 전국 실행분 → LEI별 HELOC 집계)",
                 note_lines=[f"공개된 동사 파트너 {len(parts)}개사의 HMDA 한도형 여신(HELOC) 실행 약정액 합. 파트너 명의 실행분만 집계",
                             "loan_amount는 1만 달러 구간 중간값 기준이라 합계는 근사치"])
names = top + [ETC]
p = add_chart(cv, "area", sh.cats, sh.series(*names), sheet=NO, fmt=sh.fmt(top[0]),
              units=("(십억 달러)", ""), on_tick=True)
# 원그래프 삽입: 영역이 낮은 왼쪽 위 빈 곳
d = 2.05
px, py = p.px0 + 0.15, p.top + 0.05
add_chart(cv, "pie", pi.cats, pi.series("2025 실행액"), sheet=PIE, fmt=pi.fmt("2025 실행액"),
          x=px, y=py, w=d, h=d, legend=False)
cv.label(px + d + 0.1, py + d / 2 - 0.42, "2025년 기준", DG)
cv.label(px + d + 0.1, py + d / 2 - 0.05, "파트너사별 비중", DG)
prs.save(PPTX)
print("saved", XLSX, PPTX)
print("순서:", order)
for k in order:
    print(f"  {k:<16}", " ".join(f"{v/1e9:6.2f}" for v in vol[k]))
print("  시장           ", " ".join(f"{market[y]/1e9:6.1f}" for y in YEARS))
