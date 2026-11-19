# build_figr_charts.py — FIGR 투자포인트① 도표(3-1 ~ 3-8): 원자료 엑셀(유일한 원본) → SMIC 양식 PPT
# 실행: python3 build_figr_charts.py  →  out/FIGR_IP1_도표_원자료.xlsx, out/FIGR_IP1_도표.pptx
# 대조: python3 smic_check.py out/FIGR_IP1_도표.pptx out/FIGR_IP1_도표_원자료.xlsx --render out/render
import csv, json, os
from smic_xlsx import SourceBook, load
from smic_charts_lib import *
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, "raw")
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)
XLSX = os.path.join(OUT, "FIGR_IP1_도표_원자료.xlsx")
PPTX = os.path.join(OUT, "FIGR_IP1_도표.pptx")
SRC = "SMIC 3팀"
B, M, K = 1_000_000_000, 1_000_000, 1_000
PCT1 = "0.0%;(0.0%)"
BAE = '0.0"배"'


def fred(series_id):
    d = {}
    for r in list(csv.reader(open(os.path.join(RAW, f"fred_{series_id}.csv"))))[1:]:
        if r[1] not in ("", "."):
            d[r[0]] = float(r[1])
    return d


def hmda(name):
    return json.load(open(os.path.join(RAW, f"hmda_{name}.json")))["agg"]


wb = SourceBook()

# ───────── 도표 3-1 2순위 주담대 증권화 발행액(와이드) ─────────
wb.add("도표 3-1 삽입", "금융위기 전 2순위 주담대 증권화 발행액(2005)", ["2005"],
       cols=[("발행액", [60_740_000_000], FMT["십억"])],
       source="Inside Mortgage Finance(2006), 'Second-Lien MBS Issuance Up Sharply in 2005 Despite Slowdown in HELOC "
              "Securitization' — 2005년 2순위 MBS 607.4억 달러(전년 대비 +23.6%)",
       unit="(십억 달러)", size="와이드", note="2006~2021년 연간 시계열 미확보 → 도표에서 축 생략")
wb.add("도표 3-1", "미국 2순위 주담대(HELOC·CES) 증권화 발행액 추이와 주요 사건",
       ["2022", "2023", "2024", "2025", "2026E"],
       cols=[("실적", [1 * B, 4_500_000_000, 12 * B, 28_997 * M, 29_829 * M], FMT["십억"]),
             ("26년 전망(S&P)", [None, None, None, None, 40 * B], FMT["십억"], "hide"),
             ("26년 전망 잔여", [None, None, None, None, "=C6-B6"], FMT["십억"]),
             ("22년 대비 배수", [None, None, None, None, "=C6/B2"], '0"배"', "hide")],
       row_sources=["S&P Global Ratings, 'Consumer Pulse: U.S. Second-Lien RMBS'(2026.2) — 2022년 약 10억 달러(SFA Research Corner 2026.9.17 재인용)",
                    "KBRA 집계(National Mortgage News 'How big can home equity securitizations get?') — 2023년 45억 달러",
                    "KBRA 집계(동) — 2024년 '120억 달러를 조금 넘음'",
                    "J.P. Morgan·Bloomberg(SFA Research Corner 2026.9.17 표) — HELOC/CES 2025 FY 28,997백만 달러",
                    "실적: J.P. Morgan 2026 YTD 29,829백만 달러(SFA 2026.9.17) / 전망: S&P 약 400억 달러(SFA 재인용)"],
       unit="(십억 달러)", size="와이드",
       note="26년 실적은 9월 중순 누적, 잔여는 S&P 연간 전망(약 400억 달러)과의 차이. 22·24년은 '약'으로 발표된 값",
       check="22년(약 10억)·24년(120억 '조금 넘음')은 근사치 — Finsight/KBRA 연말 집계로 교체 권장")
wb.add_facts("도표 3-1 사건", "도표 3-1 안의 사건 표기", "도식", facts=[
    ("07~08", "서브프라임 붕괴, 모노라인 보증사 AAA 상실(S&P 2008.6 Ambac·MBIA 강등)", "S&P 등급 조치(2008.6.5)"),
    ("10", "Ambac Financial Group 파산보호 신청(2010.11.8)", "Ambac 8-K(2010.11.8)"),
    ("08~18", "신규 2순위 단독 증권화 사실상 전무", "KBRA(2026.9.8) 'essentially nonexistent for years following the GFC'; Fitch(2019)"),
    ("19", "Towd Point HE Trust 2019-HE1(FirstKey/Cerberus) — 위기 후 첫 HELOC 단독 증권화", "Fitch Ratings, National Mortgage News(2019)"),
    ("23.04", "FIGRE Trust 2023-HE1 — 동사 첫 공모 등급 HELOC ABS, Class A AAA(DBRS)", "Figure 보도자료(PR Newswire 2023.4.25), DBRS Morningstar"),
    ("25.06", "동사 HELOC ABS에 S&P AAA(블록체인 자산 최초)", "Figure 보도자료(Business Wire 2025.6.12)"),
], size="와이드")

# ───────── 도표 3-2 순자산 인출 비용(예시) ─────────
wb.add("도표 3-2 삽입", "미국 모기지 금리 분포(잔액 건수 기준, 26.2Q)", ["26.2Q"],
       cols=[("3% 미만", [0.192], PCT1), ("3~4%", [0.299], PCT1), ("4~5%", [0.165], PCT1),
             ("5~6%", [0.120], PCT1), ("6% 이상", [0.225], PCT1),
             ("4% 미만", "={3% 미만}+{3~4%}", PCT1, "hide")],
       source="FHFA National Mortgage Database, Outstanding Residential Mortgage Statistics(National, All Mortgages, 2026Q2, VALUE1=건수 기준)",
       unit="(%)", first="구분")
wb.add("도표 3-2", "저금리 1순위 차주의 순자산 인출 비용: 현금인출 재융자 vs HELOC(예시)",
       ["재융자\n표면금리", "재융자\n실질금리", "HELOC\n(동사)"],
       cols=[("다시 매기는 기존 잔액", [0, "=12399600000000/51426000", 0], "#,##0", "hide"),
             ("기존 금리", [0.035, 0.035, 0.035], "0.00%", "hide"),
             ("새로 빌리는 돈", [86_000, 86_000, 86_000], "#,##0", "hide"),
             ("적용 금리", [0.0728, 0.0728, 0.092], "0.00%", "hide"),
             ("실질 조달금리", "=(({다시 매기는 기존 잔액}+{새로 빌리는 돈})*{적용 금리}-{다시 매기는 기존 잔액}*{기존 금리})/{새로 빌리는 돈}", PCT1)],
       row_sources=["PMMS 30년 고정 7.28%(Freddie Mac, 2026.11.19) — 표면금리 = 적용 금리",
                    "기존 잔액 = FHFA NMDB 2026Q2 총 잔액 12,399.6십억 달러 ÷ 5,142.6만 건 / 기존 금리 3.5%(최빈 구간 3~4%의 중간, 예시 가정) / 추가 자금 86,000달러",
                    "동사 파트너 HELOC 평균 금리 9.2%·평균 잔액 86천 달러(FIGR 10-Q 3Q26 Loan Characteristics)"],
       unit="(%)", first="구분",
       note="실질 조달금리 = (재융자 후 연 이자 − 기존 연 이자) ÷ 새로 빌린 돈. 첫해 이자 기준 단순 비교, 예시",
       check="기존 금리 3.5%는 예시 가정(FHFA 최빈 구간 3~4%의 중간값)")

# ───────── 도표 3-3 기존 은행 HELOC vs 동사 HELOC(표) ─────────
rows33 = ["금리", "자금 지급", "상환", "실행 기간", "건당 생산비", "차주", "실행 후 출구"]
wb.add("도표 3-3", "기존 은행 HELOC vs 동사 HELOC", rows33, cols=[],
       labels=[("기존 은행 HELOC", ["변동(프라임 연동)", "한도 내 수시 인출", "10년 이자만, 이후 원금",
                                  "업계 중위 약 43일", "1만 1,109달러(MBA)", "(위기 전) CLTV 95%까지", "은행 만기 보유"]),
               ("동사 HELOC", ["고정, 평균 9.2%", "실행 시 전액 지급", "첫 달부터 원리금 균등",
                             "중위 10일", "717달러(25년)", "신용점수 평균 744", "Connect 매각·증권화"])],
       row_sources=["SFA Research Corner(2026.9.17) 'variable rate tied to Prime' / FIGR 10-Q 3Q26 파트너 HELOC 평균 금리 9.2%",
                    "SFA(2026.9.17) 'revolving line… draw as needed' / 동사 HELOC 전액 일시 지급(FIGR 10-K 2025)",
                    "SFA(2026.9.17) '10-year interest-only draw period followed by a 20-year amortization' / FIGR 10-Q 'monthly, amortizing borrower payments'",
                    "FIGR 10-K 2025: 동사 LOS 중위 10일 vs 업계 중위 약 43일(업계 자료)",
                    "FIGR 10-K 2025: 25년 건당 평균 생산비용 약 717달러 vs MBA 업계 평균 11,109달러(2025.9 기준)",
                    "SFA(2026.9.17) 위기 전 피기백 CLTV 95% 근접 / FIGR 10-Q 3Q26 파트너 HELOC 평균 FICO 744",
                    "HMDA 2025(도표 3-7): 대형 은행 HELOC 당해 매각 0% / 동사 Connect·ABS"],
       first="구분", kind="표")

# ───────── 도표 3-5 분기 거래량·Connect·파트너 ─────────
# 공시가 누계(1H·9M·FY)로만 있는 분기는 '공시값 − 차감값' 수식으로 분기 값을 만든다(천 달러 원값 그대로)
# 행: 2=24.1Q 3=24.2Q 4=24.3Q 5=24.4Q 6=25.1Q 7=25.2Q 8=25.3Q 9=25.4Q 10=26.1Q 11=26.2Q 12=26.3Q
q = ["2024 1Q", "2024 2Q", "2024 3Q", "2024 4Q", "2025 1Q", "2025 2Q", "2025 3Q", "2025 4Q", "2026 1Q", "2026 2Q", "2026 3Q"]
wb.add("도표 3-5", "동사 분기 거래량(Connect·기타)과 활성 파트너 수", q,
       cols=[("공시 거래량", [1_191_916 * K, 2_506_171 * K, 1_450_472 * K, 5_128_460 * K, 1_365_136 * K, 3_203_130 * K,
                              2_469_446 * K, 8_377_133 * K, 2_902_378 * K, 12_957_483 * K, 5_796_309 * K], FMT["백만"], "hide"),
             ("차감 거래량", [0, "=B2", 0, 3_956_643 * K, 0, "=B6", 0, 5_672_576 * K, 0, "=B12+B10", 0], FMT["백만"], "hide"),
             ("총 거래량", "={공시 거래량}-{차감 거래량}", FMT["백만"], "hide"),
             ("공시 Connect", [0, 0, 0, 8_144 * K, 477_904 * K, 1_244_566 * K, 1_131_334 * K, 3_842_222 * K, 1_611_840 * K,
                               8_644_079 * K, 4_259_304 * K], FMT["백만"], "hide"),
             ("차감 Connect", [0, 0, 0, 0, 0, "=E6", 0, 2_375_900 * K, 0, "=E12+E10", 0], FMT["백만"], "hide"),
             ("Connect(LHS)", "={공시 Connect}-{차감 Connect}", FMT["백만"]),
             ("기타", "={총 거래량}-{Connect(LHS)}", FMT["백만"]),
             ("LTM 신규 파트너", [None, None, None, 77, 76, 78, None, None, None, None, None], "#,##0", "hide"),
             ("파트너(RHS)", ["=J6-I6", "=J7-I7", None, 134, 144, 168, 246, 307, 387, 489, 612], "#,##0"),
             ("Connect YoY", [None] * 10 + ["=G12/G8"], BAE, "hide"),
             ("Connect(억 달러)", [None] * 10 + ["=G12/100000000"], '0.0"억 달러"', "hide"),
             ("파트너 1년 배수", [None] * 10 + ["=J12/J8"], BAE, "hide"),
             ("출시 후 파트너 배수", [None] * 10 + ["=J12/J3"], BAE, "hide"),
             ("Connect 비중", [None] * 3 + ["={}/{}".format(f"G{r}", f"D{r}") for r in range(5, 13)], "0%", "hide")],
       row_sources=["거래량: FIGR DRS/A(2025.8.4) 1Q24 1,191,916천 달러, Connect '—' / 파트너: 1Q25 144곳 − 직전 12개월 신규 76곳(DRS/A, 이탈 0 가정)",
                    "거래량: S-1(424B4, 2025.9.11) 1H24 2,506,171천 − 1Q24, Connect '—'(24.6 출시) / 파트너: 2Q25 168곳 − 신규 78곳(S-1)",
                    "FIGR 10-Q 3Q25 전년 동기 1,450,472천 달러, Connect '—' / 파트너 미공시",
                    "FY24 5,128,460천(10-K 2025) − 9M24 3,956,643천(10-Q 3Q25), Connect FY24 8,144천 / 파트너 134곳(DRS 2025.5.9, 직전 12개월 신규 77곳)",
                    "FIGR 10-Q 1Q26 전년 동기 1,365,136 / 477,904천 / 파트너 144곳(DRS/A 2025.8.4)",
                    "S-1 1H25 3,203,130 / 1,244,566천 − 1Q25 / 파트너 168곳(S-1)",
                    "FIGR 10-Q 3Q25 2,469,446 / 1,131,334천 / 활성 파트너 246곳",
                    "FY25 8,377,133 / 3,842,222천(10-K 2025) − 9M25 5,672,576 / 2,375,900천(10-Q 3Q25) / 307곳",
                    "FIGR 10-Q 1Q26 2,902,378 / 1,611,840천 / 387곳(신규 80곳)",
                    "9M26 12,957,483 / 8,644,079천(10-Q 3Q26) − 3Q26 − 1Q26 / 489곳(신규 102곳, 2Q26 실적 발표)",
                    "FIGR 10-Q 3Q26 5,796,309 / 4,259,304천(Kiavi 677.3백만 달러 포함) / 독립 파트너 612곳(3Q26 실적 발표)"],
       unit="(백만 달러) / (곳)",
       note="Connect는 2024.6 출시. 24.1Q·2Q 파트너 수는 S-1·DRS의 '직전 12개월 신규 파트너'를 빼서 역산(이탈 0 가정), 24.3Q 파트너 수 미공시. 26.3Q 거래량에 Kiavi 6.77억 달러 포함",
       check="파트너 수 정의: ~26.2Q '활성 파트너', 26.3Q '독립(standalone) 파트너'(Kiavi 포함 894곳). 21년 말 파트너 3곳(S-1)·23년 말 57곳(역산)은 도표 밖")

# ───────── 도표 3-4 양면 플랫폼(도식) ─────────
wb.add_facts("도표 3-4", "동사 양면 플랫폼: 공장(LOS)과 장터(Connect)", "도식", facts=[
    ("612곳", "독립 파트너 수(26.3Q, Kiavi 제외)", "FIGR 3Q26 실적 발표(2026.11.12)"),
    ("20곳+", "활성 예금기관 파트너 20곳 이상", "FIGR 3Q26 실적 발표"),
    ("중위 10일", "LOS 실행 중위 기간(업계 약 43일)", "FIGR 10-K 2025"),
    ("717달러", "25년 건당 평균 생산비용(MBA 업계 평균 11,109달러)", "FIGR 10-K 2025"),
    ("DART", "Digital Asset Registry Technology — 자산 보관·담보권 완성", "FIGR 3Q26 실적 발표 About Figure"),
    ("28곳", "YLDS로 마켓플레이스 결제하는 참여자", "FIGR 3Q26 실적 발표"),
])

# ───────── 도표 3-6 1순위가 걸은 길 ─────────
T, G, P, A = fred("ASHMA"), fred("BOGZ1FL403065105Q"), fred("BOGZ1FL413065105Q"), fred("BOGZ1FL673065105Q")
yrs = list(range(1966, 2026))
key = [f"{y}-10-01" for y in yrs]
wb.add("도표 3-6", "1순위가 걸은 길: 1~4가구 주택담보대출 중 2차 시장(GSE·MBS 풀·민간 ABS) 보유 비중", [str(y) for y in yrs],
       cols=[("총액", [int(T[k] * M) for k in key], FMT["십억"], "hide"),
             ("GSE 보유", [int(G[k] * M) for k in key], FMT["십억"], "hide"),
             ("에이전시·GSE MBS 풀", [int(P[k] * M) for k in key], FMT["십억"], "hide"),
             ("민간 ABS", [int(A[k] * M) for k in key], FMT["십억"], "hide"),
             ("2차 시장 보유 비중", "=({GSE 보유}+{에이전시·GSE MBS 풀}+{민간 ABS})/{총액}", PCT1)],
       labels=[("축 레이블", ["'" + str(y)[2:] if y % 10 == 0 else "" for y in yrs])],
       source="Federal Reserve Z.1 Financial Accounts(FRED: ASHMA 총액, FL403065105Q GSE, FL413065105Q 에이전시·GSE MBS 풀, FL673065105Q 민간 ABS), 각 연도 4분기 말 잔액(백만 달러)",
       unit="(%)", note="2010년 GSE 신탁 연결(FAS 166/167)로 풀 → GSE 보유로 이동했으나 합계는 연속. GSE 보유에는 비증권화 포트폴리오 포함",
       check="본문 '1970년 5% → 30년 만에 60%' ↔ Z.1: 1970년 6.2%, 2000년 59.1%")
wb.add("도표 3-6 삽입", "HELOC 증권화율(26.3Q)", ["26.3Q"],
       cols=[("증권화 HELOC 잔액", [21_500_000_000], FMT["십억"]),
             ("HELOC 잔액", [481 * B], FMT["십억"]),
             ("HELOC 증권화율", "={증권화 HELOC 잔액}/{HELOC 잔액}", PCT1)],
       source="증권화 잔액: Deutsche Bank·Intex(SFA Research Corner 2026.9.17 '약 215억 달러') / HELOC 잔액: NY Fed Household Debt and Credit 2026Q3(0.481조 달러)",
       unit="(%)", first="구분", check="증권화 잔액은 26.9 중순, HELOC 잔액은 26.9말 — 시점 근사")

# ───────── 도표 3-7 같은 은행, 다른 출구(HMDA) ─────────
banks = [("PNC", "PNC"), ("Fifth Third", "FifthThird"), ("Citizens", "Citizens"), ("Flagstar", "Flagstar"), ("동사", "Figure")]
c1, c1s, hl, hls = [], [], [], []
for _, key_ in banks:
    a = hmda(key_)
    first = a.get("CLOSED_L1")
    c1.append(int(first[1]) if first else None)
    c1s.append(int(first[3]) if first else None)
    h = [a[k] for k in ("HELOC_L1", "HELOC_L2") if k in a]
    hl.append(int(sum(x[1] for x in h)))
    hls.append(int(sum(x[3] for x in h)))
wb.add("도표 3-7", "같은 은행, 다른 출구: 2025년 실행 당해 매각 비중(금액 기준)", [b[0] for b in banks],
       cols=[("1순위 실행액", c1, FMT["백만"], "hide"), ("1순위 매각액", c1s, FMT["백만"], "hide"),
             ("HELOC 실행액", hl, FMT["백만"], "hide"), ("HELOC 매각액", hls, FMT["백만"], "hide"),
             ("1순위 주담대", ["=C2/B2", "=C3/B3", "=C4/B4", "=C5/B5", None], "0%"),
             ("HELOC", "={HELOC 매각액}/{HELOC 실행액}", "0%")],
       row_sources=[f"CFPB HMDA Data Browser 2025 LAR(LEI {l}), action_taken=1, 매각 = purchaser_type≠0, 1순위 = 폐쇄형·lien 1, HELOC = open-end(1·2순위)"
                    for l in ("AD6GFRVSDT01YPT1CS68", "QFROUN1UWUYU0DVIWD51", "DRMSV1Q0EKMEXLAU1P80", "SS1TRMSN6BRNMOREEV51", "254900UL88QFG0E40516(Figure Lending LLC)")],
       unit="(%)", first="금융기관",
       note="HMDA 금액은 HELOC의 경우 약정 한도. '당해 매각'은 실행 연도 안에 매각된 건(이후 매각은 미포함). 동사는 Figure Lending LLC 명의분만")

# ───────── 도표 3-8 확장 경로(와이드, 도식 + 작은 차트) ─────────
wb.add("도표 3-8", "동사 거래량 중 1순위 비중", ["24.3Q", "25.3Q", "25.4Q", "26.1Q"],
       cols=[("1순위 비중", ["=B3-0.065", 0.17, 0.19, 0.20], "0%;(0%)")],
       row_sources=["FIGR 3Q25 실적 발표: 25.3Q 17%, '전년 동기 대비 650bp 상승' → 10.5%",
                    "FIGR 3Q25 실적 발표", "FIGR 4Q25 실적 발표", "FIGR 1Q26 실적 발표('Year of the First Lien')"],
       unit="(%)", size="와이드", check="본문 '1년 새 14%→20%' ↔ 공시: 24.3Q 10.5% → 25.3Q 17% → 26.1Q 20%(26.2Q·3Q 미공시)")
wb.add("도표 3-8 삽입", "1순위 모기지 잔액 vs HELOC 잔액(26.3Q)", ["26.3Q"],
       cols=[("모기지 잔액", [13_374 * B], FMT["조"]), ("HELOC 잔액", [481 * B], FMT["십억"]),
             ("배수", "={모기지 잔액}/{HELOC 잔액}", BAE),
             ("HELOC 잔액(억 달러)", "={HELOC 잔액}/100000000", '#,##0"억 달러"', "hide")],
       source="NY Fed Household Debt and Credit Report 2026Q3(2026.11.10): Mortgage 13.374조 달러, HE Revolving 0.481조 달러",
       unit="(조 달러)", first="구분", size="와이드")
wb.add_facts("도표 3-8 사실", "도표 3-8 안의 사실 표기", "도식", facts=[
    ("83억 달러", "25년 동사 홈에쿼티 거래 약 83억 달러(+62%)", "FIGR 10-K 2025"),
    ("비은행 1위", "25년 비예금기관 중 HELOC 실행 1위", "FIGR 10-K 2025"),
    ("20곳+", "활성 예금기관 파트너 20곳 이상(26.3Q)", "FIGR 3Q26 실적 발표"),
    ("Flagstar", "Flagstar Bank 계약, 26.2Q 가동", "FIGR 1Q26 실적 발표"),
    ("26.9.1", "Kiavi 인수 완료일", "FIGR 10-Q 3Q26"),
    ("80억 달러", "Kiavi 연환산 거래 규모(run-rate)", "FIGR 3Q26 실적 발표"),
    ("6.8억 달러", "Kiavi의 26.3Q 거래량 기여 677.3백만 달러", "FIGR 10-Q 3Q26"),
    ("612곳", "독립 파트너 수(26.3Q)", "FIGR 3Q26 실적 발표"),
], size="와이드")

wb.save(XLSX)
bk = load(XLSX)
F = load.facts

# ════════════════════════════ PPT ════════════════════════════
prs = new_deck()
intro_slide(prs, "Figure Technology Solutions(FIGR) 투자포인트①", [
    "도표 3-1. 미국 2순위 주담대 증권화 발행액 추이와 주요 사건 (와이드 18.50×5.10, 문단 1~4 뒤)",
    "도표 3-2. 저금리 1순위 차주의 순자산 인출 비용 (기본, 문단 5)",
    "도표 3-3. 기존 은행 HELOC vs 동사 HELOC (기본, 문단 3·8)",
    "도표 3-4. 동사 양면 플랫폼: 공장(LOS)과 장터(Connect) (기본, 문단 9·11·14)",
    "도표 3-5. 동사 분기 거래량과 활성 파트너 수 (기본, 문단 13·15·16)",
    "도표 3-6. 1순위가 걸은 길: 2차 시장 보유 비중 (기본, 문단 17)",
    "도표 3-7. 같은 은행, 다른 출구: 실행 당해 매각 비중 (기본, 문단 17·18·20)",
    "도표 3-8. 동사 확장 경로: 비은행 HELOC → 은행 → 1순위 → 비은행 여신 (와이드, 문단 21~25)"],
    checks=["22·24년 2순위 발행액은 '약' 수치 — 연말 집계로 교체 권장", "도표 3-2는 예시 계산(기존 금리 3.5% 가정)",
            "본문 수치 차이: 39일→업계 중위 43일, 1970년 5%→6.2%, 1순위 14%→24.3Q 10.5%"])

def recolor(p, ser_name, idx, color):
    """labels(points=[…])로 만든 점 레이블 중 idx 하나만 색을 바꾼다(같은 계열을 labels로 두 번 부르면 앞의 것이 지워지므로)"""
    from smic_charts_lib import _series_xml, _ser_name, _txpr
    ser = next(s for s in _series_xml(p.ch) if _ser_name(s) == ser_name)
    for d in ser.find(qn("c:dLbls")).findall(qn("c:dLbl")):
        if d.find(qn("c:idx")).get("val") == str(idx):
            old = d.find(qn("c:txPr")); old.addprevious(_txpr(CT, color)); d.remove(old)


def show_name(p, ser_name):
    """데이터레이블에 계열 이름도 함께(예: '3% 미만 19%')"""
    from smic_charts_lib import _series_xml, _ser_name
    ser = next(s for s in _series_xml(p.ch) if _ser_name(s) == ser_name)
    dls = ser.find(qn("c:dLbls"))
    dls.find(qn("c:showSerName")).set("val", "1")
    sep = etree.SubElement(dls, qn("c:separator")); sep.text = " "


def lines(cv, xc, y0, items, step=0.37):
    """세로로 한 줄씩 가운데 정렬(렌더러마다 줄 간격이 달라도 위치 고정). item = (글자, 색) 또는 ('VAL', 시트, 열, i, pre, post, fmt, 색)"""
    for k, it in enumerate(items):
        y = y0 + k * step
        if it[0] == "VAL":
            _, sh, col, i, pre, post, fmt, c = it
            cv.val(xc, y, sh, col, i, pre=pre, post=post, fmt=fmt, color=c, align="c")
        elif it[0]:
            cv.label(xc, y, it[0], it[1], align="c")


# ── 도표 3-1 (와이드) ──
s0, s1 = bk["도표 3-1 삽입"], bk["도표 3-1"]
cv = chart_slide(prs, "도표 3-1", "미국 2순위 주담대(HELOC·CES) 증권화 발행액 추이와 주요 사건", "와이드",
                 "Inside Mortgage Finance, KBRA, J.P. Morgan, S&P, SFA, Fitch, 각 사, " + SRC,
                 notes="2005: IMF 607.4억 달러 / 2022 약 10억(S&P) / 2023 45억·2024 약 120억(KBRA) / 2025 289.97억(JPM) / 2026 9월 누적 298.29억(JPM), 연간 전망 약 400억(S&P)",
                 note_lines=["2006~2021년은 축 생략(연간 발행 시계열 미확보). 26E 빨강 = 9월 중순 누적 실적, 분홍 = S&P 연간 전망(약 400억 달러)까지 잔여",
                             "22년(약 10억 달러)·24년(120억 달러 '조금 넘음')은 근사치"])
YS = (0, 70 * B, 10 * B)
INNER = None
pl = add_chart(cv, "bar", s0.cats, s0.series("발행액"), sheet=s0.name, fmt=s0.fmt("발행액"), x=0, y=0, w=2.6, h=cv.h,
               yscale=YS, units=("", ""), colors=[MG], legend=False, gap=120, inner=(0.78, 0.55, 0.12, 0.38))
cv.label(0.0, 0.05, "(십억 달러)", DG)
pl.labels("발행액", fmt="#,##0.0,,,", color=MG)
pr = add_chart(cv, "stack", s1.cats, s1.series("실적", "26년 전망 잔여"), sheet=s1.name, fmt=s1.fmt("실적"),
               x=9.7, y=0, w=cv.w - 9.7, h=cv.h, yscale=YS, colors=[LG, PINK], emph={4: RED}, gap=45, axes=False,
               inner=(0.12, 0.55, 0.05, 0.38))
pr.labels("실적", points=[4], fmt="#,##0.0,,,", color=WHT, pos="ctr")
for i_ in range(4):
    cv.val(pr.x(i_), pr.y(s1.cols["실적"][i_]) - 0.38, s1, "실적", i_, fmt="#,##0.0,,,", color=MG, align="c")
cv.label(pr.x(4), pr.y(s1.cols["실적"][4] / 2) + 0.2, "9월 누적", WHT, align="c")
# 06~21 축 생략: 세로 타임라인
bx0, bx1 = 2.75, 9.55
cv.band(bx0, bx1, pl.top, pl.bottom, BG, 0.6)
cv.label((bx0 + bx1) / 2, pl.top + 0.1, "06~21 축 생략: 10년 넘게 닫힌 2순위 증권화의 문", DG, align="c")
nx, ny0, st_ = 3.75, pl.top + 0.85, 0.66
ev = [("07~08", "서브프라임 붕괴·모노라인 AAA 상실", DG, True), ("10", "모노라인 Ambac 파산보호", DG, True),
      ("08~18", "신규 2순위 증권화 사실상 0", RED, True), ("", "→ HELOC은 은행 만기 보유 자산으로", RED, False),
      ("19", "위기 후 첫 HELOC 증권화(Towd Point)", DG, True)]
cv.line(nx, ny0, nx, ny0 + st_ * (len(ev) - 1), MG, 1.5)
for k, (d, txt, c, node) in enumerate(ev):
    yy = ny0 + st_ * k
    if node:
        cv.box(nx - 0.1, yy - 0.1, 0.2, 0.2, fill=WHT, line=c, lw=1.5, shape="oval")
        cv.label(nx - 0.2, yy - 0.18, d, c, align="r")
    cv.label(nx + 0.25, yy - 0.18, txt, c)
# 05년 정점선(오른쪽 차트에만) + 40배 화살표 + 사건
yp = pr.y(s0.cols["발행액"][0])
cv.line(pr.px0, yp, pr.px0 + pr.pw, yp, MG, 0.75, dash=True)
cv.val(pr.px0 + 0.05, yp - 0.38, s0, "발행액", 0, pre="금융위기 전 정점(05년) ", fmt="#,##0.0,,,", color=MG)
cv.arrow(pr.x(0), pr.y(20 * B), pr.x(4) - 0.72, pr.y(44 * B), RED, 1.0)
cv.val(pr.x(1) - 0.3, pr.y(36 * B), s1, "22년 대비 배수", 4, pre="4년 만에 ", fmt='0"배"', color=RED, align="c")
cv.val(pr.x(4), pr.y(40 * B) - 0.38, s1, "26년 전망(S&P)", 4, pre="전망 ", fmt="#,##0.0,,,", color=RED, align="c")
cv.label(pr.x(1), pr.y(s1.cols["실적"][1]) - 1.1, "23.04 동사 첫", DG, align="c")
cv.label(pr.x(1), pr.y(s1.cols["실적"][1]) - 0.77, "공모 AAA(DBRS)", DG, align="c")
cv.label(pr.x(3), pr.y(14 * B), "25.06", DG, align="c")
cv.label(pr.x(3), pr.y(14 * B) + 0.33, "S&P AAA", DG, align="c")

# ── 도표 3-2 ──
si, sm = bk["도표 3-2 삽입"], bk["도표 3-2"]
cv = chart_slide(prs, "도표 3-2", "저금리 1순위 차주의 순자산 인출 비용: 현금인출 재융자 vs HELOC (예시)", "기본",
                 "FHFA NMDB, Freddie Mac PMMS, 동사 10-Q, " + SRC,
                 notes="실질 조달금리 = (재융자 후 연 이자 − 기존 연 이자) ÷ 새로 빌린 돈. 기존 잔액 241,117달러(FHFA 26.2Q 평균), 기존 금리 3.5%(가정), 추가 86,000달러, 재융자 7.28%(PMMS 11/19), HELOC 9.2%(동사 3Q26)",
                 note_lines=["왼쪽: 미국 모기지 잔액 건수 기준 금리 분포(FHFA, 26.2Q). 오른쪽 예시: 기존 잔액 24.1만 달러(FHFA 평균)·금리 3.5%(가정),",
                             "추가 자금 8.6만 달러(동사 HELOC 평균), 재융자 7.28%(PMMS 26.11.19). 실질 금리 = (재융자 후 연 이자 − 기존 연 이자) ÷ 새로 빌린 돈"])
buckets = ["3% 미만", "3~4%", "4~5%", "5~6%", "6% 이상"]
bcol = [RED, PINK2, LG, MG, DG]
pi = add_chart(cv, "stack100", si.cats, si.series(*buckets), sheet=si.name, fmt=PCT1, x=0, y=0, w=2.72, h=cv.h,
               colors=bcol, legend=False, axes=False, yscale=(0, 1, 0.25), gap=8, inner=(0.05, 0.55, 0.05, 0.38))
barw = pi.pw / 1.08
bl, br = pi.px0 + (pi.pw - barw) / 2, pi.px0 + (pi.pw + barw) / 2
cum = 0.0
for b_, c_ in zip(buckets, bcol):
    v = si.cols[b_][0]
    yy = pi.y(cum + v / 2) - 0.18
    tc = DG if c_ == LG else WHT
    cv.label(bl + 0.06, yy, b_, tc)
    cv.val(br - 0.05, yy, si, b_, 0, fmt="0.0%", color=tc, align="r")
    cum += v
cv.label(0.0, 0.08, "미국 모기지 금리 분포", DG)
yb0, yb1 = pi.y(0), pi.y(si.cols["4% 미만"][0])
bxr = 2.82
cv.line(bxr, yb0, bxr, yb1, RED, 1.0)
cv.line(bxr - 0.1, yb0, bxr, yb0, RED, 1.0)
cv.line(bxr - 0.1, yb1, bxr, yb1, RED, 1.0)
cv.label(bxr + 0.08, (yb0 + yb1) / 2 - 0.36, "4% 미만", RED)
cv.val(bxr + 0.08, (yb0 + yb1) / 2 + 0.0, si, "4% 미만", 0, fmt="0.0%", color=RED)
pm = add_chart(cv, "bar", sm.cats, sm.series("실질 조달금리"), sheet=sm.name, fmt=PCT1, x=3.95, y=0, w=cv.w - 3.95, h=cv.h,
               units=("(%)", ""), colors=[LG], emph={1: DG, 2: RED}, gap=55, yscale=(0, 0.3, 0.05), inner=(0.62, 0.55, 0.12, 0.7))
pm.labels("실질 조달금리", points=[0, 1, 2], fmt="0.0%", color=DG)
recolor(pm, "실질 조달금리", 2, RED)
cv.label(pm.x(0) - 0.5, pm.top + 0.02, "재융자는 기존 3%대 잔액까지", MG)
cv.label(pm.x(0) - 0.5, pm.top + 0.35, "7%대로 갈아타는 셈", MG)

# ── 도표 3-3 ──
st = bk["도표 3-3"]
cv = chart_slide(prs, "도표 3-3", "기존 은행 HELOC vs 동사 HELOC", "기본", "SFA, MBA, 동사 10-K·10-Q, HMDA, " + SRC,
                 note_lines=["건당 생산비: 동사 25년 평균, 업계는 MBA 25.9 기준 / 실행 기간: 동사 LOS 중위, 업계 중위(동사 10-K 인용 업계 자료)",
                             "차주: 위기 전 피기백 2순위 CLTV(SFA) vs 동사 파트너 HELOC 평균 신용점수(26.3Q)"])
cv.table_from_sheet(0, 0, [1.75, 3.7, 3.8], st, ["기존 은행 HELOC", "동사 HELOC"], first="구분", rowh=0.52,
                    emph={(i, 2) for i in range(1, 8)})
cv.box(0, 4.5, cv.w, 0.55, fill=YEL, shape="rect", paras=[[("'마이너스 통장'을 '팔기 좋은 할부 대출'로 재설계", RED)]])

# ── 도표 3-4 (도식) ──
s5 = bk["도표 3-5"]
cv = chart_slide(prs, "도표 3-4", "동사 양면 플랫폼: 공장(LOS)과 장터(Connect)", "기본", "동사 10-K·10-Q·실적 발표, " + SRC,
                 note_lines=["파트너 612곳은 Kiavi 제외 독립 기준(Kiavi 포함 894곳), Connect 거래액은 26.3Q, YoY는 25.3Q 대비"])
bw = [1.62, 1.62, 1.3, 2.0, 1.62]
gx = [0.0]
for w_ in bw[:-1]:
    gx.append(gx[-1] + w_ + 0.27)
yH, yA, hA = 0.0, 0.42, 2.05
cv.label(gx[0] + bw[0] / 2, yH, "공급", MG, align="c")
cv.label((gx[1] + gx[3] + bw[3]) / 2, yH, "동사 인프라", RED, align="c")
cv.line(gx[1], yH + 0.36, gx[3] + bw[3], yH + 0.36, RED, 0.75)
cv.label(gx[4] + bw[4] / 2, yH, "수요", MG, align="c")
fills = [BG, RED, LG, RED, BG]
for k in range(5):
    cv.box(gx[k], yA, bw[k], hA, fill=fills[k])
y0 = yA + 0.12
lines(cv, gx[0] + bw[0] / 2, y0, [("파트너", BLK), ("612곳", RED), ("", None), ("예금기관", DG), ("20곳+", DG)])
lines(cv, gx[1] + bw[1] / 2, y0, [("LOS(공장)", WHT), ("", None), ("중위 10일", WHT), ("건당", WHT), ("717달러", WHT)])
lines(cv, gx[2] + bw[2] / 2, y0, [("DART", BLK), ("", None), ("전자등기", DG), ("온체인", DG), ("상환 공시", DG)])
lines(cv, gx[3] + bw[3] / 2, y0, [("Connect(장터)", WHT), ("", None), ("26.3Q 거래", WHT),
                                   ("VAL", s5, "Connect(억 달러)", 10, "", "", None, WHT),
                                   ("VAL", s5, "Connect YoY", 10, "YoY ", "", None, WHT)])
lines(cv, gx[4] + bw[4] / 2, y0, [("보험사", BLK), ("연기금", BLK), ("리츠·은행", BLK), ("YLDS 결제", DG), ("28곳", DG)])
ya = yA + hA / 2
for i in range(3):
    cv.arrow(gx[i] + bw[i] + 0.03, ya, gx[i + 1] - 0.03, ya, DG, 1.5)
cv.arrow(gx[4] - 0.03, ya, gx[3] + bw[3] + 0.03, ya, DG, 1.5)
yF, hF = 3.1, 0.72
cv.label(0.0, yF - 0.42, "양면 플라이휠: 공급과 수요가 서로를 키운다", DG)
fw = ["파트너 증가", "같은 규격\n물량 증가", "매수자 경쟁\n심화", "매각가 개선"]
fwid, fgap = 1.9, (cv.w - 4 * 1.9) / 3
fx = [i * (fwid + fgap) for i in range(4)]
for i, t_ in enumerate(fw):
    cv.box(fx[i], yF, fwid, hF, fill=WHT, line=RED if i == 0 else DG, lw=1.0, paras=t_, color=RED if i == 0 else BLK)
    if i < 3:
        cv.arrow(fx[i] + fwid + 0.03, yF + hF / 2, fx[i + 1] - 0.03, yF + hF / 2, DG, 1.5)
yr = yF + hF + 0.22
cv.line(fx[3] + fwid / 2, yF + hF, fx[3] + fwid / 2, yr, DG, 1.5)
cv.line(fx[3] + fwid / 2, yr, fx[0] + fwid / 2, yr, DG, 1.5)
cv.arrow(fx[0] + fwid / 2, yr, fx[0] + fwid / 2, yF + hF + 0.02, DG, 1.5)
cv.box(0, 4.55, cv.w, 0.5, fill=RED, shape="rect", paras=[[("공장(LOS)과 장터(Connect)를 모두 가진 곳은 동사뿐", WHT)]])

# ── 도표 3-5 ──
cv = chart_slide(prs, "도표 3-5", "동사 분기 거래량(Connect·기타)과 활성 파트너 수", "기본", "동사 10-Q·실적 발표, " + SRC,
                 notes="Loan Marketplace Volume(구 Consumer Loan Marketplace Volume), Connect = 제3자 판매자가 Connect에서 거래한 물량",
                 note_lines=["24.1Q·2Q 파트너 수는 S-1·DRS의 '직전 12개월 신규 파트너'를 빼서 역산(이탈 0 가정), 24.3Q는 미공시(선으로 이음)",
                             "26.3Q 총 거래량에 Kiavi(26.9.1 인수) 6.77억 달러 포함, 26.3Q 파트너는 Kiavi 제외 독립 기준"])
p5 = add_chart(cv, "stack", s5.cats, s5.series("Connect(LHS)", "기타", "파트너(RHS)"), sheet=s5.name,
               fmt=s5.fmt("Connect(LHS)"), units=("(백만 달러)", ""), colors=[RED, LG, DG], multilevel=True, gap=45, head=0.12)
p5.rhs(["파트너(RHS)"], fmt="#,##0", unit="(곳)", colors=[DG])
_cs = p5.ch._chartSpace.find(qn("c:chart"))
_db = _cs.find(qn("c:dispBlanksAs"))
if _db is None:
    _db = etree.SubElement(_cs, qn("c:dispBlanksAs"))
    _cs.remove(_db); _cs.find(qn("c:plotVisOnly")).addnext(_db)
_db.set("val", "span")                                     # 24.3Q 파트너 미공시 → 선은 이어서
p5.labels("파트너(RHS)", points=[1, 10], fmt="#,##0", color=DG)
xl = (p5.x(1) + p5.x(2)) / 2
cv.line(xl, p5.top + 1.62, xl, p5.bottom, MG, 0.75, dash=True)
cv.label(xl, p5.top + 1.27, "24.06 Connect 출시", MG, align="c")
cv.val(p5.px0 + 0.1, p5.top + 0.45, s5, "Connect YoY", 10, pre="Connect 거래 1년 새 ", color=RED)
cv.val(p5.px0 + 0.1, p5.top + 0.8, s5, "출시 후 파트너 배수", 10, pre="파트너 Connect 출시 후 ", color=DG)

# ── 도표 3-6 ──
s6, s6i = bk["도표 3-6"], bk["도표 3-6 삽입"]
cv = chart_slide(prs, "도표 3-6", "1순위가 걸은 길: 1~4가구 주택담보대출 중 2차 시장 보유 비중", "기본",
                 "Federal Reserve Z.1, NY Fed, SFA, " + SRC,
                 notes="2차 시장 보유 비중 = (GSE 보유 + 에이전시·GSE MBS 풀 + 민간 ABS) ÷ 1~4가구 주택담보대출 총액, 연말 기준",
                 note_lines=["2차 시장 = GSE 보유 + 에이전시·GSE MBS 풀 + 민간 ABS, 연말 잔액 기준(GSE 보유에는 비증권화 포트폴리오 포함)",
                             "HELOC 증권화율 = 증권화 HELOC 잔액(약 215억 달러, SFA) ÷ HELOC 잔액(4,810억 달러, NY Fed 26.3Q)"])
p6 = add_chart(cv, "line", s6.axis, s6.series("2차 시장 보유 비중"), sheet=s6.name, fmt=PCT1, units=("(%)", ""),
               colors=[DG], yscale=(0, 0.8, 0.2), markers=False)
col6 = s6.cols["2차 시장 보유 비중"]
i70, i81, i86, i00 = 4, 15, 20, 34
p6.labels("2차 시장 보유 비중", points=[i70, i00], fmt="0.0%", color=DG)
for i_, t1, t2, ly, al in [(i70, "'70 지니메이", "첫 MBS 발행", 0.33, "l"), (i81, "'81 패니메이", "MBS 개시", 0.52, "r"),
                           (i86, "'86 REMIC 도입", "(증권화 표준 세제)", 0.72, "r")]:
    xx, yv, yt = p6.x(i_), p6.y(col6[i_]), p6.y(ly)
    cv.line(xx, yt + 0.68, xx, yv - (0.45 if i_ == i70 else 0.1), MG, 0.75, dash=True)
    lx = p6.px0 + 0.05 if al == "l" else xx + 0.1
    cv.label(lx, yt, t1, MG, align=al)
    cv.label(lx, yt + 0.33, t2, MG, align=al)
yh = p6.y(s6i.cols["HELOC 증권화율"][0])
cv.line(p6.px0, yh, p6.x(59), yh, RED, 1.0, dash=True)
cv.diamond(p6.x(59), yh, 0.24, RED)
cv.val(p6.x(59) - 0.1, yh - 0.78, s6i, "HELOC 증권화율", 0, pre="HELOC 증권화율(26.3Q) ", color=RED, align="r")
cv.label(p6.x(59) - 0.1, yh - 0.45, "= 1순위의 1970년 수준", RED, align="r")

# ── 도표 3-7 ──
s7 = bk["도표 3-7"]
cv = chart_slide(prs, "도표 3-7", "같은 은행, 다른 출구: 2025년 실행 당해 매각 비중(금액 기준)", "기본", "CFPB HMDA(2025), " + SRC,
                 notes="HMDA 2025 LAR, action_taken=1(실행), 매각 = purchaser_type≠0. 1순위 = 폐쇄형 1순위 주담대, HELOC = 개방형(1·2순위). 동사 = Figure Lending LLC",
                 note_lines=["HMDA 2025 실행분 중 같은 해 매각된 금액 비중(HELOC은 약정 한도 기준). 동사 = Figure Lending LLC 명의분",
                             "1순위 = 폐쇄형 1순위 주담대, HELOC = 개방형(1·2순위 포함)"])
p7 = add_chart(cv, "bar", s7.cats, s7.series("1순위 주담대", "HELOC"), sheet=s7.name, fmt="0%", units=("(%)", ""),
               colors=[LG, RED], gap=60, yscale=(0, 1, 0.2))
p7.labels("1순위 주담대", fmt="0%", color=MG)
p7.labels("HELOC", fmt="0%", color=RED)
cv.label((p7.x(2) + p7.x(3)) / 2, p7.y(0.66), "은행: 1순위는 팔고", RED, align="c")
cv.label((p7.x(2) + p7.x(3)) / 2, p7.y(0.66) + 0.33, "HELOC은 전부 장부에", RED, align="c")

# ── 도표 3-8 (와이드) ──
s8, s8i = bk["도표 3-8"], bk["도표 3-8 삽입"]
cv = chart_slide(prs, "도표 3-8", "동사 확장 경로: 비은행 HELOC → 은행 → 1순위 → 비은행 여신 전체", "와이드",
                 "동사 10-K·10-Q·실적 발표, NY Fed, HMDA, " + SRC,
                 note_lines=["1순위 비중: 24.3Q는 25.3Q 17%에서 '전년 대비 650bp 상승'을 역산, 26.2Q~3Q는 미공시. 잔액 배수 = 모기지 잔액 ÷ HELOC 잔액(NY Fed 26.3Q)",
                             "은행 HELOC 당해 매각 비중은 도표 3-7(HMDA 2025, PNC·Fifth Third·Citizens·Flagstar)"])
pw_, gp = 4.35, 0.367
hx = [i * (pw_ + gp) for i in range(4)]
heads = ["① 비은행 HELOC(본진)", "② 은행 HELOC", "③ 1순위 담보대출", "④ 비은행 여신 전체"]
for i, t_ in enumerate(heads):
    cv.box(hx[i], 0.0, pw_ + (0.25 if i < 3 else 0), 0.55, fill=DG if i == 0 else MG,
           shape="pentagon" if i < 3 else "rect", paras=[[(t_, WHT)]])
    cv.box(hx[i], 0.68, pw_, 3.62, fill=BG, shape="rect")
cx = [h_ + pw_ / 2 for h_ in hx]
yl = 0.95
lines(cv, cx[0], yl, [("25년 홈에쿼티 거래", DG), ("83억 달러", RED), ("비은행 1위 HELOC 실행사", DG), ("파트너 612곳", DG), ("", None),
                      ("HELOC 시장 잔액(26.3Q)", MG), ("VAL", s8i, "HELOC 잔액(억 달러)", 0, "", "", None, MG)], step=0.43)
lines(cv, cx[1], yl, [("예금기관 파트너", DG), ("20곳+", RED), ("Flagstar 26.2Q 가동", DG), ("", None),
                      ("VAL", s7 if False else bk["도표 3-7"], "HELOC", 0, "대형 은행 HELOC 당해 매각 ", "", None, DG),
                      ("→ 장부를 비울 출구 필요", DG)], step=0.43)
p8 = add_chart(cv, "bar", s8.cats, s8.series("1순위 비중"), sheet=s8.name, fmt="0%", x=hx[2] + 0.15, y=1.05, w=pw_ - 0.3, h=1.95,
               colors=[MG], emph={3: RED}, gap=60, units=("", ""), axes=False, yscale=(0, 0.25, 0.05),
               inner=(0.05, 0.3, 0.05, 0.36))
p8.labels("1순위 비중", points=[0, 1, 2, 3], fmt="0%", color=MG)
recolor(p8, "1순위 비중", 3, RED)
cv.label(cx[2], 0.8, "동사 거래량 중 1순위 비중", DG, align="c")
lines(cv, cx[2], 3.12, [("모기지 잔액은 HELOC의", DG), ("VAL", s8i, "배수", 0, "", "", BAE, RED),
                        ("VAL", s8i, "모기지 잔액", 0, "(", "조 달러, 26.3Q)", "#,##0.0,,,,", MG)], step=0.37)
lines(cv, cx[3], yl, [("Kiavi 인수(26.9.1)", DG), ("RTL·DSCR 연 80억 달러", RED), ("26.3Q 기여 6.8억 달러", DG), ("", None),
                      ("같은 레일 위에", DG), ("비에이전시 여신 연결", DG)], step=0.43)
cv.box(0, 4.45, cv.w, 0.6, fill=RED, shape="rect",
       paras=[[("생산(LOS)·검증(DART)·유통(Connect) 레일 하나에 자산군을 계속 얹는다 — 은행 양방향 유입과 20배 넘는 1순위 시장이 다음 무대", WHT)]])

prs.save(PPTX)
print("saved", XLSX, PPTX)
