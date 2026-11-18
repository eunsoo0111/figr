"""투자포인트 1 ② 차주 입장 계산.

python3 -I research/scripts/ip1_borrower_calcs.py research
입력: research/data/fred_mortgage30us_weekly.csv, research/data/nyfed_hhdc_heloc_2003q1_2026q3.csv
출력: research/data/ip1_borrower_*.csv
"""
import csv, os, sys

root = sys.argv[1] if len(sys.argv) > 1 else "research"
d = os.path.join(root, "data")

# 1) 돈을 꺼내는 길별 「새로 빌린 돈의 실질 금리」
#    기존 1순위 잔액 B(금리 r0)를 지키느냐, 통째로 새 금리 r1로 다시 받느냐
#    재융자의 실질 금리 = (r1 * (B + C) - r0 * B) / C
B, r0 = 300_000, 0.035
products = [  # 이름, 표면 금리, 비고
    ("현금인출 재융자(30년 평균 금리)", 0.0711, "Freddie Mac PMMS 2026-11-12"),
    ("현금인출 재융자(+0.25%p 가산 가정)", 0.0736, "현금인출 가산 가정"),
    ("HELOC 은행 평균(변동)", 0.0749, "Bankrate 2026-11-11"),
    ("홈에쿼티론(폐쇄형 2순위)", 0.0803, "Bankrate 2026-05"),
    ("동사형 HELOC(고정, 25년 HMDA 중위)", 0.0910, "HMDA 2025 중위"),
    ("개인 신용대출", 0.1227, "Bankrate 2026-05"),
    ("신용카드", 0.2215, "Fed G.19"),
]
rows = []
for C in (50_000, 90_000):
    for name, r, src in products:
        if name.startswith("현금인출"):
            extra = r * (B + C) - r0 * B
        else:
            extra = r * C
        rows.append([C, name, round(r * 100, 2), round(extra), round(extra / C * 100, 1), src])
with open(os.path.join(d, "ip1_borrower_marginal_rate.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["need_$", "product", "headline_rate_%", "first_year_extra_interest_$", "effective_rate_on_new_money_%", "source"])
    w.writerows(rows)

# 2) 락인 갭: 신규 30년 금리 연평균 vs 기존 대출 평균 금리(FHFA NMDB, 확인된 점만)
fred = [(dt, float(v)) for dt, v in list(csv.reader(open(os.path.join(d, "fred_mortgage30us_weekly.csv"))))[1:] if v not in ("", ".")]
outstanding = {"2022Q1": 3.8, "2025Q2": 4.3, "2025Q4": 4.4, "2026Q2": 4.4}  # FHFA NMDB 집계(2차 인용)
def q_avg(q):
    y, n = q[:4], int(q[-1])
    ms = {1: ("01", "02", "03"), 2: ("04", "05", "06"), 3: ("07", "08", "09"), 4: ("10", "11", "12")}[n]
    s = [v for dt, v in fred if dt[:4] == y and dt[5:7] in ms]
    return sum(s) / len(s)
with open(os.path.join(d, "ip1_borrower_lockin_gap.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["quarter", "new_30y_avg_%", "outstanding_avg_%", "gap_%p"])
    for q, o in outstanding.items():
        n = q_avg(q)
        w.writerow([q, round(n, 2), o, round(n - o, 2)])
    w.writerow(["2026-11-12(주간)", 7.11, 4.4, round(7.11 - 4.4, 2)])

# 기존 평균 금리가 오르는 속도와 신규 금리와 만나는 데 걸리는 시간(단순 외삽)
speed = (4.4 - 3.8) / 4.25  # 1Q22 -> 2Q26, 연 %p
years_to_meet = (7.11 - 4.4) / speed

# 3) 30년 금리가 손익분기(3.4~5.2%) 아래였던 마지막 주
def last_below(th):
    s = [dt for dt, v in fred if v < th]
    return s[-1] if s else None

# 4) HELOC 잔액: 저점 이후 연속 증가 분기 수, 분기 증가액
ny = list(csv.DictReader(open(os.path.join(d, "nyfed_hhdc_heloc_2003q1_2026q3.csv"))))
bal = [(r["quarter"], float(r["heloc_tn"]) * 1000) for r in ny]
streak = 0
for i in range(len(bal) - 1, 0, -1):
    if bal[i][1] > bal[i - 1][1]:
        streak += 1
    else:
        break
with open(os.path.join(d, "ip1_borrower_summary.csv"), "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["metric", "value", "note"])
    w.writerow(["기존 대출 평균 금리 상승 속도(연 %p)", round(speed, 3), "3.8%(1Q22) → 4.4%(2Q26), FHFA NMDB"])
    w.writerow(["이 속도로 신규 7.11%와 만나는 데 걸리는 해", round(years_to_meet, 1), "단순 외삽, 방향만"])
    w.writerow(["30년 금리가 5.2% 아래였던 마지막 주", last_below(5.2), "FRED MORTGAGE30US"])
    w.writerow(["30년 금리가 4.8% 아래였던 마지막 주", last_below(4.8), "FRED"])
    w.writerow(["30년 금리가 3.4% 아래였던 마지막 주", last_below(3.4), "FRED"])
    w.writerow(["HELOC 잔액 연속 증가 분기", streak, "NY Fed HHDC, 3Q26까지"])
    w.writerow(["HELOC 잔액 저점(22Q1) → 26Q3", f"{bal[[q for q,_ in bal].index('2022Q1')][1]:.0f} → {bal[-1][1]:.0f} ($B)", "+%.0f%%" % ((bal[-1][1] / bal[[q for q,_ in bal].index('2022Q1')][1] - 1) * 100)])
    w.writerow(["3Q26 분기 증가액($B)", round(bal[-1][1] - bal[-2][1], 1), "NY Fed"])
    w.writerow(["카드 잔액 26Q3($B)", float(ny[-1]["credit_card_tn"]) * 1000, "NY Fed"])

for fn in ("ip1_borrower_marginal_rate.csv", "ip1_borrower_lockin_gap.csv", "ip1_borrower_summary.csv"):
    print("==", fn)
    print(open(os.path.join(d, fn)).read())
