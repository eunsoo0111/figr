"""HMDA 2022~2025 2순위 실행분(전국 전수)에서 판매자(파트너) 쪽 표를 만든다.
입력: agg_sub.py 가 만든 sub_<연도>.json, filers<연도>.json (HMDA Data Browser API)
실행: python3 -I hmda_seller_tables.py <hmda_dir> <out_dir>
주의: 금액은 한도(HELOC)·원금(폐쇄형). 역모기지·사업 목적 제외. 기관 유형은 이름으로 분류(근사).
"""
import json, sys, re, csv, collections
H, OUT = sys.argv[1], sys.argv[2]
YEARS = (2022, 2023, 2024, 2025)
names = {}
for y in YEARS:
    for i in json.load(open(f'{H}/filers{y}.json'))['institutions']:
        names[i['lei']] = i['name']
CU_HINT = r'CREDIT|\bFCU\b|\bCU\b|SECURITY SERVICE|DFCU|WSECU|ELEMENTS FINANCIAL|COMMUNITY FINANCIAL|JOVIA|CEFCU|SCHOOLSFIRST|GOLDEN 1|STATE EMPLOYEES'
def typ(lei):
    u = names.get(lei, '').upper()
    if re.search(CU_HINT, u): return '신협'
    if re.search(r'BANK|SAVINGS|\bN\.?A\.?$|NATIONAL ASSOCIATION|TRUST COMPANY|BANCORP|\bFSB\b', u): return '은행'
    return '비은행'
D = {y: json.load(open(f'{H}/sub_{y}.json')) for y in YEARS}
# lender-year aggregates
L = collections.defaultdict(lambda: collections.defaultdict(lambda: {'h': 0.0, 'hn': 0, 'c': 0.0, 'hs': 0.0, 'cs': 0.0, 'hp': collections.Counter(), 'cp': collections.Counter()}))
for y in YEARS:
    for lei, oe, pt, n, s in D[y]['A']:
        r = L[lei][y]
        if oe == '1':
            r['h'] += s; r['hn'] += n; r['hp'][pt] += s
            if pt != '0': r['hs'] += s
        else:
            r['c'] += s; r['cp'][pt] += s
            if pt != '0': r['cs'] += s
SIG = {(lei, y): (n, nf, nio, rs / ra if ra else 0) for y in YEARS for lei, oe, n, nf, nio, rs, ra, ns in D[y]['B'] if oe == '1'}
PT = {'0': '미매각', '1': 'Fannie', '2': 'Ginnie', '3': 'Freddie', '5': '민간 증권화', '6': '은행', '71': '신협·모기지·금융회사', '72': '생명보험사', '8': '계열사', '9': '기타'}

# 1) 유형별 HELOC
with open(f'{OUT}/hmda_heloc_by_lender_type_2022_2025.csv', 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['year', 'type', 'heloc_units', 'heloc_limit_$B', 'share_%', 'sold_same_year_%', 'closed_end_2nd_$B'])
    for y in YEARS:
        t = collections.defaultdict(lambda: [0, 0.0, 0.0, 0.0])
        for lei in L:
            r = L[lei][y]; k = typ(lei); t[k][0] += r['hn']; t[k][1] += r['h']; t[k][2] += r['hs']; t[k][3] += r['c']
        tot = sum(v[1] for v in t.values())
        for k in ('은행', '신협', '비은행'):
            v = t[k]; w.writerow([y, k, v[0], round(v[1] / 1e9, 2), round(v[1] / tot * 100, 1), round(v[2] / v[1] * 100, 1), round(v[3] / 1e9, 2)])

# 2) 비은행 HELOC 상위 대출기관
nb = sorted([l for l in L if typ(l) == '비은행'], key=lambda l: -L[l][2025]['h'])
with open(f'{OUT}/hmda_nonbank_heloc_top_lenders.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['lender', 'lei', 'heloc_2022_$M', 'heloc_2023_$M', 'heloc_2024_$M', 'heloc_2025_$M', 'ces_2025_$M', 'heloc_sold_2025_%', 'fixed_rate_%', 'interest_only_%', 'wtd_rate_%', 'top_purchasers_2025'])
    for l in nb[:25]:
        r = L[l]; s = SIG.get((l, 2025), (1, 0, 0, 0)); hp = r[2025]['hp']; tot = sum(hp.values()) or 1
        w.writerow([names.get(l, l), l] + [round(r[y]['h'] / 1e6) for y in YEARS] + [round(r[2025]['c'] / 1e6), round(r[2025]['hs'] / max(r[2025]['h'], 1) * 100, 1),
                    round(s[1] / s[0] * 100), round(s[2] / s[0] * 100), round(s[3], 2), '; '.join(f"{PT.get(k, k)} {v / tot * 100:.0f}%" for k, v in hp.most_common(3))])

# 3) 비은행 폐쇄형 2순위 상위(자기 레일)
nbc = sorted([l for l in L if typ(l) == '비은행'], key=lambda l: -L[l][2025]['c'])
with open(f'{OUT}/hmda_nonbank_closed_end_2nd_top_lenders.csv', 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['lender', 'ces_2022_$M', 'ces_2023_$M', 'ces_2024_$M', 'ces_2025_$M', 'heloc_2025_$M', 'ces_sold_2025_%', 'top_purchasers_2025'])
    for l in nbc[:15]:
        r = L[l]; cp = r[2025]['cp']; tot = sum(cp.values()) or 1
        w.writerow([names.get(l, l)] + [round(r[y]['c'] / 1e6) for y in YEARS] + [round(r[2025]['h'] / 1e6), round(r[2025]['cs'] / max(r[2025]['c'], 1) * 100, 1),
                    '; '.join(f"{PT.get(k, k)} {v / tot * 100:.0f}%" for k, v in cp.most_common(3))])

# 4) 비은행 2순위의 출구(누가 사나), 2025
with open(f'{OUT}/hmda_nonbank_second_lien_exit_2025.csv', 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['product', 'purchaser', '$B', 'share_of_sold_%'])
    for prod, key in (('HELOC', 'hp'), ('폐쇄형', 'cp')):
        agg = collections.Counter()
        for l in L:
            if typ(l) == '비은행': agg.update(L[l][2025][key])
        sold = sum(v for k, v in agg.items() if k != '0')
        for k, v in agg.most_common():
            w.writerow([prod, PT.get(k, k), round(v / 1e9, 2), '' if k == '0' else round(v / sold * 100, 1)])

# 5) 이름이 확인된 동사 파트너(대출기관으로 신고한 HELOC) + Figure Lending 실행·매입
PARTNERS = {  # LEI: (이름, 확인 근거)
 '5493001WHVQBGRSWEU75': ('HomeBridge', '2022 파트너(회사 보도)'), '254900DTLHVWQ7NP7R34': ('The Loan Store', '상위 20 IMB, 2023 파트너'),
 '549300U3721PJGQZYY68': ('Guaranteed Rate(Rate)', '2022 파트너'), '54930001NSTOD85LT125': ('Guaranteed Rate Affinity', 'Rate 계열'),
 '549300VZVN841I2ILS84': ('CrossCountry', '2023 파트너'), '549300GKFNPRWNS0GF29': ('CMG', '2023 파트너'),
 '549300DD4R4SYK5RAQ92': ('Movement', '2025 파트너'), '549300E2UX99HKDBR481': ('New American Funding', 'EC v1·v3'),
 '549300YN94MOAVYW0F52': ('GoodLeap', '2025 파트너'), '549300PQLK8E48ILXV41': ('Valon', 'FIG6 풀 실행사(DBRS)'),
 '549300MGPZBLQDIL7538': ('Fairway', '2023 파트너'), '549300OPCWU6E72WUT29': ('Mutual of Omaha', '1Q26 콜'),
 '2549004FXPNKCLIN8250': ('Synergy One', '2022 파트너'), '5493006MA7WP1WL8U431': ('Valley Strong CU', '화이트라벨(HousingWire)'),
}
with open(f'{OUT}/hmda_figure_partner_cohort.csv', 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['lender', 'basis', 'heloc_2022_$M', 'heloc_2023_$M', 'heloc_2024_$M', 'heloc_2025_$M', 'fixed_rate_2025_%', 'wtd_rate_2025_%', 'top_purchasers_2025'])
    tot = [0.0] * 4
    for l, (nm, why) in PARTNERS.items():
        r = L[l]; s = SIG.get((l, 2025), (1, 0, 0, 0)); hp = r[2025]['hp']; t = sum(hp.values()) or 1
        vals = [r[y]['h'] / 1e6 for y in YEARS]; tot = [a + b for a, b in zip(tot, vals)]
        w.writerow([nm, why] + [round(v) for v in vals] + [round(s[1] / s[0] * 100), round(s[3], 2), '; '.join(f"{PT.get(k, k)} {v / t * 100:.0f}%" for k, v in hp.most_common(2))])
    w.writerow(['합계(14곳)', ''] + [round(v) for v in tot] + ['', '', ''])
print(open(f'{OUT}/hmda_heloc_by_lender_type_2022_2025.csv').read())
print(open(f'{OUT}/hmda_figure_partner_cohort.csv').read())
print(open(f'{OUT}/hmda_nonbank_second_lien_exit_2025.csv').read())
print(open(f'{OUT}/hmda_nonbank_closed_end_2nd_top_lenders.csv').read())

# 6) 비은행 HELOC: 동사(실행) / 이름 확인 파트너 / 나머지, 그리고 판매자 저변
FIG = '254900UL88QFG0E40516'
named_nb = [l for l in PARTNERS if typ(l) == '비은행']
with open(f'{OUT}/hmda_nonbank_heloc_split.csv', 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['year', 'nonbank_heloc_$B', 'figure_lending_originated_$B', 'named_partners_$B', 'other_nonbanks_$B', 'nonbank_lenders_ge_$10M', 'nonbank_lenders_ge_$100M'])
    for y in YEARS:
        nbl = [l for l in L if typ(l) == '비은행']
        tot = sum(L[l][y]['h'] for l in nbl); fg = L[FIG][y]['h']; pn = sum(L[l][y]['h'] for l in named_nb)
        w.writerow([y, round(tot / 1e9, 2), round(fg / 1e9, 2), round(pn / 1e9, 2), round((tot - fg - pn) / 1e9, 2),
                    sum(1 for l in nbl if L[l][y]['h'] >= 1e7), sum(1 for l in nbl if L[l][y]['h'] >= 1e8)])
print(open(f'{OUT}/hmda_nonbank_heloc_split.csv').read())
