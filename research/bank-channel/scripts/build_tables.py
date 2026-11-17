"""HMDA·FDIC·SLOOS 원자료에서 은행 채널 보강용 표를 만든다.
입력: HMDA Data Browser 내려받기(기관별 LAR CSV, 전국 2순위 집계 JSON), FDIC RIS API, 연준 SLOOS 표.
실행: python3 -I build_tables.py <hmda_dir> <out_dir>
"""
import csv, json, re, sys, collections, os
H, OUT = sys.argv[1], sys.argv[2]

def amt(r):
    try: return float(r['loan_amount'] or 0)
    except ValueError: return 0.0

# 1) 은행별: 1순위 모기지 vs HELOC 같은 해 매각 비중(금액 기준), 2025
rows_out = []
for key, label in [('citizens','Citizens'),('pnc','PNC'),('fifththird','Fifth Third'),('flagstar','Flagstar'),('valleystrong','Valley Strong CU'),('figure','Figure Lending'),('loandepot','loanDepot')]:
    R = [r for r in csv.DictReader(open(f'{H}/{key}_2025.csv')) if r['reverse_mortgage']!='1' and r['business_or_commercial_purpose']!='1']
    fl = [r for r in R if r['open-end_line_of_credit']!='1' and r['lien_status']=='1']
    he = [r for r in R if r['open-end_line_of_credit']=='1']
    def sold(L):
        t = sum(amt(r) for r in L)
        return (sum(amt(r) for r in L if r['purchaser_type']!='0')/t*100) if t else None
    def gse(L):
        t = sum(amt(r) for r in L)
        return (sum(amt(r) for r in L if r['purchaser_type'] in ('1','2','3'))/t*100) if t else None
    rows_out.append([label, len(fl), round(sum(amt(r) for r in fl)/1e9,2), None if sold(fl) is None else round(sold(fl),1), None if gse(fl) is None else round(gse(fl),1),
                     len(he), round(sum(amt(r) for r in he)/1e9,2), None if sold(he) is None else round(sold(he),1)])
with open(f'{OUT}/hmda2025_sold_share_by_lender.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['lender','first_lien_units','first_lien_$B','first_lien_sold_same_year_%','first_lien_sold_to_GSE_%','heloc_units','heloc_limit_$B','heloc_sold_same_year_%']); w.writerows(rows_out)
for r in rows_out: print(r)

# 2) 전국 2순위: 기관 유형별 점유·매각률, 은행 매입액
names = {}
for y in ['2023','2024','2025']:
    for i in json.load(open(f'{H}/filers{y}.json'))['institutions']: names[i['lei']] = i['name']
def typ(n):
    u = n.upper()
    if re.search(r'CREDIT UNION|\bFCU\b|\bCU\b|FEDERAL C\.?U|C\.U\.', u): return 'credit_union'
    if re.search(r'\bBANK\b|\bBANC|N\.A\.|NATIONAL ASSOCIATION|SAVINGS|\bTRUST\b|\bFSB\b|BANKING|\bS&L\b|LOAN ASSOCIATION|BANCORP', u) and not re.search(r'MORTGAGE (CO|CORP|LLC|INC)|HOME LOANS?,? (LLC|INC)', u): return 'bank'
    return 'nonbank'
nat, buy, lender = [], [], collections.defaultdict(dict)
for y in ['2023','2024','2025']:
    T = collections.defaultdict(lambda: collections.defaultdict(float))
    sellers = collections.defaultdict(float)
    for lei, oe, pt, bus, c, a in json.load(open(f'{H}/sub_{y}.json'))['agg']:
        if bus=='1': continue
        prod = 'HELOC' if oe=='1' else 'closed_end_second'
        t = typ(names.get(lei, lei))
        T[(prod,t)][pt] += a
        if prod=='HELOC': lender[lei][y] = lender[lei].get(y,0)+a
        if pt=='6': sellers[(names.get(lei,lei), prod)] += a
    for prod in ['HELOC','closed_end_second']:
        tot = sum(sum(T[(prod,t)].values()) for t in ['bank','credit_union','nonbank'])
        for t in ['bank','credit_union','nonbank']:
            g = T[(prod,t)]; a = sum(g.values())
            nat.append([y, prod, t, round(a/1e9,2), round(a/tot*100,1), round((a-g['0'])/a*100,1), round(g['6']/1e9,3)])
    h6 = sum(T[('HELOC',t)]['6'] for t in ['bank','credit_union','nonbank'])
    c6 = sum(T[('closed_end_second',t)]['6'] for t in ['bank','credit_union','nonbank'])
    top = sorted(sellers.items(), key=lambda kv: -kv[1])[:5]
    buy.append([y, round(h6/1e9,2), round(c6/1e9,2), round((h6+c6)/1e9,2)] + [f'{n} {p} ${v/1e6:.0f}M' for (n,p),v in top])
with open(f'{OUT}/hmda_second_lien_by_lender_type.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['year','product','lender_type','limit_or_amount_$B','share_%','sold_same_year_%','sold_to_banks_$B']); w.writerows(nat)
with open(f'{OUT}/hmda_second_lien_bought_by_banks.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['year','heloc_$B','closed_end_second_$B','total_$B','top1','top2','top3','top4','top5']); w.writerows(buy)
for b in buy: print(b)

# 3) 주요 기관 2순위 HELOC 실행 한도(금액) 2023~2025
pick = ['CITIZENS BANK, NA','Bank of America NA','PNC BANK N.A.','US BANK, N.A.','TRUIST BANK','Huntington National Bank','TD Bank','Fifth Third Bank, National Association','JPMorgan Chase Bank, NA','M&T BANK','KEYBANK NATIONAL ASSOCIATION','Flagstar Bank NA','COASTAL COMMUNITY BANK','FIGURE LENDING LLC','United Wholesale Mortgage','SPRING EQ, LLC','Navy Federal Credit Union']
inv = {v:k for k,v in names.items()}
with open(f'{OUT}/hmda_heloc_originations_major_lenders.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['lender','type','2023_$B','2024_$B','2025_$B','chg_23_25_%'])
    for n in pick:
        l = inv.get(n); v = lender.get(l, {})
        a = [v.get(y,0)/1e9 for y in ['2023','2024','2025']]
        w.writerow([n, typ(n)] + [round(x,2) for x in a] + [round((a[2]/a[0]-1)*100) if a[0] else ''])

# 4) Flagstar HMDA 2022~2025
with open(f'{OUT}/flagstar_hmda_2022_2025.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['year','heloc_units','heloc_limit_$B','heloc_sold_$B','first_lien_units','first_lien_$B','first_lien_sold_%'])
    for y in ['2022','2023','2024','2025']:
        R = [r for r in csv.DictReader(open(f'{H}/flagstar_{y}.csv')) if r['reverse_mortgage']!='1' and r['business_or_commercial_purpose']!='1']
        he = [r for r in R if r['open-end_line_of_credit']=='1']; fl = [r for r in R if r['open-end_line_of_credit']!='1' and r['lien_status']=='1']
        fa = sum(amt(r) for r in fl)
        w.writerow([y, len(he), round(sum(amt(r) for r in he)/1e9,2), round(sum(amt(r) for r in he if r['purchaser_type']!='0')/1e9,2), len(fl), round(fa/1e9,2), round(sum(amt(r) for r in fl if r['purchaser_type']!='0')/fa*100,1)])

# 5) Figure Lending 매각처 2023~2025
PT = {'0':'not_sold','5':'private_securitizer','6':'bank_or_savings','71':'cu_mortgage_finance_co','8':'affiliate','9':'other'}
with open(f'{OUT}/figure_hmda_purchaser_mix.csv','w',newline='') as f:
    w = csv.writer(f); w.writerow(['year','purchaser_type','units','limit_$M','share_of_$_%'])
    for y in ['2023','2024','2025']:
        R = [r for r in csv.DictReader(open(f'{H}/figure_{y}.csv')) if r['open-end_line_of_credit']=='1' and r['reverse_mortgage']!='1']
        tot = sum(amt(r) for r in R); c = collections.Counter(); a = collections.Counter()
        for r in R: c[r['purchaser_type']] += 1; a[r['purchaser_type']] += amt(r)
        for k,_ in c.most_common(): w.writerow([y, PT.get(k,k), c[k], round(a[k]/1e6), round(a[k]/tot*100,1)])
print('done')
