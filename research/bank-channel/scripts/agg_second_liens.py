import csv, sys, json, collections
year=sys.argv[1]; out=sys.argv[2]
agg=collections.defaultdict(lambda:[0,0.0])
r=csv.reader(sys.stdin)
h=next(r); ix={k:i for i,k in enumerate(h)}
L=ix['lei'];O=ix['open-end_line_of_credit'];P=ix['purchaser_type'];A=ix['loan_amount'];RV=ix['reverse_mortgage'];B=ix['business_or_commercial_purpose']
n=0
for row in r:
    n+=1
    if row[RV]=='1': continue
    try: amt=float(row[A])
    except: amt=0.0
    k=(row[L],row[O],row[P],row[B])
    v=agg[k]; v[0]+=1; v[1]+=amt
json.dump({'year':year,'rows':n,'agg':[[*k,*v] for k,v in agg.items()]},open(out,'w'))
print(year,n)
