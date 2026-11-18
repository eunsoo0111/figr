"""Stream HMDA nationwide subordinate-lien originations (CSV on stdin) and aggregate by lender.
Key A: (lei, open_end, purchaser_type) -> [count, sum_amount]   (excl. reverse mortgages and business purpose)
Key B: (lei, open_end) -> product signature: n, n_fixed(no intro rate), n_io, rate*amt sum, amt with rate, n_term_le20y
"""
import csv, sys, json, collections
year, out = sys.argv[1], sys.argv[2]
A = collections.defaultdict(lambda: [0, 0.0])
B = collections.defaultdict(lambda: [0, 0, 0, 0.0, 0.0, 0])
r = csv.reader(sys.stdin); h = next(r); ix = {k: i for i, k in enumerate(h)}
L, O, P, AM, RV, BU = ix['lei'], ix['open-end_line_of_credit'], ix['purchaser_type'], ix['loan_amount'], ix['reverse_mortgage'], ix['business_or_commercial_purpose']
IR, IO, RT, TM = ix['intro_rate_period'], ix['interest_only_payment'], ix['interest_rate'], ix['loan_term']
n = 0
for row in r:
    n += 1
    if row[RV] == '1' or row[BU] == '1':
        continue
    try: amt = float(row[AM])
    except ValueError: amt = 0.0
    a = A[(row[L], row[O], row[P])]; a[0] += 1; a[1] += amt
    b = B[(row[L], row[O])]; b[0] += 1
    if row[IR] in ('NA', 'Exempt', ''): b[1] += 1
    if row[IO] == '1': b[2] += 1
    try:
        rt = float(row[RT]); b[3] += rt * amt; b[4] += amt
    except ValueError: pass
    try:
        if float(row[TM]) <= 240: b[5] += 1
    except ValueError: pass
json.dump({'year': year, 'rows': n, 'A': [[*k, *v] for k, v in A.items()], 'B': [[*k, *v] for k, v in B.items()]}, open(out, 'w'))
print(year, n, flush=True)
