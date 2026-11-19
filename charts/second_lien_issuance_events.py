# -*- coding: utf-8 -*-
"""wide 18.5×5.1: 2순위(HELOC+CES) 증권화 발행액 막대(LHS) + HELOC 잔액 선(RHS) + 주요 사건.

데이터: data/second_lien_securitization_issuance.csv (발행액·출처) · data/B_heloc_balance_annual_2004_2026.csv (잔액)
실행:  python charts/second_lien_issuance_events.py  → outputs/Ex_second_lien_issuance_events.{png,svg}
"""
import csv, os, sys, glob
HC = os.environ.get('HOUSE_CHART_LIB') or next(iter(glob.glob(os.path.expanduser('~/.claude/skills/**/house-chart/lib'), recursive=True)), '')
sys.path.insert(0, HC)
from house_style import *

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ISS = list(csv.DictReader(open(f'{ROOT}/data/second_lien_securitization_issuance.csv', encoding='utf-8-sig')))
BAL = list(csv.DictReader(open(f'{ROOT}/data/B_heloc_balance_annual_2004_2026.csv', encoding='utf-8-sig')))

years = [int(r['year']) for r in BAL]
x0 = years[0]; X = lambda y: y - x0
row = lambda y, k: next(r for r in ISS if int(r['year']) == y and r['kind'] == k)

fig, ax, ax2 = new_fig('wide', twin=True)
W = 0.62
COL = {'actual': MID, 'derived': MID, 'deal_sum': LIGHT}
for r in ISS:
    y, k = int(r['year']), r['kind']
    if k in COL:
        ax.bar(X(y), float(r['issuance_bn']), width=W, color=COL[k], lw=0, zorder=3)
# 2026: YTD(실적) + 연간 전망까지의 잔여분(빗금)
ytd, fc = float(row(2026, 'ytd')['issuance_bn']), float(row(2026, 'forecast')['issuance_bn'])
ax.bar(X(2026), ytd, width=W, color=MID, lw=0, zorder=3)
ax.bar(X(2026), fc - ytd, bottom=ytd, width=W, color='white', edgecolor=MID, hatch='//////', lw=0.5, zorder=3)
# 2008~18 사실상 0: 바닥 짧은 선
nz = [int(r['year']) for r in ISS if r['kind'] == 'near_zero']
ax.plot([X(min(nz)) - W / 2, X(max(nz)) + W / 2], [0.35, 0.35], color=DARK, lw=0.8, zorder=4, solid_capstyle='butt')
ax.text((X(min(nz)) + X(max(nz))) / 2, 2.6, '2008~18 HELOC·CES 증권화 사실상 0', ha='center', va='bottom', fontsize=PT)
# 2006·07 연간 합계 미확보
for y in (2006, 2007):
    ax.text(X(y), 1.0, 'n.a.', ha='center', va='bottom', fontsize=PT, color=MID)

# 값 라벨: 처음·정점·최근
def vlab(y, v, txt=None, dy=1.2):
    ax.text(X(y), v + dy, txt or f'{v:,.1f}', ha='center', va='bottom', fontsize=PT, zorder=6)
for y in (2004, 2005, 2023, 2024, 2025):
    r = next(r for r in ISS if int(r['year']) == y); vlab(y, float(r['issuance_bn']))
ax.text(X(2026), (fc + ytd) / 2, f'{fc:,.0f}E', ha='center', va='center', fontsize=PT, zorder=6, bbox=dict(facecolor='white', edgecolor='none', pad=0.2))
ax.text(X(2026), ytd / 2, f'{ytd:,.1f}', ha='center', va='center', fontsize=PT, color='white', zorder=6)
for y in (2019, 2020, 2021, 2022):
    r = next(r for r in ISS if int(r['year']) == y); v = float(r['issuance_bn'])
    vlab(y, v, f'{v:,.1f}' if v >= 1 else f'{v:,.2f}', dy=0.8)

# HELOC 잔액(RHS, CSV)
bal = [float(r['heloc_balance_bn']) for r in BAL]
ax2.plot([X(y) for y in years], bal, color=RED, lw=1.0, marker='o', ms=1.8, zorder=5)

# 주요 사건: (연도, 문구, 높이[LHS 단위], 정렬)
EV = [
    (2005, '4Q05 2순위 신규취급 정점', 112, 'left'),
    (2007, '서브프라임 붕괴 → 2순위 증권화 중단', 103, 'left'),
    (2008, '모노라인 붕괴(MBIA·Ambac 2순위 보증손실) · 리먼(9월)', 94, 'left'),
    (2019, 'FirstKey(Cerberus) 위기 후 첫 all-HELOC 딜', 67, 'right'),
    (2020, 'Figure 원천 HELOC 첫 증권화(GRADE 2020-FIG1)', 76, 'right'),
    (2022, '연준 급격한 인상 → 1순위 금리 잠김', 85, 'right'),
    (2023, 'Figure 첫 등급 HELOC ABS(FIGRE 2023-HE1)', 94, 'right'),
    (2025, 'Figure IPO(9월) · HELOC 증권화 누적 19건 USD 5.8bn', 103, 'right'),
]
top = {int(r['year']): float(r['issuance_bn'] or 0) for r in ISS if r['kind'] in ('actual', 'derived', 'deal_sum', 'near_zero')}
top[2006] = top[2007] = 0
for y, t, h, side in EV:
    base = top.get(y, 0) + (7.5 if top.get(y, 0) >= 1 else 3.0)
    ax.plot([X(y), X(y)], [base, h - 0.6], color=BLACK, lw=0.35, ls=(0, (1.5, 1.2)), zorder=2)
    ax.plot([X(y)], [h], marker='o', ms=1.6, color=BLACK, zorder=6)
    dx = 0.18 if side == 'left' else -0.18
    ax.text(X(y) + dx, h, t, ha=side, va='center', fontsize=PT, zorder=7,
            bbox=dict(facecolor='white', edgecolor='none', pad=0.3))

ax.set_xticks([X(y) for y in years]); ax.set_xticklabels([f'{y}' for y in years])
ax.set_xlim(-0.6, len(years) - 0.4)
ax.set_ylim(0, 120); ax.set_yticks(range(0, 121, 20))
ax2.set_ylim(0, 1200); ax2.set_yticks(range(0, 1201, 200))

h = [Patch(facecolor=MID), Patch(facecolor=LIGHT), Patch(facecolor='white', edgecolor=MID, hatch='//////'),
     Line2D([], [], color=RED, marker='o', ms=1.8, lw=1.0)]
l = ['2순위(HELOC+CES) 증권화 발행액(LHS)', '확인된 딜 합계·하한(LHS)', '2026 연간 전망(S&P, LHS)', 'HELOC 잔액(RHS)']
legend(fig, h, l, ncol=4)
finish(fig, ax, ax2, '(USD bn)', '(USD bn)')

out = f'{ROOT}/outputs/Ex_second_lien_issuance_events.png'
png, svg = save(fig, out)
bad = check_svg(svg, kind='wide')
print('saved', png, '| check:', bad or 'OK')
