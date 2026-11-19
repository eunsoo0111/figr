# -*- coding: utf-8 -*-
"""wide 18.5×5.1 — 2순위(HELOC+CES) RMBS 발행액(막대, LHS) + HELOC 잔액(분기 선, RHS) + 주요 사건.

정본: data/C_second_lien_rmbs_issuance_annual_2004_2026.csv (발행액·등급·출처키)
      data/HELOC_2004_2026_rawdata.xlsx  A_잔액_분기(잔액) · E_2026_YTD(연간 전망) · F_증권화잔액·발행사(FIGRE 점유)
실행:  python charts/second_lien_issuance_events.py → outputs/Ex_second_lien_issuance_events.{png,svg}
"""
import csv, os, sys, glob
import openpyxl
HC = os.environ.get('HOUSE_CHART_LIB') or next(iter(glob.glob(os.path.expanduser('~/.claude/skills/**/house-chart/lib'), recursive=True)), '')
sys.path.insert(0, HC)
from house_style import *

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
C = {int(r['year']): r for r in csv.DictReader(open(f'{ROOT}/data/C_second_lien_rmbs_issuance_annual_2004_2026.csv', encoding='utf-8-sig'))}
WB = openpyxl.load_workbook(f'{ROOT}/data/HELOC_2004_2026_rawdata.xlsx', data_only=True)
QTR = [(r[0], float(r[1])) for r in WB['A_잔액_분기'].iter_rows(min_row=2, values_only=True)
       if isinstance(r[0], str) and len(r[0]) == 6 and r[0][4] == 'Q' and r[0] >= '2004Q1']
FC = next(float(str(r[2]).lstrip('~>')) for r in WB['E_2026_YTD'].iter_rows(min_row=2, values_only=True) if r[1] == 'S&P Global Ratings')
FIGRE = next(r[1] for r in WB['F_증권화잔액·발행사'].iter_rows(min_row=2, values_only=True) if str(r[0]).startswith('FIGRE'))

Y0, Y1 = 2004, 2026
X = lambda y: y - Y0
xq = lambda q: X(int(q[:4])) + (int(q[5]) - 2.5) * 0.22          # 분기를 해당 연도 칸 안에 펼친다
val = lambda y: float(C[y]['chart_value_bn']) if C[y]['chart_value_bn'] else None
DORM = [y for y in C if '휴면' in C[y]['grade']]                   # 2008~2021

fig, ax, ax2 = new_fig('wide', twin=True)
ax.set_xlim(-0.6, X(Y1) + 0.6); YMAX, RMAX = 160, 1200                                              # 두 축 눈금 4칸으로 맞춤
U = YMAX / 100                                                      # 높이 1% = U
ax.set_ylim(0, YMAX); ax.set_yticks(range(0, YMAX + 1, 40))
ax2.set_ylim(0, RMAX); ax2.set_yticks(range(0, RMAX + 1, 300))

# ── 휴면기 음영 ───────────────────────────────────────────────
d0, d1 = X(min(DORM)) - 0.5, X(max(DORM)) + 0.5
ax.axvspan(d0, d1, color='#F2F2F2', lw=0, zorder=0)
ax.text((d0 + d1) / 2, 3.2 * U, f'휴면기 {min(DORM)}~{str(max(DORM))[2:]}: 연도별 공개 발행치 없음',
        ha='center', va='center', fontsize=PT, color=DARK, zorder=2)

# ── 발행액 막대 ──────────────────────────────────────────────
W = 0.6
for y in range(Y0, Y1 + 1):
    v = val(y)
    if v is None or y in DORM:
        continue
    if y == Y1:                                                     # 2026: YTD + 연간 전망까지 빗금
        ax.bar(X(y), v, width=W, color=MID, lw=0, zorder=3)
        ax.bar(X(y), FC - v, bottom=v, width=W, color='white', edgecolor=MID, hatch='//////', lw=0.5, zorder=3)
        ax.text(X(y), v / 2, f'{v:,.1f}', ha='center', va='center', fontsize=PT, color='white', zorder=6)
        ax.text(X(y), FC + 1.0 * U, f'~{FC:,.0f}E', ha='center', va='bottom', fontsize=PT, zorder=6)
        continue
    pre = y < min(DORM)                                             # 위기 전 막대는 잔액 선이 지나가므로 라벨을 막대 안에
    ax.bar(X(y), v, width=W, color=DARK if pre else MID, lw=0, zorder=3)
    ax.text(X(y), v / 2 if pre else v + 1.0 * U, f'{v:,.1f}', ha='center', va='center' if pre else 'bottom',
            fontsize=PT, color='white' if pre else BLACK, zorder=6)
for y in (2006, 2007):                                              # 연간 총액 원문 미확보(분기 단편만)
    ax.text(X(y), 1.0 * U, 'n.a.', ha='center', va='bottom', fontsize=PT, color=MID, zorder=6)

# ── HELOC 잔액(분기) ────────────────────────────────────────
xs, ys = [xq(q) for q, _ in QTR], [v for _, v in QTR]
ax2.plot(xs, ys, color=RED, lw=1.0, zorder=5, solid_joinstyle='round')
pk = max(range(len(ys)), key=lambda i: ys[i])
lo = min((i for i, (q, _) in enumerate(QTR) if q >= '2012Q1'), key=lambda i: ys[i])
for i, txt, dx, dy, ha in ((pk, f'{ys[pk]:,.0f} 정점({QTR[pk][0][2:]})', 0.2, 0.025 * RMAX, 'left'),
                           (lo, f'{ys[lo]:,.0f} 저점({QTR[lo][0][2:]})', 0, -0.04 * RMAX, 'center'),
                           (len(ys) - 1, f'{ys[-1]:,.0f}', 0, 0.045 * RMAX, 'center')):
    ax2.plot(xs[i], ys[i], marker='o', ms=2.2, color=RED, mec='white', mew=0.4, zorder=6)
    ax2.text(xs[i] + dx, ys[i] + dy, txt, ha=ha, va='center', fontsize=PT, color=RED, zorder=6)

# ── 사건: 번호 마커 + 휴면기 음영 안 목록(Figure 관련 = 검은 원) ─────
EV = [  # (연도, 문구, Figure 관련)
    (2005, '4Q05 2순위 신규취급 정점', False),
    (2007, '서브프라임 붕괴 → 2순위 증권화 중단', False),
    (2008, '리먼 파산 · 모노라인(MBIA·Ambac) 붕괴', False),
    (2019, 'FirstKey, 위기 후 첫 all-HELOC 딜', False),
    (2020, 'Figure 원천 HELOC 첫 증권화(GRADE)', True),
    (2022, '연준 급격한 인상 → 1순위 금리 잠김', False),
    (2023, 'FIGRE 첫 등급딜 · 2순위 RMBS 재개', True),
    (2025, 'Figure IPO(9월)', True),
    (Y1, f'FIGRE, HELOC RMBS 잔액 점유 {FIGRE:.0f}%(26년 9월)', True),
]
q4 = {int(q[:4]): v for q, v in QTR if q.endswith('Q4')}; q4[Y1] = ys[-1]
def badge(x, y, n, filled):
    ax.plot(x, y, marker='o', ms=6.2, mfc=BLACK if filled else 'white', mec=BLACK, mew=0.5, zorder=8, clip_on=False)
    ax.text(x, y, str(n), ha='center', va='center', fontsize=PT, color='white' if filled else BLACK, zorder=9)
for n, (y, t, f) in enumerate(EV, 1):
    v = val(y) or 0
    top = max(v + (2 if y < min(DORM) else 6) * U if v >= 1 else 0, (FC + 6 * U) if y == Y1 else 0, q4[y] * YMAX / RMAX)
    badge(X(y), top + (12 if y == Y1 else 7) * U, n, f)
COLS = [(X(2008) + 0.45, EV[:4], 1), (X(2014) + 0.85, EV[4:], 5)]
for x0c, items, n0 in COLS:
    for k, (y, t, f) in enumerate(items):
        yy = (94 - k * 6.9) * U
        badge(x0c, yy, n0 + k, f)
        ax.text(x0c + 0.33, yy, t, ha='left', va='center', fontsize=PT, zorder=9)

ax.set_xticks([X(y) for y in range(Y0, Y1 + 1)]); ax.set_xticklabels([str(y) for y in range(Y0, Y1 + 1)])
ax.tick_params(axis='x', length=0)

h = [Patch(facecolor=DARK), Patch(facecolor=MID), Patch(facecolor='white', edgecolor=MID, hatch='//////'),
     Line2D([], [], color=RED, lw=1.0), Line2D([], [], ls='', marker='o', ms=4, mfc=BLACK, mec=BLACK)]
l = ['2순위 RMBS 발행 · 위기 전(LHS)', '2순위 RMBS 발행 · 2022~(LHS)', '2026 연간 전망 S&P(LHS)', 'HELOC 잔액, 분기(RHS)', 'Figure 관련 사건']
legend(fig, h, l, ncol=5)
finish(fig, ax, ax2, '(USD bn)', '(USD bn)')
ax.tick_params(axis='x', length=0)

png, svg = save(fig, f'{ROOT}/outputs/Ex_second_lien_issuance_events.png')
print('saved', png, '| check:', check_svg(svg, kind='wide') or 'OK')
