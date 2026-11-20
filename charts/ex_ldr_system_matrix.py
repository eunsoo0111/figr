# -*- coding: utf-8 -*-
"""table 9×3.6: 은행 예대율(낮음/높음) × 자체 취급 시스템(있음/없음) 2×2 — 은행 유형별 Figure 접점."""
import sys, os, json
sys.path.insert(0, os.environ['HOUSE_LIB'])
from house_style import *
HERE = os.path.dirname(os.path.abspath(__file__)); D = json.load(open(os.path.join(HERE, 'ldr_system_matrix.json'), encoding='utf-8'))
fig, ax, _ = new_fig('table'); w, h = SIZES['table']; ax.set_xlim(0, w); ax.set_ylim(0, h); ax.axis('off')
x0, x1 = 0.10, w - 0.10; xl = 1.55; cw = (x1 - xl) / 2; xs = [xl, xl + cw]          # 라벨 열 + 두 칸
top = h - 0.08; hh = 0.62; rh = 1.22; LS = 0.25                                       # 머리행·본문행 높이, 줄 간격(cm)
# 머리행
ax.add_patch(plt.Rectangle((x0, top - hh), x1 - x0, hh, facecolor='#E9E9E9', edgecolor='none'))
for c, xc in zip(D['cols'], xs):
    ax.text(xc + 0.08, top - hh / 2 + 0.12, c['head'], fontweight='bold', ha='left', va='center')
    ax.text(xc + 0.08, top - hh / 2 - 0.13, c['sub'], ha='left', va='center')
ax.plot([x0, x1], [top - hh, top - hh], color=BLACK, lw=0.5)
y = top - hh
for i, r in enumerate(D['rows']):
    yc = y - rh / 2
    ax.text(x0 + 0.05, yc + 0.12, r['head'], fontweight='bold', ha='left', va='center')
    ax.text(x0 + 0.05, yc - 0.13, r['sub'], ha='left', va='center')
    for c, xc in zip(D['cols'], xs):
        k = f"{r['key']}|{c['key']}"; cell = D['cells'][k]; hi = k == D['highlight']
        if hi: ax.add_patch(plt.Rectangle((xc, y - rh), cw, rh, facecolor='#F5F5F5', edgecolor='none', zorder=0))
        yy = yc + LS
        ax.text(xc + 0.08, yy, cell['role'], fontweight='bold', ha='left', va='center', color=RED if hi else BLACK)
        for ln in cell['lines']: yy -= LS; ax.text(xc + 0.08, yy, ln, ha='left', va='center')
    y -= rh
    if i < len(D['rows']) - 1: ax.plot([x0, x1], [y, y], color=LIGHT, lw=0.35)
ax.plot([x0, x1], [y, y], color=BLACK, lw=0.5)
ax.plot([xl, xl], [y, top], color=LIGHT, lw=0.35); ax.plot([xl + cw, xl + cw], [y, top], color=LIGHT, lw=0.35)
ax.text(x0 + 0.05, y - 0.17, D['title_note'], ha='left', va='center')
fig.canvas.draw(); rr = fig.canvas.get_renderer()
ext = lambda t: t.get_window_extent(rr)
over = [t.get_text()[:20] for t in ax.texts if ext(t).x1 / fig.dpi * 2.54 > x1]; assert not over, ('오른쪽 넘침', over)
# 칸 안 넘침: 각 칸 글자가 다음 칸 경계를 넘지 않는지
for t in ax.texts:
    x = ext(t).x0 / fig.dpi * 2.54; r_ = ext(t).x1 / fig.dpi * 2.54
    edge = xl if x < xl else (xl + cw if x < xl + cw else x1)
    if t.get_text() != D['title_note']: assert r_ <= edge - 0.03, ('칸 넘침', t.get_text())
assert ext(ax.texts[-1]).y0 >= 0, '아래 넘침'
png, svg = save(fig, os.path.join(os.path.dirname(HERE), 'outputs', 'Exhibit_LDR_system_matrix.png'))
print('saved', png, '| check:', check_svg(svg, kind='table') or 'OK')
