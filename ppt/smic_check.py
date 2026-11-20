# smic_check.py — PPT 차트 값 ↔ 원자료 엑셀 1:1 대조 + 양식 검사 + (선택) 렌더
# 사용: python3 smic_check.py 도표.pptx 원자료.xlsx [--render 폴더]
# 결과: 콘솔 보고 + 엑셀에 '대조' 시트(칸 단위 대조표)와 '목록' 시트 '대조 결과' 열을 채운다. 불일치 있으면 종료코드 1.
import os, re, subprocess, sys
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from smic_xlsx import load, INDEX, DG, B00, M00, ARIAL  # noqa: E402
from smic_charts_lib import fmt_value, tw  # noqa: E402

CM = 360000
OK_FONTS = {"AppleSDGothicNeoB00", "AppleSDGothicNeoM00", "AppleSDGothicNeoH00",
            "AppleSDGothicNeoL00", "Arial", "+mn-lt", "+mn-ea", "+mj-lt", "+mj-ea"}
OK_COLORS = {"C00000", "FF7F7F", "FFBFBF", "D0CECE", "7F7F7F", "595959", "E7E6E6", "FFF2CC",
             "000000", "FFFFFF"}


def walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:                      # 그룹
            yield from walk(sh.shapes)


def same(a, b):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= max(1e-9, 1e-9 * max(abs(a), abs(b)))


def chart_data(ch):
    out = []
    for pl in ch.plots:
        cats = pl.categories
        if cats.depth > 1:
            labs = [" ".join(str(x) for x in t) for t in cats.flattened_labels]
        else:
            labs = ["" if str(c) == "None" else str(c) for c in cats]      # 빈 칸 = 'None'으로 읽힘
        for s in pl.series:
            out.append((s.name, labs, list(s.values)))
    return out


def style_issues(ch, nser):
    iss = []
    cs = ch._chartSpace
    x = cs.xml
    if "majorGridlines" in x or "minorGridlines" in x:
        iss.append("눈금선이 있음")
    sp = cs.find(qn("c:spPr"))
    if sp is None or sp.find(qn("a:noFill")) is None:
        iss.append("차트 배경 채우기 있음")
    for f in set(re.findall(r'typeface="([^"]+)"', x)) - OK_FONTS:
        iss.append(f"양식 밖 글꼴 {f}")
    for c in set(re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', x)) - OK_COLORS:
        iss.append(f"팔레트 밖 색 #{c}")
    if "schemeClr" in x.split("<c:plotArea")[1] if "<c:plotArea" in x else False:
        iss.append("테마 색(schemeClr)이 남아 있음 — 계열 색 지정 누락")
    if nser > 1 and ch.has_legend and ch.legend.position is not None and int(ch.legend.position) != -4160:
        iss.append("범례가 상단이 아님")
    if nser > 1 and not ch.has_legend and ch.plots[0].__class__.__name__ != "PiePlot":
        iss.append("계열이 2개 이상인데 범례 없음")
    for ser in cs.iter(qn("c:ser")):
        if ser.getparent().tag == qn("c:lineChart"):
            ln = ser.find(qn("c:spPr") + "/" + qn("a:ln"))
            if ln is None or ln.get("w") not in ("19050",):
                iss.append("꺾은선 두께가 1.5pt가 아님")
                break
    vax = cs.findall(".//" + qn("c:valAx"))
    if len(vax) > 1:
        names = [s.name for pl in ch.plots for s in pl.series]
        if not any("LHS" in n for n in names) or not any("RHS" in n for n in names):
            iss.append("축이 둘인데 계열 이름에 (LHS)/(RHS) 표기가 없음")
    for ax in vax:
        if ax.find(qn("c:title")) is None and ax.find(qn("c:delete")).get("val") == "0":
            iss.append("세로축 단위(축제목) 없음")
            break
    for ax in vax:
        t = ax.find(qn("c:title"))
        nf = ax.find(qn("c:numFmt"))
        if t is None or nf is None:
            continue
        unit = "".join(x.text or "" for x in t.iter(qn("a:t")))
        code = nf.get("formatCode", "").split(";")[0]
        scaled = re.search(r"[0#],+(?![0#])", re.sub(r'"[^"]*"', "", code))
        if re.search(r"천|백만|십억|조|mn|bn|tn|\bM\b|\bB\b", unit) and not scaled:
            iss.append(f"축 단위 '{unit}'인데 표시 형식에 쉼표 축약이 없음 — 원값을 이미 단위 변환해 넣은 것 같음(규칙 2)")
    return iss


def unit_issue(name, vals, fmt):
    body = fmt.split(";")[0]
    m = re.search(r"[0#][0#,.]*", re.sub(r'"[^"]*"', "", body))
    if not m or "%" in body:
        return None
    k = len(m.group(0)) - len(m.group(0).rstrip(","))
    vs = [abs(v) for v in vals if v]
    if k and vs and max(vs) < 1000 ** k:
        return f"'{name}': 표시 형식이 ÷{1000**k:,}인데 원값 최대 {max(vs):,.4g} — 이미 단위를 바꾼 값 같음(원값 1의 자리로)"
    return None


def _norm(t):
    return re.sub(r"\s+", "", t or "")


def check(pptx_path, xlsx_path, render=None):
    book = load(xlsx_path)
    facts = getattr(load, "facts", {})
    prs = Presentation(pptx_path)
    rows, warns, summary = [], [], {}
    used = {k: set() for k in book}
    for si, slide in enumerate(prs.slides, 1):
        shapes = list(walk(slide.shapes))
        title = next((sh.text_frame.text for sh in shapes if sh.has_text_frame and
                      sh.text_frame.text.startswith("도표 ")), None)
        slide_txt = _norm(" ".join(sh.text_frame.text for sh in shapes if sh.has_text_frame) + " " +
                          " ".join(c.text for sh in shapes if getattr(sh, "has_table", False) and sh.has_table
                                   for r in sh.table.rows for c in r.cells))
        no = re.match(r"(도표 [0-9A-Za-z]+-[0-9]+)", title).group(1) if title else None
        guide = next((sh for sh in shapes if sh.name == "GUIDE"), None)
        frames = [sh for sh in shapes if getattr(sh, "has_chart", False) and sh.has_chart]
        tag = no or f"슬라이드 {si}"
        if no and not title.startswith(no + "."):
            warns.append(f"[{tag}] 제목 형식 '도표 N-M. 제목'이 아님")
        if guide is not None:
            gx0, gy0 = guide.left, guide.top
            gx1, gy1 = gx0 + guide.width, gy0 + guide.height
            tol = int(0.03 * CM)
            for sh in shapes:
                if sh is guide or sh.shape_type == 6:
                    continue
                inside_grp = sh._element.getparent().tag == qn("p:grpSp")
                if not inside_grp:
                    continue
                if (sh.left < gx0 - tol or sh.top < gy0 - tol or
                        sh.left + sh.width > gx1 + tol or sh.top + sh.height > gy1 + tol):
                    warns.append(f"[{tag}] '{sh.name}'가 가이드({guide.width/CM:.2f}×{guide.height/CM:.2f}cm) 밖으로 나감")
        for sh in shapes:
            if not sh.has_text_frame or sh._element.getparent().tag != qn("p:grpSp"):
                continue
            xml = sh._element.xml
            for f in set(re.findall(r'typeface="([^"]+)"', xml)) - OK_FONTS:
                warns.append(f"[{tag}] 도표 글자 '{sh.text_frame.text[:12]}…' 양식 밖 글꼴 {f}")
            for z in set(re.findall(r' sz="(\d+)"', xml)) - {"800"}:
                warns.append(f"[{tag}] 도표 글자 '{sh.text_frame.text[:12]}…' {int(z)/100:g}pt — 도표 안 글자는 8pt")
        for gf in frames:
            sheet = gf.name.split(":", 1)[1] if gf.name.startswith("CHART:") else no
            if sheet not in book:
                rows.append((tag, sheet or "?", "(시트 없음)", "", "", "", "", "✗"))
                summary[sheet or tag] = "✗ 엑셀 시트 없음"
                continue
            sh_ = book[sheet]
            bad = 0
            data = chart_data(gf.chart)
            for iss in style_issues(gf.chart, len(data)):
                warns.append(f"[{tag}] {iss}")
            cax = gf.chart._chartSpace.find(".//" + qn("c:catAx"))
            horiz = 'barDir val="bar"' in gf.chart._chartSpace.xml
            if data and cax is not None and gf.chart.plots[0].categories.depth == 1 and not horiz:
                sk = cax.find(qn("c:tickLblSkip"))
                labs = [l for l in data[0][1][::int(sk.get("val")) if sk is not None else 1] if l]
                need = sum(tw(l) + 0.15 for l in labs)
                if need > gf.width / CM * 0.82:
                    warns.append(f"[{tag}] 가로축 레이블이 겹칠 수 있음(필요 {need:.1f}cm) — "
                                 "엑셀 기간을 짧게('24, 26E, 25.06) 쓰거나 skip")
            for name, labs, vals in data:
                if name not in sh_.cols:
                    rows.append((tag, sheet, name, "(엑셀에 열 없음)", "", "", "", "✗")); bad += 1
                    continue
                used[sheet].add(name)
                if labs != sh_.cats and labs != sh_.axis:
                    rows.append((tag, sheet, name, "기간 불일치", " / ".join(sh_.cats), " / ".join(labs), "", "✗"))
                    bad += 1
                ev = sh_.cols[name]
                for i, (lab, pv) in enumerate(zip(sh_.cats, vals)):
                    xv = ev[i] if i < len(ev) else None
                    okv = same(xv, pv)
                    bad += not okv
                    rows.append((tag, sheet, name, lab, xv, pv,
                                 None if (xv is None or pv is None) else pv - xv, "○" if okv else "✗"))
                if len(ev) != len(vals):
                    rows.append((tag, sheet, name, "개수", len(ev), len(vals), "", "✗")); bad += 1
                u = unit_issue(name, ev, sh_.fmt(name))
                if u:
                    warns.append(f"[{tag}] {u}")
            if guide is not None and len(frames) == 1:
                dw = abs(gf.width - guide.width) / CM
                dh = abs(gf.height - guide.height) / CM
                if dw > 0.6 or dh > 0.6:
                    warns.append(f"[{tag}] 차트 크기 {gf.width/CM:.2f}×{gf.height/CM:.2f}가 가이드와 다름")
            summary[sheet] = "○ 일치" if bad == 0 else f"✗ 불일치 {bad}칸"
        for tb in [sh for sh in shapes if getattr(sh, "has_table", False) and sh.has_table
                   and sh.name.startswith("TABLE:")]:
            sheet, *cols = tb.name[6:].split("|")
            if sheet not in book:
                summary[sheet] = "✗ 엑셀 시트 없음"; continue
            sh_, bad, t = book[sheet], 0, tb.table
            for i, cat in enumerate(sh_.cats):
                for j, n in enumerate(cols):
                    want = fmt_value(sh_.cols[n][i], sh_.fmt(n)) if n in sh_.cols else sh_.texts.get(n, [""] * 999)[i]
                    got = t.cell(i + 1, j + 1).text
                    okv = want == got
                    bad += not okv
                    rows.append((tag, sheet, n, cat, want, got, "", "○" if okv else "✗"))
            summary[sheet] = "○ 일치(표)" if bad == 0 else f"✗ 표 불일치 {bad}칸"
        for sh in shapes:                            # 엑셀 값으로 만든 주석 글자
            if not sh.name.startswith("VAL:"):
                continue
            sheet, col, i, f = sh.name[4:].split("|")
            i = int(i)
            if sheet not in book:
                rows.append((tag, sheet, col, "주석", "", sh.text_frame.text, "", "✗")); continue
            S = book[sheet]
            want = fmt_value(S.cols[col][i], f or S.fmt(col)) if col in S.cols else S.texts.get(col, [""] * 999)[i]
            okv = _norm(want) in _norm(sh.text_frame.text)
            rows.append((tag, sheet, col, f"주석 {S.cats[i]}", want, sh.text_frame.text, "", "○" if okv else "✗"))
            prev = summary.get(sheet, "")
            if not okv:
                summary[sheet] = "✗ 주석 불일치"
            elif not prev:
                summary[sheet] = "○ 일치(주석 글자)"
        if no and no in facts:                       # 숫자 없는 도표: 적힌 글자(날짜·이름·숫자)가 엑셀 '표기'와 같은지
            fl = facts[no]
            bad = 0
            for lab, what, src in fl:
                okv = _norm(lab) in slide_txt
                bad += not okv
                rows.append((tag, no, "글자", what[:30], lab, lab if okv else "(도표에 없음)", "", "○" if okv else "✗"))
            summary[no] = ("— 숫자 없음(대조 대상 아님)" if not fl else
                           f"○ 글자 대조 {len(fl)}건" if bad == 0 else f"✗ 글자 불일치 {bad}건")
    for k in facts:
        summary.setdefault(k, "— PPT에 해당 슬라이드 없음")
    for k, s in book.items():
        if k not in summary:
            summary[k] = "— PPT에 차트·표 없음(확인)"
        if any(not x for x in s.sources):
            warns.append(f"[{k}] 출처 열이 빈 행이 있음")
    _write(xlsx_path, rows, summary)
    nbad = sum(1 for r in rows if r[-1] == "✗")
    print(f"■ 대조: {len(rows)}칸 중 불일치 {nbad}칸")
    for k, v in summary.items():
        print(f"  {k}: {v}")
    for r in rows:
        if r[-1] == "✗":
            print("  ✗", r)
    print(f"■ 양식 경고 {len(warns)}건")
    for w in dict.fromkeys(warns):
        print("  !", w)
    if render:
        do_render(pptx_path, render)
    return nbad, warns


def _write(xlsx_path, rows, summary):
    wb = load_workbook(xlsx_path)
    if "대조" in wb.sheetnames:
        del wb["대조"]
    ws = wb.create_sheet("대조", 1)
    head = ["슬라이드", "시트", "계열", "기간", "엑셀 값", "PPT 값", "차이", "결과"]
    ws.append(head)
    for j in range(1, 9):
        c = ws.cell(1, j)
        c.font = Font(name=B00, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=DG)
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(list(r))
        rr = ws.max_row
        for j in (5, 6, 7):
            ws.cell(rr, j).font = Font(name=ARIAL); ws.cell(rr, j).number_format = "#,##0.########"
        if r[-1] == "✗":
            for j in range(1, 9):
                ws.cell(rr, j).fill = PatternFill("solid", fgColor="FFBFBF")
    for col, w in zip("ABCDEFGH", (10, 12, 22, 12, 22, 22, 12, 6)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    idx = wb[INDEX]
    col = [c.value for c in idx[1]].index("대조 결과") + 1
    for r in range(2, idx.max_row + 1):
        k = idx.cell(r, 1).value
        if k in summary:
            idx.cell(r, col).value = summary[k]
    wb.calculation.fullCalcOnLoad = True
    wb.save(xlsx_path)


def do_render(pptx_path, outdir):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx_path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    subprocess.run(["pdftoppm", "-png", "-r", "110", pdf, os.path.join(outdir, "slide")], timeout=300)
    print(f"■ 렌더: {outdir}/slide-*.png — 가이드 부분을 확대해서 눈으로 확인")


if __name__ == "__main__":
    a = sys.argv[1:]
    r = a[a.index("--render") + 1] if "--render" in a else None
    nbad, _ = check(a[0], a[1], r)
    sys.exit(1 if nbad else 0)
