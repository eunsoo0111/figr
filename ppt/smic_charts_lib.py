# smic_charts_lib.py — SMIC 3팀 양식 네이티브 PPT 도표. 값은 반드시 smic_xlsx.load()로 읽은 엑셀에서 넣는다.
import math, re
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

RED, PINK2, PINK, LG, MG, DG = "C00000", "FF7F7F", "FFBFBF", "D0CECE", "7F7F7F", "595959"
BG, YEL, BLK, WHT = "E7E6E6", "FFF2CC", "000000", "FFFFFF"
PAL = [RED, PINK2, PINK, LG, MG, DG]            # 계열 기본 순서(누적·원그래프: 빨강 진→연, 회색 연→진)
B00, M00, H00, ARIAL = "AppleSDGothicNeoB00", "AppleSDGothicNeoM00", "AppleSDGothicNeoH00", "Arial"
CT, CM = 8, 360000
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SIZES = {"기본": (9.25, 5.10), "대형": (18.50, 7.30)}
GX, GY = 1.60, 2.90                               # 가이드 좌상단(cm)
# 표시 형식: 원값은 그대로, 단위는 쉼표로만 (쉼표 1개 = ÷1,000). 음수는 괄호.
FMT = {"원": "#,##0;(#,##0)", "천": "#,##0,;(#,##0,)", "백만": "#,##0,,;(#,##0,,)",
       "십억": "#,##0.0,,,;(#,##0.0,,,)", "조": "#,##0.0,,,,;(#,##0.0,,,,)",
       "%": "0.0%;(0.0%)", "%0": "0%", "배": '0.0"x"', "소수1": "#,##0.0;(#,##0.0)"}


def E(v):
    return Emu(int(round(v * CM)))


def _el(tag, **attrs):
    e = etree.Element(qn(tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def _txpr(size=CT, color=DG, bold=False, font=B00):
    return etree.fromstring(
        f'<c:txPr xmlns:c="{C}" xmlns:a="{A}"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
        f'<a:defRPr sz="{int(size*100)}" b="{1 if bold else 0}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/></a:defRPr>'
        f'</a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>')


def _ln(width_pt=0.75, color=DG):
    return etree.fromstring(
        f'<c:spPr xmlns:c="{C}" xmlns:a="{A}"><a:ln w="{int(width_pt*12700)}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>')


def _insert_before(parent, el, tags):
    for t in tags:
        x = parent.find(qn(t))
        if x is not None:
            x.addprevious(el)
            return el
    parent.append(el)
    return el


# ───────────── 숫자 표시 (축 폭·레이블 계산용, 엑셀 형식 일부 흉내) ─────────────
def fmt_value(v, fmt="#,##0"):
    if v is None:
        return ""
    secs = fmt.split(";")
    f = secs[1] if (v < 0 and len(secs) > 1) else secs[0]
    neg = v < 0 and len(secs) == 1
    x = abs(v) if (v < 0) else v
    lit = "".join(re.findall(r'"([^"]*)"', f))
    core = re.sub(r'"[^"]*"', "", f)
    if "%" in core:
        x *= 100
    m = re.search(r"[0#][0#,.]*", core)
    body = m.group(0) if m else "0"
    trail = len(body) - len(body.rstrip(","))
    x /= 1000 ** trail
    dec = len(body.split(".")[1].rstrip(",")) if "." in body else 0
    s = f"{x:,.{dec}f}" if "," in body.rstrip(",") else f"{x:.{dec}f}"
    pre = core[:m.start()] if m else ""
    post = core[m.end():] if m else ""
    out = pre.replace("(", "(") + s + post + lit
    return ("-" + out) if neg else out


def nice_scale(vals, ticks=5, zero=True, head=0.0, steps=(1, 2, 2.5, 5, 10)):
    """보기 좋은 (min, max, 눈금 간격). head = 위쪽 여유(레이블·주석 자리, 0.2 = 20%)"""
    vs = [v for v in vals if v is not None]
    lo, hi = min(vs), max(vs)
    if zero:
        lo, hi = min(0, lo), max(0, hi)
    hi += (hi - lo) * head
    if hi == lo:
        hi = lo + 1
    raw = (hi - lo) / ticks
    mag = 10 ** math.floor(math.log10(raw))
    step = next(s * mag for s in steps if s * mag >= raw)
    mn = math.floor(lo / step) * step
    mx = math.ceil(hi / step) * step
    if not zero and mn == lo and lo != 0:
        mn -= step
    return mn, mx, step


def fit_scale(vals, k, zero=True, head=0.0):
    """눈금 개수를 k개로 맞춘 scale — 두 축 눈금 높이 맞추기용"""
    vs = [v for v in vals if v is not None]
    lo, hi = min(vs), max(vs)
    if zero:
        lo, hi = min(0, lo), max(0, hi)
    hi += (hi - lo) * head
    raw = (hi - lo) / k if hi > lo else 1
    mag = 10 ** math.floor(math.log10(raw))
    for m in (mag, mag * 10):
        for st in (1, 2, 2.5, 5):
            step = st * m
            mn = math.floor(lo / step) * step
            if math.ceil(hi / step) * step <= mn + k * step + 1e-12:
                return (mn, mn + k * step, step)
    return nice_scale(vals, k, zero, head)


def axis_fmt(fmt, step):
    """눈금 간격이 표시 단위로 정수면 축 형식에서 소수점을 뗀다(10,000.0 → 10,000)"""
    body = fmt.split(";")[0]
    m = re.search(r"[0#][0#,.]*", re.sub(r'"[^"]*"', "", body))
    if not m or "." not in m.group(0):
        return fmt
    k = len(m.group(0)) - len(m.group(0).rstrip(","))
    disp = step / 1000 ** k * (100 if "%" in body else 1)
    if abs(disp - round(disp)) > 1e-9:
        return fmt
    return re.sub(r"\.0+", "", fmt)


_FONTS = {}


def tw(s, pt=CT):
    """렌더 대체 글꼴(Inter + WenQuanYi/Noto) 기준 글자 폭(cm) — Apple SD Gothic보다 넓어 보수적"""
    from PIL import ImageFont
    if not _FONTS:
        import glob
        lat = (glob.glob("/usr/share/fonts/**/Inter-Regular.otf", recursive=True) +
               glob.glob("/usr/share/fonts/**/LiberationSans-Regular.ttf", recursive=True))
        cjk = (glob.glob("/usr/share/fonts/**/wqy-zenhei.ttc", recursive=True) +
               glob.glob("/usr/share/fonts/**/NotoSansCJK-Regular.ttc", recursive=True))
        _FONTS["lat"] = ImageFont.truetype(lat[0], 1000) if lat else None
        _FONTS["cjk"] = ImageFont.truetype(cjk[0], 1000) if cjk else None
    tot = 0.0
    for ch in str(s):
        wide = ord(ch) > 0x2000 or ch in "·×○△"
        f = _FONTS["cjk"] if wide else _FONTS["lat"]
        tot += f.getlength(ch) if f else (1000 if wide else 600)
    return tot / 1000 * pt / 72 * 2.54 * 1.06      # 렌더 대체 글꼴의 자간·한글 간격 여유


def circled(n):
    return chr(0x2460 + n - 1)          # ① ② ③ …


# ───────────── 덱 · 슬라이드 ─────────────
def new_deck(template=None):
    """양식 pptx가 있으면 열어서 슬라이드를 모두 지우고 테마만 쓴다. 없으면 빈 16:9(33.867×19.05)."""
    prs = Presentation(template) if template else Presentation()
    sld = prs.slides._sldIdLst
    for sid in list(sld):
        prs.part.drop_rel(sid.get(qn("r:id")))
        sld.remove(sid)
    prs.slide_width, prs.slide_height = E(33.867), E(19.05)
    return prs


def _blank(prs):
    for l in prs.slide_layouts:
        if l.name in ("빈 화면", "Blank"):
            return l
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def strip_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def font_xml(rPr, size=None, color=None, font=B00, bold=False):
    if size is not None:
        rPr.set("sz", str(int(round(size * 100))))
    rPr.set("b", "1" if bold else "0")
    rPr.set("lang", "ko-KR")
    for tag in ("a:solidFill", "a:latin", "a:ea", "a:cs"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    if color:
        sf = etree.SubElement(rPr, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    for tag in ("a:latin", "a:ea", "a:cs"):
        etree.SubElement(rPr, qn(tag)).set("typeface", font)


def _fill_runs(tf, paras, size, color, align, spacing, font, bold):
    """paras: str(줄바꿈 \\n) 또는 [para]; para = str 또는 [(글자, 색, 크기, 글꼴, 굵게)]"""
    if isinstance(paras, str):
        paras = paras.split("\n")
    for k, para in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pPr.set("algn", {"l": "l", "c": "ctr", "r": "r"}[align])
        ln = etree.SubElement(pPr, qn("a:lnSpc"))
        etree.SubElement(ln, qn("a:spcPct")).set("val", str(int(spacing * 100000)))
        runs = [(para,)] if isinstance(para, str) else para
        for it in runs:
            r = p.add_run()
            r.text = it[0]
            g = lambda i, d: it[i] if len(it) > i and it[i] is not None else d
            font_xml(r._r.get_or_add_rPr(), g(2, size), g(1, color), g(3, font), g(4, bold))


def slide_text(slide, x, y, w, h, text, size, color=BLK, font=B00):
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    bp = tb.text_frame._txBody.find(qn("a:bodyPr"))
    bp.set("wrap", "square")
    for k in ("lIns", "rIns", "tIns", "bIns"):
        bp.set(k, "0")
    _fill_runs(tb.text_frame, text, size, color, "l", 1.0, font, False)
    return tb


def chart_slide(prs, no, title, size="기본", source="SMIC 3팀", notes="", note_lines=(), concept=False):
    """도표 1개 = 슬라이드 1장. 제목 = 빨강 '도표 N-M.' + 검정 제목(개념도면 '(개념도)').
    반환 Canvas — 좌표는 가이드 좌상단 기준 cm, 크기는 cv.w × cv.h"""
    s = prs.slides.add_slide(_blank(prs))
    w, h = SIZES[size]
    slide_text(s, 1.6, 0.95, 30.0, 0.7,
               [[(f"{no}. ", RED), (title + (" (개념도)" if concept else ""), BLK)]], 15)
    slide_text(s, 1.6, 1.75, 30.0, 0.55, f"{size}도표 {w:.2f} × {h:.2f} cm   |   출처: {source}", 10, MG)
    g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(GX), E(GY), E(w), E(h))
    strip_style(g); g.name = "GUIDE"; g.fill.background()
    g.line.color.rgb = RGBColor.from_string(LG); g.line.width = Emu(6350)
    etree.SubElement(g.line._get_or_add_ln(), qn("a:prstDash")).set("val", "dash")
    if note_lines:
        slide_text(s, 1.6, GY + h + 0.35, 30.0, 1.2,
                   ["주: " + note_lines[0]] + ["     " + t for t in note_lines[1:]], 9, MG)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    cv = Canvas(s, GX, GY, no.replace(" ", ""))
    cv.w, cv.h, cv.no = w, h, no
    return cv


def intro_slide(prs, report, charts, checks=()):
    """첫 장 안내: 규격·서식·워드 삽입 방법·도표 목록·발간 전 확인"""
    s = prs.slides.add_slide(_blank(prs))
    slide_text(s, 1.6, 0.95, 30, 0.8, f"{report} 도표 모음 — 안내", 18)
    body = ["규격: 기본 9.25 × 5.10 cm / 대형 18.50 × 7.30 cm — 정사이즈로 제작(워드에서 크기 조정 금지)",
            "서식: 도표 안 글자는 전부 8pt(AppleSDGothicNeoB00, 595959), 강조만 C00000, 표 숫자 Arial",
            "워드 삽입: 차트 = 그룹 복사 → Ctrl+V / 도식·표·사진 도표 = Ctrl+Alt+V → 그림(확장 메타파일) 또는 SVG",
            "원자료: 같은 이름의 엑셀 시트(도표 번호) — 숫자·표는 칸 대조, 개념도·도식은 적힌 날짜·이름 대조('대조' 시트)", "",
            "도표 목록"] + [f"  {c}" for c in charts]
    if checks:
        body += ["", "발간 전 확인"] + [f"  · {c}" for c in checks]
    slide_text(s, 1.6, 2.2, 30, 15, body, 11, DG)
    return s


# ───────────── 도표 안 그리기(Canvas): 좌표 cm, 가이드 좌상단 기준. 글자는 전부 8pt ─────────────
class Canvas:
    def __init__(self, slide, fx, fy, name):
        self.slide, self.fx, self.fy = slide, fx, fy
        self.grp = slide.shapes.add_group_shape()
        self.grp.name = name
        self.sh = self.grp.shapes
        self.w, self.h = SIZES["기본"]

    def X(self, x):
        return E(self.fx + x)

    def Y(self, y):
        return E(self.fy + y)

    def text(self, x, y, w, h, paras, size=CT, color=DG, align="l", anchor="t",
             wrap=True, spacing=1.0, font=B00, bold=False, name=None):
        tb = self.sh.add_textbox(self.X(x), self.Y(y), E(w), E(h))
        bp = tb.text_frame._txBody.find(qn("a:bodyPr"))
        bp.set("wrap", "square" if wrap else "none")
        for k in ("lIns", "rIns", "tIns", "bIns"):
            bp.set(k, "0")
        bp.set("anchor", {"t": "t", "m": "ctr", "b": "b"}[anchor])
        for el in list(bp):
            bp.remove(el)
        etree.SubElement(bp, qn("a:noAutofit"))
        _fill_runs(tb.text_frame, paras, size, color, align, spacing, font, bold)
        if name:
            tb.name = name
        return tb

    def label(self, x, y, s, color=DG, size=CT, align="l", bold=False):
        """한 줄 글자: (x, y) = 글자 왼쪽(가운데·오른쪽) 위. 폭은 tw()로 자동"""
        w = tw(s, size) + 0.1
        x0 = {"l": x, "c": x - w / 2, "r": x - w}[align]
        return self.text(x0, y, w, size * 0.0353 * 1.25, s, size, color, align, wrap=False, bold=bold)

    def box(self, x, y, w, h, fill=None, line=None, lw=0.75, dash=False, paras=None, size=CT,
            color=BLK, align="c", anchor="m", margin=0.05, shape="round", radius=0.1,
            spacing=1.0, font=B00, bold=False, alpha=None, name=None):
        kind = {"round": MSO_SHAPE.ROUNDED_RECTANGLE, "rect": MSO_SHAPE.RECTANGLE,
                "diamond": MSO_SHAPE.DIAMOND, "tri": MSO_SHAPE.ISOSCELES_TRIANGLE,
                "oval": MSO_SHAPE.OVAL, "chevron": MSO_SHAPE.CHEVRON,
                "pentagon": MSO_SHAPE.PENTAGON, "down": MSO_SHAPE.DOWN_ARROW}[shape]
        s = self.sh.add_shape(kind, self.X(x), self.Y(y), E(w), E(h))
        strip_style(s)
        if shape == "round":
            s.adjustments[0] = min(0.5, radius / max(0.01, min(w, h)))
        if fill:
            s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
            if alpha is not None:
                clr = s.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
                etree.SubElement(clr, qn("a:alpha")).set("val", str(int(alpha * 100000)))
        else:
            s.fill.background()
        if line:
            s.line.color.rgb = RGBColor.from_string(line)
            s.line.width = Emu(int(lw * 12700))
            if dash:
                etree.SubElement(s.line._get_or_add_ln(), qn("a:prstDash")).set("val", "sysDash")
        else:
            s.line.fill.background()
        bp = s.text_frame._txBody.find(qn("a:bodyPr"))
        bp.set("wrap", "square")
        for k in ("lIns", "rIns"):
            bp.set(k, str(int(margin * CM)))
        for k in ("tIns", "bIns"):
            bp.set(k, "0")
        bp.set("anchor", {"t": "t", "m": "ctr", "b": "b"}[anchor])
        for el in list(bp):
            bp.remove(el)
        etree.SubElement(bp, qn("a:noAutofit"))
        if paras is not None:
            _fill_runs(s.text_frame, paras, size, color, align, spacing, font, bold)
        if name:
            s.name = name
        return s

    def line(self, x1, y1, x2, y2, color=DG, lw=0.75, dash=False, arrow=False, name=None):
        c = self.sh.add_connector(MSO_CONNECTOR.STRAIGHT, self.X(x1), self.Y(y1), self.X(x2), self.Y(y2))
        strip_style(c)
        c.line.color.rgb = RGBColor.from_string(color)
        c.line.width = Emu(int(lw * 12700))
        ln = c.line._get_or_add_ln()
        if dash:
            etree.SubElement(ln, qn("a:prstDash")).set("val", "sysDash" if dash is True else dash)
        if arrow:
            te = etree.SubElement(ln, qn("a:tailEnd"))
            te.set("type", "triangle"); te.set("w", "med"); te.set("len", "med")
        if name:
            c.name = name
        return c

    def arrow(self, x1, y1, x2, y2, color=DG, lw=1.0, dash=False):
        """차트 안 화살표 1pt(CAGR·추세) / 도식의 흐름 화살표는 lw=1.5"""
        return self.line(x1, y1, x2, y2, color, lw, dash, arrow=True)

    def back(self, shape):
        """그룹 맨 뒤로(차트 뒤에 깔기)"""
        el = shape._element
        self.grp._element.remove(el)
        self.grp._element.insert(2, el)            # nvGrpSpPr, grpSpPr 다음
        return shape

    def band(self, x0, x1, y0, y1, fill=PINK, alpha=0.45, back=True):
        """강조 구간 음영(평균 구간·추정 구간·이벤트 기간). 기본으로 차트 뒤에 깐다"""
        s = self.box(x0, y0, x1 - x0, y1 - y0, fill=fill, shape="rect", alpha=alpha)
        return self.back(s) if back else s

    def hline(self, y, x0, x1, color=RED, lw=1.0, dash=True, text=None, text_pos="r"):
        """평균선·목표선 + 라벨(예: 'Avg. 5.32x')"""
        self.line(x0, y, x1, y, color, lw, dash)
        if text:
            if text_pos == "r":
                self.label(x1, y - 0.36, text, color, align="r")
            else:
                self.label(x0, y - 0.36, text, color)

    def num(self, n, x, y, color=DG, size=CT):
        """주가 차트 번호 주석 ①②③ — 본문에 같은 번호로 설명 문단"""
        return self.label(x, y, circled(n), color, size, align="c")

    def poly(self, pts, fill=None, line=None, lw=0.75, dash=False, closed=True, name=None):
        x0, y0 = pts[0]
        fb = self.sh.build_freeform(self.X(x0), self.Y(y0), scale=1.0)
        fb.add_line_segments([(self.X(x), self.Y(y)) for x, y in pts[1:]], close=closed)
        s = fb.convert_to_shape()
        strip_style(s)
        if fill:
            s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
        else:
            s.fill.background()
        if line:
            s.line.color.rgb = RGBColor.from_string(line)
            s.line.width = Emu(int(lw * 12700))
            if dash:
                etree.SubElement(s.line._get_or_add_ln(), qn("a:prstDash")).set("val", "sysDash")
        else:
            s.line.fill.background()
        if name:
            s.name = name
        return s

    def photo(self, x, y, w, h, path=None, what="", where="", fit="contain"):
        """사진 도표. path가 있으면 비율 유지해 (x,y,w,h) 안에 넣는다(contain=다 보이게 / cover=꽉 채우고 자름).
        없으면 점선 자리 + '사진: 무엇 / 받을 곳' — 사용자가 넣는다(사진 제안 표에도 같은 내용)."""
        if path is None:
            return self.box(x, y, w, h, line=MG, dash=True, shape="rect",
                            paras=[f"사진: {what}", f"받을 곳: {where}"] if where else [f"사진: {what}"],
                            color=MG, name="PHOTO_SLOT")
        from PIL import Image
        iw, ih = Image.open(path).size
        r = iw / ih
        if fit == "contain":
            pw, ph = (w, w / r) if w / r <= h else (h * r, h)
            pic = self.sh.add_picture(path, self.X(x + (w - pw) / 2), self.Y(y + (h - ph) / 2), E(pw), E(ph))
        else:
            pic = self.sh.add_picture(path, self.X(x), self.Y(y), E(w), E(h))
            box = w / h
            if box > r:                       # 상자가 더 넓음 → 위아래 자름
                pic.crop_top = pic.crop_bottom = (1 - r / box) / 2
            else:                             # 상자가 더 높음 → 좌우 자름
                pic.crop_left = pic.crop_right = (1 - box / r) / 2
        pic.name = "PHOTO"
        return pic

    def val(self, x, y, sheet, col, i, pre="", post="", fmt=None, color=DG, align="l", bold=False, wrap_at=None):
        """엑셀 값으로 만든 주석 글자(예: '1년 만에 8.1배'). 이름 VAL:시트|열|i → 대조 때 엑셀 값과 비교"""
        v = sheet.cols[col][i] if col in sheet.cols else None
        t = fmt_value(v, fmt or sheet.fmt(col)) if col in sheet.cols else sheet.texts[col][i]
        s_ = pre + t + post
        if wrap_at:
            s_ = s_.replace(wrap_at, "\n", 1)
        if "\n" in s_:                              # 여러 줄(예: '26말\n첫 출하')
            w = max(tw(l) for l in s_.split("\n")) + 0.1
            x0 = {"l": x, "c": x - w / 2, "r": x - w}[align]
            tb = self.text(x0, y, w, 0.36 * (s_.count("\n") + 1), s_, CT, color, align, wrap=False, bold=bold)
        else:
            tb = self.label(x, y, s_, color, align=align, bold=bold)
        tb.name = f"VAL:{sheet.name}|{col}|{i}|{fmt or ''}"
        return tb

    def slot(self, x, y, w, h, *lines):
        """아직 못 만든 자리(원자료 없는 차트 등) — 점선 상자 + 무엇이 필요한지"""
        return self.box(x, y, w, h, line=MG, dash=True, shape="rect", paras=list(lines), color=MG, name="SLOT")

    def check(self, x, y, d=0.32, color=RED):
        """체크 표시(도식의 충족 항목) — 글자 대신 도형"""
        return self.poly([(x, y + d * 0.5), (x + d * 0.38, y + d * 0.88), (x + d, y + d * 0.1)],
                         line=color, lw=2.0, closed=False)

    def diamond(self, cx, cy, d, fill=RED, line=None):
        return self.box(cx - d / 2, cy - d / 2, d, d, fill=fill, line=line, shape="diamond")

    def tri(self, cx, cy, d, fill=MG, line=None):
        return self.box(cx - d / 2, cy - d / 2, d, d * 0.9, fill=fill, line=line, shape="tri")

    def timeline(self, x0, x1, y, events, color=RED, above=None):
        """events: [(라벨(날짜), 설명, 강조여부)] — 빨강 원 노드 + 회색 축. 위·아래 번갈아 설명"""
        self.line(x0, y, x1, y, MG, 1.5)
        n = len(events)
        for i, (d, txt, hot) in enumerate(events):
            cx = x0 + (x1 - x0) * (i + 0.5) / n
            self.box(cx - 0.13, y - 0.13, 0.26, 0.26, fill=WHT, line=color, lw=2.25, shape="oval")
            up = (i % 2 == 0) if above is None else above
            c = RED if hot else DG
            dw = tw(d) + 0.1
            self.label(min(max(dw / 2, cx), self.w - dw / 2), y - 0.55 if up else y + 0.25, d, c, align="c", bold=hot)
            if txt:
                ty = y - 0.55 - 0.32 * (txt.count("\n") + 1) if up else y + 0.55
                tw_ = max(tw(t_) for t_ in txt.split("\n")) + 0.2
                tx = min(max(0.0, cx - tw_ / 2), self.w - tw_)
                self.text(tx, ty, tw_, 0.32 * (txt.count("\n") + 1), txt, CT, c, "c")

    def table_from_sheet(self, x, y, colw, sheet, cols, unit="", first="구분", total_rows=(), heads=None, **kw):
        """엑셀 시트에서 바로 표를 만든다(숫자는 시트 표시 형식대로). 대조 검사 대상(이름 TABLE:시트|열…)
        colw=None이면 글자 폭으로 자동(남는 폭은 고르게). heads: 머리행 표시 글자(줄바꿈 등) — 없으면 엑셀 열 이름"""
        head = [f"{first}{' ' + unit if unit else ''}"] + list(heads or cols)
        body = []
        for i, c in enumerate(sheet.cats):
            row = [c]
            for n in cols:
                row.append(fmt_value(sheet.cols[n][i], sheet.fmt(n)) if n in sheet.cols else sheet.texts[n][i])
            body.append(row)
        emph = set(kw.pop("emph", ())) | {(r + 1, j) for r in total_rows for j in range(len(head))}
        gf = self.table(x, y, colw, [head] + body, emph=emph, **kw)
        gf.name = "TABLE:" + sheet.name + "|" + "|".join(cols)
        return gf

    def table(self, x, y, colw, rows, rowh=0.45, emph=(), red=(), size=CT, last_thick=True,
              head_fill=DG, first_col_left=False, gray=()):
        """SMIC 도표 표: 머리행 595959+흰 B00 / 본문 M00 검정(숫자는 Arial) / 가운데 / 세로줄 없음 /
        마지막 줄 1.5pt / emph = {(행,열)} 노랑 FFF2CC 칸 / red = {(행,열)} 빨강 글자. 단위는 '구분' 옆 괄호"""
        nr, nc = len(rows), len(rows[0])
        if colw is None:                             # 자동 폭: 가장 긴 줄 + 여백, 남는 폭은 고르게
            need = [max(tw(l) for r in rows for l in str(r[j]).split("\n")) + 0.3 for j in range(nc)]
            extra = (self.w - x) - sum(need)
            colw = [n + extra / nc for n in need]
        gf = self.slide.shapes.add_table(nr, nc, self.X(x), self.Y(y), E(sum(colw)), E(rowh * nr))
        self.grp._element.append(gf._element)
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        for a in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
            tblPr.attrib.pop(a, None)
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is not None:
            tblPr.remove(sid)
        for j, w in enumerate(colw):
            tbl.columns[j].width = E(w)
        for i in range(nr):
            tbl.rows[i].height = E(rowh)
        isnum = re.compile(r"^[\s\(\)\-+$₩%x.,0-9]+[A-Za-z%x]*$")
        for i in range(nr):
            for j in range(nc):
                cell = tbl.cell(i, j)
                txt = str(rows[i][j])
                for k, ln_ in enumerate(txt.split("\n")):
                    p = cell.text_frame.paragraphs[0] if k == 0 else cell.text_frame.add_paragraph()
                    p._p.get_or_add_pPr().set("algn", "l" if (first_col_left and j == 0 and i > 0) else "ctr")
                    r = p.add_run(); r.text = ln_
                    if i == 0:
                        f, col = B00, WHT
                    else:
                        f = ARIAL if isnum.match(ln_ or "x") and ln_ not in ("○", "△", "×") else M00
                        col = RED if (i, j) in red else (MG if (i, j) in gray else BLK)
                    font_xml(r._r.get_or_add_rPr(), size, col, f)
                tcPr = cell._tc.get_or_add_tcPr()
                for el in list(tcPr):
                    tcPr.remove(el)
                tcPr.set("marL", str(int(0.03 * CM))); tcPr.set("marR", str(int(0.03 * CM)))
                tcPr.set("marT", "0"); tcPr.set("marB", "0"); tcPr.set("anchor", "ctr")

                def ln(tag, w=None, col=DG, cmpd=None):
                    el = etree.SubElement(tcPr, qn(tag))
                    if w is None:
                        el.set("w", "0"); etree.SubElement(el, qn("a:noFill"))
                    else:
                        el.set("w", str(int(w * 12700)))
                        sf = etree.SubElement(el, qn("a:solidFill"))
                        etree.SubElement(sf, qn("a:srgbClr")).set("val", col)
                ln("a:lnL"); ln("a:lnR")
                ln("a:lnT") if i <= 1 else ln("a:lnT", 0.5)
                if i == nr - 1:
                    ln("a:lnB", 1.5 if last_thick else 0.5)
                elif i == 0:
                    ln("a:lnB")
                else:
                    ln("a:lnB", 0.5)
                fill = head_fill if i == 0 else (YEL if (i, j) in emph else None)
                if fill:
                    sf = etree.SubElement(tcPr, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr")).set("val", fill)
                else:
                    etree.SubElement(tcPr, qn("a:noFill"))
        self.sh._recalculate_extents()
        return gf


# ───────────── 네이티브 차트 ─────────────
_KIND = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "barh": XL_CHART_TYPE.BAR_CLUSTERED,
         "stack": XL_CHART_TYPE.COLUMN_STACKED, "stack100": XL_CHART_TYPE.COLUMN_STACKED_100,
         "line": XL_CHART_TYPE.LINE_MARKERS, "area": XL_CHART_TYPE.AREA_STACKED,
         "pie": XL_CHART_TYPE.PIE}


def _plot_area(ch):
    return ch._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))


def style_axis(ax, fmt=None, mn=None, mx=None, major=None, delete=False):
    """축선 595959 0.75pt + 주눈금 바깥쪽, 눈금선 없음, 글자 B00 8pt 595959"""
    for g in ax.findall(qn("c:majorGridlines")) + ax.findall(qn("c:minorGridlines")):
        ax.remove(g)
    for tag in ("c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr"):
        for e in ax.findall(qn(tag)):
            ax.remove(e)
    d = ax.find(qn("c:delete"))
    if d is not None:
        d.set("val", "1" if delete else "0")
    anchor = ax.find(qn("c:numFmt"))
    if anchor is None:
        anchor = ax.find(qn("c:title")) if ax.find(qn("c:title")) is not None else ax.find(qn("c:axPos"))
    for it in (_el("c:majorTickMark", val="out"), _el("c:minorTickMark", val="none"),
               _el("c:tickLblPos", val="nextTo"), _ln(0.75), _txpr()):
        anchor.addnext(it); anchor = it
    if fmt:
        nf = ax.find(qn("c:numFmt"))
        if nf is None:
            nf = _el("c:numFmt"); ax.find(qn("c:axPos")).addnext(nf)
        nf.set("formatCode", fmt); nf.set("sourceLinked", "0")
    sc = ax.find(qn("c:scaling"))
    for tag, v in (("c:max", mx), ("c:min", mn)):       # 스키마 순서: orientation, max, min
        if v is not None:
            e = sc.find(qn(tag))
            if e is None:
                e = _el(tag); sc.append(e)
            e.set("val", repr(float(v)))
    mxe, mne = sc.find(qn("c:max")), sc.find(qn("c:min"))
    if mxe is not None and mne is not None:
        sc.remove(mxe); mne.addprevious(mxe)
    if major is not None and ax.tag == qn("c:valAx"):
        mu = ax.find(qn("c:majorUnit"))
        if mu is None:
            mu = _el("c:majorUnit")
            cb = ax.find(qn("c:crossBetween"))
            (cb if cb is not None else ax[-1]).addnext(mu)
        mu.set("val", repr(float(major)))


def _ser_name(ser):
    v = ser.find(qn("c:tx") + "//" + qn("c:v"))
    return v.text if v is not None else ""


def _series_xml(ch):
    return _plot_area(ch).findall(".//" + qn("c:ser"))


def series_color(ser):
    for path in ("c:spPr/a:solidFill/a:srgbClr", "c:spPr/a:ln/a:solidFill/a:srgbClr"):
        e = ser.find(path.replace("c:", "{%s}" % C).replace("a:", "{%s}" % A))
        if e is not None:
            return e.get("val")
    return DG


def _color_series(ch, kind, colors):
    for i, s in enumerate(ch.plots[0].series if kind != "pie" else []):
        col = colors[i % len(colors)]
        if kind == "line":
            style_line(s, col)
        else:
            s.format.fill.solid(); s.format.fill.fore_color.rgb = RGBColor.from_string(col)
            s.format.line.fill.background()
            if kind in ("bar", "barh", "stack", "stack100"):
                s.invert_if_negative = False


def style_line(series, color=DG, width=1.5, marker=True):
    """꺾은선: 1.5pt, 원형 표식 3, 표식 채우기·테두리 = 선 색(테두리 0.75pt — 1.5pt는 발표 장표 규칙)"""
    series.format.line.color.rgb = RGBColor.from_string(color)
    series.format.line.width = Pt(width)
    series.smooth = False
    if marker:
        series.marker.style = XL_MARKER_STYLE.CIRCLE
        series.marker.size = 3
        series.marker.format.fill.solid()
        series.marker.format.fill.fore_color.rgb = RGBColor.from_string(color)
        series.marker.format.line.color.rgb = RGBColor.from_string(color)
        series.marker.format.line.width = Pt(0.75)
    else:
        series.marker.style = XL_MARKER_STYLE.NONE


def add_chart(cv, kind, cats, series, sheet, fmt="#,##0", x=0.0, y=0.0, w=None, h=None,
              colors=None, emph=None, yscale="auto", multilevel=False, gap=60, overlap=None,
              markers=None, on_tick=None, legend=True, units=("", ""), inner=None, skip=None, head=0.0,
              min_pct=0.03, zero=None, axes=True):
    """네이티브 차트를 도표 그룹 안에 넣는다.
    sheet: 엑셀 시트 이름(=도표 번호). 대조할 때 이 이름으로 엑셀을 찾는다 — 반드시 넣기.
    series: [(이름, 값)] — sheet.series(...)로 엑셀에서 꺼낸 것. 이름 = 엑셀 머리행.
    kind: bar | barh | stack | stack100 | line | area | pie
    emph: {점 인덱스: 색} — 단일 계열 강조(나머지 colors[0]).  colors: 계열 색 순서(기본 PAL)
    yscale: 'auto'(nice 눈금 고정 → 주석 좌표 정확) | (min, max, major) | None(엑셀 자동)
    head: 위쪽 여유 비율(합계 레이블·화살표 자리) / min_pct: 원그래프에서 이보다 작은 조각은 숫자 생략
    zero: 축을 0부터(None = 막대는 0부터, 꺾은선은 데이터 범위) / axes=False: 축 숨김(작은 삽입 차트)
    units: (왼쪽 축 단위, 오른쪽 축 단위) 예: ('(십억 원)', '(%)')
    반환 Plot — .gf(그래픽 프레임) .x(i) .y(v) 로 주석 좌표(cm, 가이드 기준)"""
    w = cv.w if w is None else w
    h = cv.h if h is None else h
    cd = CategoryChartData(number_format=fmt)
    if multilevel:                                   # '2024 1Q' → 2단 축(연도/분기)
        groups = {}
        for c in cats:
            top, sub = str(c).split(" ", 1)
            groups.setdefault(top, []).append(sub)
        for top, subs in groups.items():
            g = cd.add_category(top)
            for s in subs:
                g.add_sub_category(s)
    else:
        cd.categories = list(cats)
    for name, vals in series:
        cd.add_series(name, list(vals))
    gf = cv.slide.shapes.add_chart(_KIND[kind], cv.X(x), cv.Y(y), E(w), E(h), cd)
    gf.name = "CHART:" + sheet
    ch = gf.chart
    _smic_style(ch, kind, legend and (len(series) > 1 or kind == "pie"))
    colors = colors or ([RED] if (len(series) == 1 and kind != "pie") else PAL)
    _color_series(ch, kind, colors)
    if kind == "pie":
        _pie_style(ch, series[0][1], min_pct)
    else:
        pl = ch.plots[0]
        if kind in ("bar", "barh", "stack", "stack100"):
            pl.gap_width = gap
            pl.overlap = 100 if kind.startswith("stack") else (overlap or 0)
        if emph:
            color_points(pl.series[0], colors[0], emph)
        n = len(cats)
        if kind == "line":
            mk = (n <= 24) if markers is None else markers
            for s in pl.series:
                style_line(s, series_color(s._element), marker=mk)
        if (on_tick if on_tick is not None else (kind in ("line", "area") and n > 24)):
            for cb in _plot_area(ch).findall(qn("c:valAx") + "/" + qn("c:crossBetween")):
                cb.set("val", "midCat")
        cat_ax = _plot_area(ch).find(qn("c:catAx"))
        if any(str(c) == "" for c in cats):            # '축 레이블'(월초만 표기) → 모두 표시, 눈금 없음
            skip = 1
            cat_ax.find(qn("c:majorTickMark")).set("val", "none")
        if kind == "barh":                              # 첫 항목이 위로
            cat_ax.find(qn("c:scaling")).find(qn("c:orientation")).set("val", "maxMin")
            vax = _plot_area(ch).find(qn("c:valAx"))
            cr = vax.find(qn("c:crosses"))
            if cr is not None:
                cr.set("val", "max")
        if skip is None and n > 12 and not multilevel:
            lw_ = max(tw(c) for c in cats) + 0.25
            skip = max(1, math.ceil(n * lw_ / (w * 0.85)))
        if skip and skip > 1 and cat_ax is not None:
            for tag, v in (("c:tickLblSkip", skip), ("c:tickMarkSkip", skip)):
                e = cat_ax.find(qn(tag))
                if e is None:
                    e = _insert_before(cat_ax, _el(tag), ["c:noMultiLvlLbl", "c:extLst"])
                e.set("val", str(v))
        va = _plot_area(ch).find(qn("c:valAx"))
        allv = [v for _, vals in series for v in vals if v is not None]
        if kind == "stack":
            allv = [sum(v for v in col if v is not None) for col in zip(*[s[1] for s in series])] + allv
        auto = yscale == "auto"
        if auto and kind != "stack100":
            z = (kind != "line" or min(allv) <= 0) if zero is None else zero
            whole = "." not in fmt.split(";")[0] and "%" not in fmt   # 정수 표시면 2.5 간격 금지(3, 5, 8 같은 눈금 방지)
            yscale = nice_scale(allv, zero=z, head=head, steps=(1, 2, 5, 10) if whole else (1, 2, 2.5, 5, 10))
        if isinstance(yscale, tuple):
            style_axis(va, fmt=axis_fmt(fmt, yscale[2]), mn=yscale[0], mx=yscale[1], major=yscale[2])
        else:
            style_axis(va, fmt=fmt)
            yscale = None
    if not axes and kind != "pie":
        for ax in _plot_area(ch).findall(qn("c:valAx")):
            ax.find(qn("c:delete")).set("val", "1")
        units = ("", "")
    p = Plot(cv, gf, kind, x, y, w, h, len(cats), len(series), gap, yscale, fmt, multilevel)
    p.axes = axes
    p.auto_y, p.head, p.zero = (kind != "pie" and auto), head, zero
    if units[0] and kind != "pie":
        axis_unit(_plot_area(ch).find(qn("c:valAx")), units[0], 0.0, 0.0)
    p.units = units
    p.layout(inner)
    cv.grp._element.append(gf._element)               # 그룹 안으로(워드에 그룹째 복사)
    cv.sh._recalculate_extents()
    return p


def _smic_style(ch, kind, legend):
    """3팀 양식: 배경·윤곽선 없음, 글자 B00 8pt 595959, 눈금선 없음, 범례 상단"""
    cs = ch._chartSpace
    rc = cs.find(qn("c:roundedCorners"))
    if rc is None:
        rc = _el("c:roundedCorners", val=0)
        _insert_before(cs, rc, ["c:AlternateContent", "c:clrMapOvr", "c:pivotSource", "c:protection", "c:chart"])
    rc.set("val", "0")
    for old in cs.findall(qn("c:spPr")) + cs.findall(qn("c:txPr")):
        cs.remove(old)
    sp = etree.fromstring(f'<c:spPr xmlns:c="{C}" xmlns:a="{A}"><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')
    cs.find(qn("c:chart")).addnext(sp)
    sp.addnext(_txpr())
    ch.has_title = False
    atd = cs.find(qn("c:chart")).find(qn("c:autoTitleDeleted"))
    if atd is not None:
        atd.set("val", "1")
    for ax in _plot_area(ch).findall(qn("c:catAx")) + _plot_area(ch).findall(qn("c:valAx")):
        style_axis(ax)
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        lg = cs.find(qn("c:chart")).find(qn("c:legend"))
        for t in lg.findall(qn("c:txPr")):
            lg.remove(t)
        _insert_before(lg, _txpr(), ["c:extLst"])


def _pie_style(ch, vals, min_pct=0.03):
    """원그래프: 엑셀 순서(비중 큰 순, '기타'는 마지막)대로 빨강 진→연, 회색 연→진 /
    숫자 흰색·안쪽·B00 8pt, min_pct 미만 조각은 숫자 생략"""
    pl = ch.plots[0]
    pl.vary_by_categories = True
    s = pl.series[0]
    for i in range(len(vals)):
        pt = s.points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(PAL[i % len(PAL)])
        pt.format.line.color.rgb = RGBColor.from_string(WHT); pt.format.line.width = Pt(0.75)
    dl = s.data_labels
    dl.show_percentage, dl.show_value = True, False
    dl.number_format, dl.number_format_is_linked = "0%", False
    from pptx.enum.chart import XL_LABEL_POSITION
    dl.position = XL_LABEL_POSITION.INSIDE_END
    dls = s._element.find(qn("c:dLbls"))
    for t in dls.findall(qn("c:txPr")):
        dls.remove(t)
    _insert_before(dls, _txpr(CT, WHT), ["c:dLblPos", "c:showLegendKey"])
    tot = sum(v or 0 for v in vals)
    for i, v in enumerate(vals):
        if tot and (v or 0) / tot < min_pct:
            d = _el("c:dLbl"); d.append(_el("c:idx", val=i)); d.append(_el("c:delete", val=1))
            dls.insert(0, d)


def color_points(series, base, per_point=None):
    """강조할 막대만 SMIC 빨강, 나머지 회색 같은 식: color_points(s, LG, {4: RED})"""
    series.format.fill.solid(); series.format.fill.fore_color.rgb = RGBColor.from_string(base)
    for i, col in (per_point or {}).items():
        pt = series.points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(col)


def axis_unit(ax, text, x, y):
    """단위 = 가로 축제목을 축 상단으로 옮긴 것. x, y는 차트 안 비율(0~1)"""
    for old in ax.findall(qn("c:title")):
        ax.remove(old)
    t = etree.fromstring(
        f'<c:title xmlns:c="{C}" xmlns:a="{A}"><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/>'
        f'<a:p><a:pPr><a:defRPr sz="800" b="0"><a:solidFill><a:srgbClr val="{DG}"/></a:solidFill>'
        f'<a:latin typeface="{B00}"/><a:ea typeface="{B00}"/></a:defRPr></a:pPr>'
        f'<a:r><a:rPr lang="ko-KR" sz="800" b="0"><a:solidFill><a:srgbClr val="{DG}"/></a:solidFill>'
        f'<a:latin typeface="{B00}"/><a:ea typeface="{B00}"/></a:rPr><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx>'
        f'<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="{x}"/><c:y val="{y}"/>'
        f'</c:manualLayout></c:layout><c:overlay val="0"/></c:title>')
    ax.find(qn("c:axPos")).addnext(t)


def plot_inner(ch, x, y, w, h):
    """플롯 영역(눈금 레이블 제외) 위치를 차트 안 비율로 고정"""
    pa = _plot_area(ch)
    lay = pa.find(qn("c:layout"))
    if lay is None:
        lay = _el("c:layout"); pa.insert(0, lay)
    for e in list(lay):
        lay.remove(e)
    ml = etree.SubElement(lay, qn("c:manualLayout"))
    for tag, v in (("c:layoutTarget", "inner"), ("c:xMode", "edge"), ("c:yMode", "edge"),
                   ("c:x", x), ("c:y", y), ("c:w", w), ("c:h", h)):
        etree.SubElement(ml, qn(tag)).set("val", v if isinstance(v, str) else str(round(v, 4)))


class Plot:
    """add_chart 반환값. 플롯 영역을 cm 단위로 고정해 두고 주석 좌표를 계산한다."""

    def __init__(self, cv, gf, kind, x, y, w, h, n, nser, gap, yscale, fmt, multilevel):
        self.cv, self.gf, self.ch, self.kind = cv, gf, gf.chart, kind
        self.cx, self.cy, self.w, self.h, self.n, self.nser = x, y, w, h, n, nser
        self.gap, self.yscale, self.fmt, self.ml = gap, yscale, fmt, multilevel
        self.y2scale, self.fmt2, self.units = None, None, ("", "")

    def layout(self, inner=None):
        """inner=(왼, 위, 오른, 아래 여백 cm). 없으면 눈금 레이블 폭·범례·단위로 자동"""
        self.inner_arg = inner
        if self.kind == "pie":
            return
        if inner is None:                       # 플롯 영역 최대: 눈금 글자·단위·범례가 쓰는 만큼만 비운다
            def lab_w(sc, fmt):
                if not sc:
                    return 1.0
                k = int(round((sc[1] - sc[0]) / sc[2]))
                return max(tw(fmt_value(sc[0] + i * sc[2], fmt)) for i in range(k + 1)) + 0.18
            if self.kind == "barh":
                names = [str(c) for c in self.ch.plots[0].categories]
                left = max(tw(n) for n in names) + 0.2
            else:
                left = lab_w(self.yscale, self.fmt) if getattr(self, "axes", True) else 0.05
            right = lab_w(self.y2scale, self.fmt2) if self.y2scale else 0.12
            top = 0.55 if (self.ch.has_legend or any(self.units)) else 0.12   # 단위·범례는 한 줄에 같이
            if self.ch.has_legend:
                names = [_ser_name(e) for e in _series_xml(self.ch)]
                lw_ = sum(tw(n) + (0.75 if self.kind == "line" else 0.45) for n in names)
                room = self.w - tw(self.units[0] or "") - tw(self.units[1] or "") - 0.4
                if lw_ > room:
                    top += 0.33 * math.ceil(lw_ / (self.w - 0.4) - 1 + 1e-9)
            bottom = 0.38 + (0.33 if self.ml else 0)
            if self.kind == "barh":
                bottom = 0.38 if getattr(self, "axes", True) else 0.05
            inner = (left, top, right, bottom)
        l, t, r, b = inner
        self.inner = inner
        self.px0, self.py0 = self.cx + l, self.cy + t
        self.pw, self.ph = self.w - l - r, self.h - t - b
        plot_inner(self.ch, l / self.w, t / self.h, self.pw / self.w, self.ph / self.h)
        if self.units[1]:
            ax2 = _plot_area(self.ch).findall(qn("c:valAx"))[-1]
            axis_unit(ax2, self.units[1], max(0.0, 1 - (tw(self.units[1]) + 0.05) / self.w), 0.0)

    def x(self, i, s=None):
        """i번째 기간의 가로 중심(cm). 묶은 막대에서 s번째 계열 막대 중심은 s 지정"""
        on_tick = any(cb.get("val") == "midCat" for cb in
                      _plot_area(self.ch).findall(qn("c:valAx") + "/" + qn("c:crossBetween")))
        if on_tick:
            return self.px0 + self.pw * i / max(1, self.n - 1)
        slot = self.pw / self.n
        cx = self.px0 + slot * (i + 0.5)
        if s is not None and self.kind == "bar" and self.nser > 1:
            grp = slot / (1 + self.gap / 100)
            bw = grp / self.nser
            cx = self.px0 + slot * i + (slot - grp) / 2 + bw * (s + 0.5)
        return cx

    def y(self, v, rhs=False):
        sc = self.y2scale if rhs else self.yscale
        if not sc:
            raise ValueError("yscale을 고정해야 주석 좌표를 계산할 수 있다(add_chart yscale='auto')")
        return self.py0 + self.ph * (sc[1] - v) / (sc[1] - sc[0])

    @property
    def top(self):
        return self.py0

    @property
    def bottom(self):
        return self.py0 + self.ph

    def rhs(self, names, kind="line", fmt="#,##0", yscale="auto", colors=None, unit="", zero=None):
        """이름이 names인 계열을 오른쪽 축(RHS)으로. kind='line'이면 꺾은선(1.5pt·표식 3)"""
        pa = _plot_area(self.ch)
        prim = pa.find(qn("c:barChart")) if pa.find(qn("c:barChart")) is not None else pa.find(qn("c:lineChart"))
        moved, vals = [], []
        for ser in list(prim.findall(qn("c:ser"))):
            if _ser_name(ser) in names:
                prim.remove(ser); moved.append(ser)
                vals += [float(v.text) for v in ser.findall(".//" + qn("c:val") + "//" + qn("c:v"))]
        line = _el("c:lineChart")
        line.append(_el("c:grouping", val="standard")); line.append(_el("c:varyColors", val=0))
        for k, ser in enumerate(moved):
            col = (colors or [DG, MG, PINK2])[k % 3]
            for e in ser.findall(qn("c:invertIfNegative")) + ser.findall(qn("c:spPr")) + ser.findall(qn("c:marker")):
                ser.remove(e)
            sp = etree.fromstring(
                f'<c:spPr xmlns:c="{C}" xmlns:a="{A}"><a:ln w="19050" cap="rnd"><a:solidFill>'
                f'<a:srgbClr val="{col}"/></a:solidFill><a:round/></a:ln></c:spPr>')
            mk = etree.fromstring(
                f'<c:marker xmlns:c="{C}" xmlns:a="{A}"><c:symbol val="circle"/><c:size val="3"/>'
                f'<c:spPr><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill></a:ln></c:spPr></c:marker>')
            ser.find(qn("c:tx")).addnext(sp); sp.addnext(mk)
            if ser.find(qn("c:smooth")) is None:
                ser.append(_el("c:smooth", val=0))
            line.append(ser)
        line.append(_el("c:marker", val=1))
        line.append(_el("c:axId", val=900001)); line.append(_el("c:axId", val=900002))
        prim.addnext(line)
        last_ax = pa.findall(qn("c:valAx"))[-1]
        cat2 = etree.fromstring(
            f'<c:catAx xmlns:c="{C}"><c:axId val="900001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            f'<c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            f'<c:tickLblPos val="nextTo"/><c:crossAx val="900002"/><c:crosses val="autoZero"/>'
            f'<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')
        val2 = etree.fromstring(
            f'<c:valAx xmlns:c="{C}"><c:axId val="900002"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            f'<c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="{fmt}" sourceLinked="0"/>'
            f'<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'<c:crossAx val="900001"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>')
        last_ax.addnext(cat2); cat2.addnext(val2)
        zero2 = (min(vals) <= 0 or kind != "line") if zero is None else zero
        if self.auto_y and self.kind != "stack100":         # 남은 왼쪽 계열로 왼쪽 축 다시
            rest = [float(v.text) for ser in prim.findall(qn("c:ser"))
                    for v in ser.findall(".//" + qn("c:val") + "//" + qn("c:v"))]
            if prim.tag == qn("c:barChart") and prim.find(qn("c:grouping")).get("val") == "stacked":
                n_ = self.n
                cols_ = [[float(v.text) for v in ser.findall(".//" + qn("c:val") + "//" + qn("c:v"))]
                         for ser in prim.findall(qn("c:ser"))]
                rest += [sum(c[i] for c in cols_ if i < len(c)) for i in range(n_)]
            z1 = (self.kind != "line" or min(rest) <= 0) if self.zero is None else self.zero
            self.yscale = nice_scale(rest, zero=z1, head=self.head)
            style_axis(pa.findall(qn("c:valAx"))[0], fmt=axis_fmt(self.fmt, self.yscale[2]),
                       mn=self.yscale[0], mx=self.yscale[1], major=self.yscale[2])
        if yscale == "auto" and self.yscale:                 # 오른쪽 축 눈금 개수 = 왼쪽(같은 높이)
            k1 = round((self.yscale[1] - self.yscale[0]) / self.yscale[2])
            yscale = fit_scale(vals, k1, zero=zero2, head=self.head)
        elif yscale == "auto":
            yscale = nice_scale(vals, zero=zero2, head=self.head)
        self.y2scale, self.fmt2 = (yscale if isinstance(yscale, tuple) else None), fmt
        if self.y2scale:
            style_axis(val2, fmt=axis_fmt(fmt, yscale[2]), mn=yscale[0], mx=yscale[1], major=yscale[2])
        else:
            style_axis(val2, fmt=fmt)
        self.units = (self.units[0], unit or self.units[1])
        self.layout(self.inner_arg)
        return val2

    def labels(self, ser_name, points=None, fmt=None, color=None, pos=None, bold=False):
        """데이터레이블 = 계열 색(기본). points=None이면 전부, [i…]면 그 점만(강조는 color=RED).
        pos: bar 'outEnd'/'inEnd'/'ctr', stack 'ctr'/'inEnd', line 't'/'b'/'r'"""
        ser = next(s for s in _series_xml(self.ch) if _ser_name(s) == ser_name)
        is_line = ser.getparent().tag == qn("c:lineChart")
        pos = pos or ("t" if is_line else ("ctr" if self.kind.startswith("stack") else "outEnd"))
        col = color or series_color(ser)
        on_rhs = is_line and ser.getparent().find(qn("c:axId")).get("val") == "900001"
        f = fmt or (self.fmt2 if on_rhs else self.fmt)
        for old in ser.findall(qn("c:dLbls")):
            ser.remove(old)
        dls = _el("c:dLbls")

        def body(parent, show):
            parent.append(_el("c:numFmt", formatCode=f, sourceLinked=0))
            parent.append(etree.fromstring(
                f'<c:spPr xmlns:c="{C}" xmlns:a="{A}"><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'))
            parent.append(_txpr(CT, col, bold))
            parent.append(_el("c:dLblPos", val=pos))
            for tag, v in (("showLegendKey", 0), ("showVal", show), ("showCatName", 0),
                           ("showSerName", 0), ("showPercent", 0), ("showBubbleSize", 0)):
                parent.append(_el("c:" + tag, val=v))
        if points is not None:
            for i in points:
                d = _el("c:dLbl"); d.append(_el("c:idx", val=i)); body(d, 1); dls.append(d)
        body(dls, 0 if points is not None else 1)
        _insert_before(ser, dls, ["c:trendline", "c:errBars", "c:cat", "c:val"])
        return dls
