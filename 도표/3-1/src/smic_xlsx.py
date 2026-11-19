# smic_xlsx.py — 도표 원자료 엑셀(유일한 원본). 시트 = 도표 하나, 머리행 한 줄 + 데이터.
# 원값은 1의 자리까지 그대로, 단위는 표시 형식으로만. 파생값(비율·YoY·지수)은 엑셀 수식으로 남긴다.
import os, re, shutil, subprocess, tempfile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

B00, M00, ARIAL = "AppleSDGothicNeoB00", "AppleSDGothicNeoM00", "Arial"
DG, YEL = "595959", "FFF2CC"
INDEX = "목록"
INDEX_HEAD = ["도표", "제목", "종류", "크기", "단위", "출처", "주", "대조 결과", "확인 필요"]
KINDS = ("숫자", "표", "개념도", "도식", "사진")   # 숫자·표 = 칸 대조 / 개념도·도식·사진 = 글자(사실) 대조 또는 대조 없음
THIN = Side(style="thin", color=DG)
THICK = Side(style="medium", color=DG)


def _num_cell(c, fmt):
    c.font = Font(name=ARIAL, size=10)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="right", vertical="center")


def _txt_cell(c, bold=False, color="000000", fill=None, align="center"):
    c.font = Font(name=B00 if bold else M00, size=10, color=color)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)


class SourceBook:
    """wb = SourceBook()
    wb.add("도표 2-3", "동사 매출과 마진 추이", periods=["2022", ...],
           cols=[("매출액(LHS)", [4223754000000, ...], "#,##0,,,"),            # 원값 그대로
                 ("OPM(RHS)", "={영업이익}/{매출액(LHS)}", "0.0%"),             # 파생 = 수식
                 ("YoY", "={매출액(LHS)}/{매출액(LHS)@-1}-1", "0.0%"),           # @-1 = 윗행
                 ("영업이익", [...], "#,##0,,,", "hide")],                        # 차트에 안 쓰는 보조 열
           labels=[("이벤트", ["", "PQ 통과", ...])],
           source="LS전선 사업보고서(2022~2025)", row_sources=[...],           # 행별 출처(선택)
           unit="(십억 원) / (%)", size="기본", concept=False, note="")
    wb.save("원자료.xlsx")"""

    def __init__(self):
        self.wb = Workbook()
        ws = self.wb.active
        ws.title = INDEX
        for j, h in enumerate(INDEX_HEAD, 1):
            _txt_cell(ws.cell(1, j, h), bold=True, color="FFFFFF", fill=DG)
        for j, w in enumerate([10, 40, 8, 6, 18, 40, 40, 18, 40], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = "A2"

    def add(self, no, title, periods, cols, labels=(), source="", row_sources=None,
            unit="", size="기본", concept=False, note="", first="기간", kind=None, check=""):
        if len(no) > 31 or re.search(r"[:\\/?*\[\]]", no):
            raise ValueError("시트 이름 불가: " + no)
        ws = self.wb.create_sheet(no)
        names = [c[0] for c in cols]
        letter = {n: get_column_letter(i + 2) for i, n in enumerate(names)}
        heads = [first] + names + [l[0] for l in labels] + ["출처"]   # first: 기간 | 구분(원그래프·표)
        for j, h in enumerate(heads, 1):
            _txt_cell(ws.cell(1, j, h), bold=True, color="FFFFFF", fill=DG)
        n = len(periods)
        for i, p in enumerate(periods):
            r = i + 2
            _txt_cell(ws.cell(r, 1, p))
            for j, col in enumerate(cols):
                name, vals, fmt = col[0], col[1], col[2] if len(col) > 2 else "#,##0"
                if isinstance(vals, str):            # 수식: {열 이름} = 같은 행, {열 이름@-k} = k행 위
                    lag = max([int(k) for k in re.findall(r"@-(\d+)", vals)] or [0])
                    f = re.sub(r"\{([^}@]+)(?:@-(\d+))?\}",
                               lambda m: f"{letter[m.group(1)]}{r - int(m.group(2) or 0)}", vals)
                    v = None if i < lag else f
                else:
                    if len(vals) != n:
                        raise ValueError(f"{no} '{name}': 값 {len(vals)}개 ≠ 기간 {n}개")
                    v = vals[i]
                _num_cell(ws.cell(r, j + 2, v), fmt)
            for k, (lname, lvals) in enumerate(labels):    # 글자 열. '='로 시작하면 수식(예: COUNTIF)
                _txt_cell(ws.cell(r, len(cols) + 2 + k, lvals[i] or None), align="left")
            src = (row_sources[i] if row_sources else source)
            _txt_cell(ws.cell(r, len(heads), src), align="left")
        for j in range(1, len(heads) + 1):
            ws.cell(n + 1, j).border = Border(bottom=THICK)
            ws.column_dimensions[get_column_letter(j)].width = max(10, min(40, len(str(heads[j - 1])) * 2 + 4))
        ws.column_dimensions[get_column_letter(len(heads))].width = 45
        for r in range(1, n + 2):
            ws.row_dimensions[r].height = 16
        ws.freeze_panes = "B2"
        hidden = [c[0] for c in cols if len(c) > 3 and c[3] == "hide"]
        idx = self.wb[INDEX]
        kind = kind or ("개념도" if concept else ("숫자" if any(not isinstance(c[1], str) or str(c[1]).startswith("=")
                                                         for c in cols) else "표"))
        need = "; ".join(x for x in (check, "보조 열(차트 미사용): " + ", ".join(hidden) if hidden else "") if x)
        idx.append([no, title, kind, size, unit, source, note, "", need])
        for j in range(1, len(INDEX_HEAD) + 1):
            _txt_cell(idx.cell(idx.max_row, j), align="left")
        return ws

    def add_facts(self, no, title, kind, facts=(), size="기본", note="", check=""):
        """숫자 없는 도표(개념도·도식·사진). facts = [(도표에 적힌 글자, 내용, 출처)].
        도표 안의 날짜·이름·숫자는 여기 적고, 대조 때 슬라이드 글자에 그대로 있는지 본다.
        facts가 없으면 대조 없음(순수 개념도·사진)."""
        assert kind in ("개념도", "도식", "사진"), kind
        ws = self.wb.create_sheet(no)
        for j, h in enumerate(["표기", "내용", "출처"], 1):
            _txt_cell(ws.cell(1, j, h), bold=True, color="FFFFFF", fill=DG)
        for i, (lab, what, src) in enumerate(facts, 2):
            for j, v in enumerate((lab, what, src), 1):
                _txt_cell(ws.cell(i, j, v), align="left")
        for col, w in zip("ABC", (28, 50, 50)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        idx = self.wb[INDEX]
        idx.append([no, title, kind, size, "", "; ".join(dict.fromkeys(f[2] for f in facts)), note, "", check])
        for j in range(1, len(INDEX_HEAD) + 1):
            _txt_cell(idx.cell(idx.max_row, j), align="left")
        return ws

    def save(self, path):
        self.wb.calculation.fullCalcOnLoad = True
        self.wb.save(path)
        return path


def recalc(path):
    """LibreOffice로 수식을 계산한 사본 경로를 돌려준다(원본은 그대로)."""
    out = tempfile.mkdtemp(prefix="smic_recalc_")
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    p = os.path.join(out, os.path.basename(path))
    if not os.path.exists(p):
        raise RuntimeError("LibreOffice 재계산 실패: " + path)
    return p


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


class Sheet:
    def __init__(self, name, cats, cols, fmts, sources, texts):
        self.name, self.cats, self.cols, self.fmts, self.sources = name, cats, cols, fmts, sources
        self.texts = texts                          # 글자 열(이벤트·날짜 등)

    def series(self, *names):
        """차트에 넣을 [(이름, 값)] — 이름은 엑셀 머리행과 똑같아야 대조가 된다."""
        miss = [n for n in names if n not in self.cols]
        if miss:
            raise KeyError(f"{self.name}: 엑셀에 없는 열 {miss} (있는 열: {list(self.cols)})")
        return [(n, self.cols[n]) for n in names]

    def fmt(self, name):
        return self.fmts.get(name, "#,##0")

    @property
    def axis(self):
        """차트 가로축 글자: '축 레이블' 열이 있으면 그것(일별 데이터의 월초만 표기 등), 없으면 기간"""
        return self.texts.get("축 레이블", self.cats)


def month_labels(dates, every=1, out="%y.%m", parse="%y.%m.%d"):
    """일별 날짜 → 각 달 첫 거래일에만 'yy.mm', 나머지 '' (every=3이면 분기마다)"""
    import datetime as dt
    labs, last = [], None
    for d in dates:
        t = dt.datetime.strptime(str(d), parse)
        key = (t.year, (t.month - 1) // every)
        labs.append(t.strftime(out) if key != last else "")
        last = key
    return labs


def load(path, recalc_first=True):
    """{시트 이름: Sheet} — 수식 열은 계산된 값으로 읽는다."""
    src = recalc(path) if recalc_first else path
    wb = load_workbook(src, data_only=True)
    raw = load_workbook(path)                      # 표시 형식은 원본에서
    book = {}
    book_facts = {}
    for ws in wb.worksheets:
        if ws.title in (INDEX, "대조"):
            continue
        heads = [c.value for c in ws[1]]
        if heads[:3] == ["표기", "내용", "출처"]:
            book_facts[ws.title] = [(_txt(r[0]), _txt(r[1]), _txt(r[2]))
                                    for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
            continue
        if not heads or heads[-1] != "출처":          # 첫 열 이름은 자유(기간·구분·요구 조건…), 마지막 열 = 출처
            continue
        rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
        cats = [str(r[0]) for r in rows]
        cols, fmts, texts = {}, {}, {}
        for j, h in enumerate(heads[1:-1], 1):
            vals = [None if r[j] == "" else r[j] for r in rows]   # 수식이 ""(빈칸)를 돌려준 칸
            if all(v is None or isinstance(v, (int, float)) for v in vals):
                cols[h] = vals
                fmts[h] = raw[ws.title].cell(2, j + 1).number_format
            else:
                texts[h] = [_txt(v) for v in vals]
        book[ws.title] = Sheet(ws.title, cats, cols, fmts, [r[-1] for r in rows], texts)
    load.facts = book_facts                        # 숫자 없는 도표의 글자 목록: load.facts[시트]
    return book
