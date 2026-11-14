# -*- coding: utf-8 -*-
"""지주 양식 정본 헬퍼 (2026-11-14, 정밀 실측판)

정본: 261026 「식자재마트 현황 및 당사 비교」(전략경영부문 전략기획팀) PPTX 를 XML 단위로 실측.
지주 보고 본장표는 이 모듈로 빌드한다 (docs/HOLDING_REPORT_GUIDE.md 2~5장과 같은 값).
frames.py 의 f_* 템플릿은 이전 양식(■ 리드, 좌우 1.16)이라 지주 보고에는 쓰지 않는다.

실측 요약 (전부 이 모듈이 강제):
  글상자   wrap none, 안쪽 여백 0, 세로 가운데, 문단 앞 6pt, 높이 0.53cm (한 줄에 하나)
  글자     맑은 고딕, 검정 #000000. 괄호 부기는 자동 축소하지 않고 호출부에서 9pt #595959 로 지정
  제목     20pt 볼드 가운데 y 0.99 (부제 14pt 볼드), 밑줄 1.5pt y 2.02 x 8.40~19.11
  영역     ❏ 겹친 사각(0.35×0.34, 선 #808080 0.75pt, 뒤 #595959) + 14pt 볼드, 글자 x 기호+0.61 y 2.97
  ● 요지   11pt 볼드, Wingdings 'l' 불릿 120%, 내어쓰기 0.476cm
  표       표 스타일 '스타일 없음, 눈금 없음', 셀 여백 0, 세로 가운데, 선 #BFBFBF 0.5pt(cap flat, 실선),
           헤더 #EBEFF5 10pt 볼드(줄 간격 100%), 본문 10pt(줄 간격 130%), 1열 볼드, 합계·결론 행 #F2F2F2,
           굵은 테두리는 셀 선 1.5pt 검정(큰 표는 2.25pt). 카드 안 소표의 헤더는 #F2F2F2
  빨간 밑줄 선 도형 #FF0000 1.5pt
  회색 박스 #F2F2F2 테두리 없음, 글 11pt 볼드 줄 간격 150%(줄 피치 0.70cm), 왼쪽 안여백 0.40cm
  카드     라벨 칸 #EBEFF5 폭 2.14 12pt 볼드 줄 간격 130% + 내용 칸 폭 9.86, 선 #BFBFBF 0.5pt
           내용: ① 11pt 볼드 (칸 x+2.60, 위+0.40) / - 세부 9pt (x+3.13, 0.73 아래) / → 9~11pt
  각주     8pt, x 0.76 y 18.34

    from holding_std import *
    prs = new_deck()
    s = h_chrome(prs, "카카오 제휴 재무 효과")
    h_head(s, "L", "신규 고객 획득 비용"); h_head(s, "R", "오카도 최소수수료")
    h_bullet(s, LT, 4.07, [("제휴 활용 시 ", True), ("획득 비용 최소화", True, NAVY)])
    h_table(s, LT, 4.62, [2.4, 3.2, 3.2, 3.2], [0.87, 1.07, 1.07], rows, thick=[(0, 3, 2, 3)])
    save_deck(prs, report_filename("카카오 제휴 재무 효과", "0.1"))

좌표는 cm. 줄·셀 값: 문자열 또는 [(글, 볼드, 색, 크기)] 리스트(색·크기 생략 가능, None = 기본).
"""
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from zetta_ppt_standard import (_rect, _line, _set_font, _blank, new_deck, save_deck,
                                report_filename, WHITE)

__all__ = ["new_deck", "save_deck", "report_filename", "WHITE",
           "BLK", "INK", "HDR", "GRAY", "LINE", "SUB", "NAVY", "RED", "MARK",
           "LT", "LTW", "RT", "RTW", "CARD_RT", "BOT", "PAGE_W", "LINE_H", "CARD_LABEL_W",
           "L", "C", "R_", "TOP", "MID",
           "tw", "h_text", "h_chrome", "h_head", "h_bullet", "h_caption", "h_table",
           "h_graybox", "h_redline", "h_redline_center", "h_thick", "h_tri", "h_line",
           "h_dot", "h_card", "h_footnote", "h_summary"]

# ── 토큰 (지침 4장)
BLK = RGBColor(0, 0, 0)
INK = BLK                             # 본문 글자 = 순검정
HDR = RGBColor(0xEB, 0xEF, 0xF5)      # 표 헤더·카드 라벨 면
GRAY = RGBColor(0xF2, 0xF2, 0xF2)     # 시사점 박스, 합계·결론 행, 카드 안 소표 헤더
LINE = RGBColor(0xBF, 0xBF, 0xBF)     # 격자 0.5pt
SUB = RGBColor(0x59, 0x59, 0x59)      # 보충 글자, 기호 뒤 사각, 도식 선
MARK = RGBColor(0x80, 0x80, 0x80)     # 영역 기호 선
NAVY = RGBColor(0x00, 0x00, 0x99)     # 강조 남색 볼드
RED = RGBColor(0xFF, 0x00, 0x00)      # 빨간 밑줄 1.5pt
TRI_LINE = RGBColor(0xD9, 0xD9, 0xD9)
NO_GRID_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # 스타일 없음, 눈금 없음

# ── 그리드 (지침 2장)
PAGE_W = 27.52
LT, LTW = 1.26, 12.00                 # 좌 영역 본문
RT, RTW = 14.76, 12.00                # 우 영역 본문
CARD_RT = 14.26                       # 카드형(영역 머리글 없는 2×2) 우측 카드 시작
BOT = 17.53                           # 콘텐츠 하한선
LINE_H = 0.53                         # 한 줄 글상자 높이
CARD_LABEL_W = 2.14
L, C, R_ = PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT
TOP, MID = MSO_ANCHOR.TOP, MSO_ANCHOR.MIDDLE


def tw(text, pt, bold=False):
    """맑은 고딕 기준 글자 폭 추정(cm). 빨간 밑줄 길이·위치 계산용, 5~10% 오차는 사용자가 조정"""
    w = 0.0
    for ch in text:
        w += 0.30 if ch == " " else (0.94 if ord(ch) > 0x1100 else (0.30 if ch in ",.'()" else 0.56))
    return w * pt * 0.03528 * (1.04 if bold else 1.0)


def _segs(ln, bold):
    return [(ln, bold)] if isinstance(ln, str) else ln


def _runs(p, segs, size, color):
    """괄호 자동 축소 없이 지정한 크기 그대로 run 생성 (정본 덱과 같은 방식)"""
    for seg in segs:
        r = p.add_run(); r.text = seg[0]
        col = seg[2] if len(seg) > 2 and seg[2] is not None else color
        _set_font(r, seg[3] if len(seg) > 3 and seg[3] is not None else size, seg[1], col)


def _para_fmt(p, ln_pct=None, bef_pt=None):
    """문단 줄 간격(%)·문단 앞 간격(pt). pPr 자식 순서: lnSpc, spcBef, ..."""
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:lnSpc", "a:spcBef"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    pos = 0
    if ln_pct is not None:
        el = pPr.makeelement(qn("a:lnSpc"), {})
        el.append(el.makeelement(qn("a:spcPct"), {"val": str(int(ln_pct * 1000))}))
        pPr.insert(pos, el); pos += 1
    if bef_pt is not None:
        el = pPr.makeelement(qn("a:spcBef"), {})
        el.append(el.makeelement(qn("a:spcPts"), {"val": str(int(bef_pt * 100))}))
        pPr.insert(pos, el)


def _zero_inset(tf):
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)


def h_text(s, l, t, w, h, lines, size=11, bold=False, color=INK, align=L, anchor=MID, wrap=False):
    """글상자: 안쪽 여백 0, 세로 가운데, wrap none(한 줄에 하나), 문단 앞 6pt.
    lines = 문자열 | [줄], 줄 = 문자열 | [(글, 볼드, 색, 크기)]"""
    tb = _rect(s, l, t, w, h, fill=None, line=None)      # 정본과 같이 채움·선 없는 사각형 도형에 글을 넣음
    tf = tb.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    _zero_inset(tf)
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        _para_fmt(p, bef_pt=6)
        _runs(p, _segs(ln, bold), size, color)
    return tb


def h_chrome(prs, title, sub=None):
    """제목 20pt 볼드 가운데(y 0.99, 높이 0.89) + 밑줄 1.5pt(y 2.02, x 8.40~19.11, 긴 제목 7.28~20.24).
    sub = 시리즈 부제. 「주제 _ ① 부제」로 쓰고 부제는 14pt 볼드. ■ 리드는 두지 않는다."""
    s = _blank(prs)
    segs = [(title, True)] + ([(" _ ", True), (sub, True, None, 14)] if sub else [])
    h_text(s, 1.76, 0.99, 24.0, 0.89, [segs], size=20, bold=True, align=C, wrap=True)
    used = tw(title, 20, True) + (tw(" _ ", 20, True) + tw(sub, 14, True) if sub else 0)
    w = 10.71 if used < 10.2 else 12.96
    _line(s, (PAGE_W - w) / 2, 2.02, (PAGE_W + w) / 2, 2.02, color=BLK, w_pt=1.5)
    return s


def h_head(s, side, text):
    """❏ 영역 머리글 14pt 볼드 좌측. side = "L" | "R" 또는 기호 x 좌표(cm)"""
    x = {"L": 0.76, "R": 14.26}.get(side, side)
    _rect(s, x + 0.05, 3.11, 0.35, 0.34, fill=SUB, line=MARK, line_w=0.75)
    _rect(s, x, 3.05, 0.35, 0.34, fill=WHITE, line=MARK, line_w=0.75)
    h_text(s, x + 0.61, 2.97, 11.38, LINE_H, text, size=14, bold=True)


def h_bullet(s, x, y, segs, w=11.38, size=11):
    """● 요지 한 줄 11pt 볼드. Wingdings 'l' 불릿 120%, 내어쓰기 0.476cm (영역 안 y 4.07, 카드형 전폭 y 2.85 / 3.75)"""
    segs = [(segs, True)] if isinstance(segs, str) else list(segs)
    tb = h_text(s, x, y, w, LINE_H, [[(" ", True)] + segs], size=size, bold=True)
    p = tb.text_frame.paragraphs[0]
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "171450"); pPr.set("indent", "-171450")
    for tag, attrs in (("a:buSzPct", {"val": "120000"}),
                       ("a:buFont", {"typeface": "Wingdings", "pitchFamily": "2", "charset": "2"}),
                       ("a:buChar", {"char": "l"})):
        pPr.append(pPr.makeelement(qn(tag), attrs))
    return tb


def h_caption(s, x, y, text, size=10, color=INK, bold=False, w=11.38):
    """표 위 기준 표기: (A Vs. B) 10pt, ※ 9pt, [ 소제목 ] 11pt 볼드, (예시) 9pt #595959"""
    return h_text(s, x, y, w, LINE_H, text, size=size, color=color, bold=bold)


_EDGES = (("L", "a:lnL"), ("R", "a:lnR"), ("T", "a:lnT"), ("B", "a:lnB"))


def _borders(cell, spec):
    """spec = {"L"|"R"|"T"|"B": (색, pt)}. 정본과 같은 선 속성(cap flat, 단선, 실선)으로 lnL·lnR·lnT·lnB 순서 기록"""
    tcPr = cell._tc.get_or_add_tcPr()
    for _, tag in _EDGES:
        el = tcPr.find(qn(tag))
        if el is not None:
            tcPr.remove(el)
    for pos, (key, tag) in enumerate(_EDGES):
        color, pt = spec[key]
        ln = tcPr.makeelement(qn(tag), {"w": str(int(pt * 12700)), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        f = ln.makeelement(qn("a:solidFill"), {})
        f.append(f.makeelement(qn("a:srgbClr"), {"val": str(color)}))
        ln.append(f)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
        ln.append(ln.makeelement(qn("a:round"), {}))
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "none", "w": "med", "len": "med"}))
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "none", "w": "med", "len": "med"}))
        tcPr.insert(pos, ln)


def h_table(s, l, t, cw, rh, data, hdr=1, aligns=None, size=10, hdr_size=None, merges=None, fills=None,
            bold_cols=(0,), gray_rows=(), hdr_fill=HDR, thick=None, thick_pt=1.5):
    """표 (정본 실측): 스타일 없음·눈금 없음, 셀 여백 0, 세로 가운데, 선 #BFBFBF 0.5pt.
    헤더 hdr_fill(기본 #EBEFF5, 카드 안 소표는 GRAY) 볼드 가운데 줄 간격 100%, 본문 줄 간격 130%, 1열 볼드.
    cw·rh = 열 폭·행 높이(cm). gray_rows = 합계·결론 행 번호(면 #F2F2F2 + 볼드).
    thick = [(r0, c0, r1, c1)] 굵은 검정 테두리로 감쌀 셀 범위(셀 선으로 구현, thick_pt 1.5 또는 2.25).
    fills = {(r, c): RGB}. aligns = 열별 "l"/"c"/"r" (헤더는 항상 가운데). 행 높이 기준: 헤더 0.82~0.97, 본문 1.07"""
    nr, nc = len(data), len(cw)
    g = s.shapes.add_table(nr, nc, Cm(l), Cm(t), Cm(sum(cw)), Cm(sum(rh)))
    tb = g.table
    tblPr = tb._tbl.tblPr
    tblPr.set("firstRow", "1"); tblPr.set("bandRow", "1")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = tblPr.makeelement(qn("a:tableStyleId"), {}); tblPr.append(sid)
    sid.text = NO_GRID_STYLE
    for j, w in enumerate(cw):
        tb.columns[j].width = Cm(w)
    for i, h in enumerate(rh):
        tb.rows[i].height = Cm(h)
    edge = {(i, j): {k: (LINE, 0.5) for k in "LRTB"} for i in range(nr) for j in range(nc)}
    for (r0, c0, r1, c1) in (thick or []):
        for i in range(r0, r1 + 1):
            edge[(i, c0)]["L"] = (BLK, thick_pt); edge[(i, c1)]["R"] = (BLK, thick_pt)
            if c0 > 0:
                edge[(i, c0 - 1)]["R"] = (BLK, thick_pt)
            if c1 < nc - 1:
                edge[(i, c1 + 1)]["L"] = (BLK, thick_pt)
        for j in range(c0, c1 + 1):
            edge[(r0, j)]["T"] = (BLK, thick_pt); edge[(r1, j)]["B"] = (BLK, thick_pt)
            if r0 > 0:
                edge[(r0 - 1, j)]["B"] = (BLK, thick_pt)
            if r1 < nr - 1:
                edge[(r1 + 1, j)]["T"] = (BLK, thick_pt)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.vertical_anchor = MID
            c.margin_left = c.margin_right = c.margin_top = c.margin_bottom = 0
            base = hdr_fill if i < hdr else (GRAY if i in gray_rows else WHITE)
            c.fill.solid(); c.fill.fore_color.rgb = (fills or {}).get((i, j), base)
            _borders(c, edge[(i, j)])
            if val is None:
                continue
            tf = c.text_frame; tf.word_wrap = True
            lines = val.split("\n") if isinstance(val, str) else val
            al = C if i < hdr else {"l": L, "c": C, "r": R_}[(aligns or "c" * nc)[j]]
            fs = (hdr_size or size) if i < hdr else size
            for k, ln in enumerate(lines):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph(); p.alignment = al
                _para_fmt(p, ln_pct=100 if i < hdr else 130, bef_pt=0)
                _runs(p, _segs(ln, i < hdr or j in bold_cols or i in gray_rows), fs, INK)
    for (r0, c0), (r1, c1) in (merges or []):
        tb.cell(r0, c0).merge(tb.cell(r1, c1))
    return tb


def h_graybox(s, l, t, w, h, lines=None, size=11, pad=0.40, pitch=0.70):
    """회색 시사점 박스(#F2F2F2, 테두리 없음). lines 를 주면 세로 가운데에 줄 피치 0.70cm(11pt, 줄 간격 150%)로
    줄마다 글상자 하나를 놓고 각 줄의 (x, 밑줄 y) 리스트를 돌려준다(빨간 밑줄 위치 맞춤용)."""
    _rect(s, l, t, w, h, fill=GRAY, line=None)
    pos = []
    if lines:
        y0 = t + (h - pitch * len(lines)) / 2
        for k, ln in enumerate(lines):
            y = y0 + k * pitch + (pitch - LINE_H) / 2
            h_text(s, l + pad, y, w - pad * 2, LINE_H, [_segs(ln, True)], size=size, bold=True)
            pos.append((l + pad, y + LINE_H + 0.04))
    return pos


def h_redline(s, x, y, w):
    """빨간 밑줄 = 선 도형 #FF0000 1.5pt (글자 서식 underline 금지). y = 글상자 아랫변 + 0.04 안팎"""
    return _line(s, x, y, x + w, y, color=RED, w_pt=1.5)


def h_redline_center(s, cx, y, text, size, bold=True, pad=0.25):
    """가운데 정렬 글자(표 셀 값) 아래 빨간 밑줄. cx = 셀 중심 x"""
    w = tw(text, size, bold) + pad * 2
    return h_redline(s, cx - w / 2, y, w)


def h_thick(s, l, t, w, h, pt=1.5):
    """굵은 검정 테두리 사각(표 밖 묶음용). 표 안의 열·행 묶음은 h_table(thick=...) 로 셀 선에 넣는다"""
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(l), Cm(t), Cm(w), Cm(h))
    sp.fill.background(); sp.line.color.rgb = BLK; sp.line.width = Pt(pt); sp.shadow.inherit = False
    return sp


def h_tri(s, cx, t, w=4.15, h=0.40):
    """회색 역삼각: 표 → 시사점 박스 연결 (면 #F2F2F2, 선 #D9D9D9 0.5pt)"""
    sp = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Cm(cx - w / 2), Cm(t), Cm(w), Cm(h))
    sp.rotation = 180
    sp.fill.solid(); sp.fill.fore_color.rgb = GRAY
    sp.line.color.rgb = TRI_LINE; sp.line.width = Pt(0.5); sp.shadow.inherit = False
    return sp


def h_line(s, x1, y1, x2, y2, color=SUB, w_pt=0.5):
    """직선(사선 가능). 도식 선은 정본 값만 사용: #595959 0.5pt(화살표·축), 1.0pt(구분), 검정 1.5pt(강조)"""
    ln = s.shapes.add_connector(1, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.shadow.inherit = False; ln.line.color.rgb = color; ln.line.width = Pt(w_pt)
    return ln


def h_dot(s, x, y, d=0.20, fill=BLK):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x - d / 2), Cm(y - d / 2), Cm(d), Cm(d))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.color.rgb = BLK; sp.line.width = Pt(0.75); sp.shadow.inherit = False
    return sp


def h_card(s, l, t, label, h=5.81, w=12.00):
    """카드 1개: 라벨 칸(#EBEFF5, 폭 2.14, 12pt 볼드 가운데, 줄 간격 130%) + 내용 칸(흰색), 선 #BFBFBF 0.5pt.
    내용 기준 좌표 (x1, x2, y) 반환: ① 줄 x1 = l+2.60, - 세부 x2 = l+3.13, 첫 줄 y = t+0.40.
    줄 간격: ① 다음 - 는 +0.73, - 다음 - 는 +0.65, - 다음 ① 은 +0.89 (정본 7/10)"""
    lab = _rect(s, l, t, CARD_LABEL_W, h, fill=HDR, line=LINE, line_w=0.5)
    tf = lab.text_frame; tf.word_wrap = False; tf.vertical_anchor = MID
    for i, ln in enumerate(label.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = C
        _para_fmt(p, ln_pct=130)
        _runs(p, [(ln, True)], 12, INK)
    _rect(s, l + CARD_LABEL_W, t, w - CARD_LABEL_W, h, fill=WHITE, line=LINE, line_w=0.5)
    return l + 2.60, l + 3.13, t + 0.40


def h_footnote(s, text):
    """좌하단 번호 각주 8pt (「1. 내용」, 본문 위첨자 번호와 대응)"""
    return h_text(s, 0.76, 18.34, 20.0, 0.37, text, size=8)


def h_summary(prs, paragraphs, title="- Summary -"):
    """덱 끝 Summary 장: 좌측 제목 20pt 볼드(x 0.46, y 1.05, 밑줄 없음) + ■ 12pt 볼드 문단(각 2~3줄).
    paragraphs = [문단], 문단 = [줄], 줄 = 문자열 | [(글, 볼드, 색, 크기)]. 괄호 부기는 (글, False, None, 10)"""
    s = _blank(prs)
    h_text(s, 0.46, 1.05, 11.91, 1.11, title, size=20, bold=True, anchor=TOP)
    n = len(paragraphs); pitch = (BOT - 3.17) / n
    for i, para in enumerate(paragraphs):
        y = 3.17 + pitch * i
        for k, ln in enumerate(para):
            segs = list(_segs(ln, True))
            if k == 0:
                segs = [("■  ", True, None, 9)] + segs
            h_text(s, 1.22 if k == 0 else 1.76, y + k * 0.95, 24.5, 0.86, [segs], size=12, bold=True)
    return s
