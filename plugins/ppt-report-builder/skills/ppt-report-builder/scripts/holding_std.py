# -*- coding: utf-8 -*-
"""지주 양식 정본 헬퍼 (2026-11-14)

정본: 261026 「식자재마트 현황 및 당사 비교」(전략경영부문 전략기획팀) PPTX 실측.
지주 보고 본장표는 이 모듈로 빌드한다 (docs/HOLDING_REPORT_GUIDE.md 2~5장과 같은 값).
frames.py 의 f_* 템플릿은 이전 양식(■ 리드, 좌우 1.16)이라 지주 보고에는 쓰지 않는다.

    from holding_std import *
    prs = new_deck()
    s = h_chrome(prs, "카카오 제휴 재무 효과")            # 제목 + 밑줄
    h_head(s, "L", "신규 고객 획득 비용")                  # ❏ 영역 머리글 (좌)
    h_head(s, "R", "오카도 최소수수료")                    # ❏ 영역 머리글 (우)
    h_bullet(s, LT, 4.07, [("제휴 활용 시 ", True), ("획득 비용 최소화", True, NAVY)])
    h_caption(s, LT, 4.82, "(자체 획득 Vs. 카카오 제휴)")
    h_table(s, LT, 5.38, [2.5, 3.2, 3.2, 3.1], [0.8, 1.07, 1.07], rows)
    h_thick(s, LT + 8.9, 5.38, 3.1, 2.94)                  # 핵심 열 굵은 검정 테두리
    h_graybox(s, LT, 15.9, LTW, BOT - 15.9)                # 시사점 박스
    save_deck(prs, report_filename("카카오 제휴 재무 효과", "0.1"))

좌표는 cm. 좌 영역 본문 LT~LT+LTW (1.26~13.26), 우 영역 RT~RT+RTW (14.76~26.76), 하한 BOT 17.53.
셀·줄 값: 문자열 또는 [(글, 볼드, 색, 크기)] 리스트(색·크기는 생략 가능).
"""
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from zetta_ppt_standard import (_rect, _line, _emit_runs, _blank, new_deck, save_deck,
                                report_filename, INK, WHITE)

__all__ = ["new_deck", "save_deck", "report_filename", "INK", "WHITE",
           "BLK", "HDR", "GRAY", "LINE", "SUB", "NAVY", "RED", "MARK",
           "LT", "LTW", "RT", "RTW", "CARD_RT", "BOT", "PAGE_W",
           "L", "C", "R_", "TOP", "MID",
           "tw", "h_text", "h_chrome", "h_head", "h_bullet", "h_caption", "h_table",
           "h_graybox", "h_redline", "h_redline_center", "h_thick", "h_tri", "h_line",
           "h_dot", "h_card", "h_footnote", "h_summary"]

# ── 토큰 (지침 4장)
BLK = RGBColor(0, 0, 0)
HDR = RGBColor(0xEB, 0xEF, 0xF5)      # 표 헤더·카드 라벨 면
GRAY = RGBColor(0xF2, 0xF2, 0xF2)     # 시사점 박스, 합계·결론 행, 표 안 설명 박스
LINE = RGBColor(0xBF, 0xBF, 0xBF)     # 격자 0.5pt
SUB = RGBColor(0x59, 0x59, 0x59)      # 보충 글자, 기호 뒤 사각
MARK = RGBColor(0x80, 0x80, 0x80)     # 영역 기호 선
NAVY = RGBColor(0x00, 0x00, 0x99)     # 강조 남색 볼드
RED = RGBColor(0xFF, 0x00, 0x00)      # 빨간 밑줄 1.5pt
TRI_LINE = RGBColor(0xD9, 0xD9, 0xD9)

# ── 그리드 (지침 2장)
PAGE_W = 27.52
LT, LTW = 1.26, 12.00                 # 좌 영역 본문
RT, RTW = 14.76, 12.00                # 우 영역 본문
CARD_RT = 14.26                       # 카드형 우측 카드 시작
BOT = 17.53                           # 콘텐츠 하한선
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


def h_text(s, l, t, w, h, lines, size=11, bold=False, color=INK, align=L, anchor=MID):
    """여백 0 글상자. lines = 문자열 | [줄], 줄 = 문자열 | [(글, 볼드, 색, 크기)]"""
    tb = s.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for seg in _segs(ln, bold):
            col = seg[2] if len(seg) > 2 and seg[2] is not None else color
            _emit_runs(p, seg[0], seg[3] if len(seg) > 3 else size, seg[1], col)
    return tb


def h_chrome(prs, title, sub=None):
    """제목 20pt 볼드 가운데(y 0.99) + 밑줄 1.5pt(y 2.02, 폭 10.71 고정, 긴 제목 12.96).
    sub = 시리즈 부제. 「주제 _ ① 부제」로 쓰고 부제는 14pt 볼드. ■ 리드는 두지 않는다."""
    s = _blank(prs)
    segs = [(title, True)] + ([("  _ " + sub, True, None, 14)] if sub else [])
    h_text(s, 1.76, 0.99, 24.0, 0.89, [segs], size=20, bold=True, align=C)
    used = tw(title, 20, True) + (tw("  _ " + sub, 14, True) if sub else 0)
    w = 10.71 if used < 10.2 else 12.96
    _line(s, (PAGE_W - w) / 2, 2.02, (PAGE_W + w) / 2, 2.02, color=BLK, w_pt=1.5)
    return s


def h_head(s, side, text):
    """❏ 영역 머리글 14pt 볼드 좌측. side = "L" | "R" 또는 기호 x 좌표(cm)"""
    x = {"L": 0.76, "R": 14.26}.get(side, side)
    _rect(s, x + 0.05, 3.11, 0.35, 0.34, fill=SUB, line=MARK, line_w=0.75)
    _rect(s, x, 3.05, 0.35, 0.34, fill=WHITE, line=MARK, line_w=0.75)
    h_text(s, x + 0.61, 2.97, 11.38, 0.53, text, size=14, bold=True)


def h_bullet(s, x, y, segs, w=11.6, size=11):
    """● 요지 한 줄 11pt 볼드 (영역 안 y 4.07, 카드형 전폭은 y 2.85 / 3.75)"""
    segs = [(segs, True)] if isinstance(segs, str) else list(segs)
    h_text(s, x, y, w, 0.53, [[("●  ", True, None, size - 2)] + segs], size=size, bold=True)


def h_caption(s, x, y, text, size=10, color=INK, bold=False, w=11.6):
    """표 위 기준 표기: (A Vs. B) 10pt, ※ 9pt, [ 소제목 ] 11pt 볼드, (예시) 9pt 회색"""
    h_text(s, x, y, w, 0.53, text, size=size, color=color, bold=bold)


def _border(cell, color=LINE, w=0.5):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        el = tcPr.find(qn(tag))
        if el is not None:
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {"w": str(int(w * 12700)), "cap": "flat"})
        f = ln.makeelement(qn("a:solidFill"), {})
        f.append(f.makeelement(qn("a:srgbClr"), {"val": str(color)}))
        ln.append(f); tcPr.insert(0, ln)


def h_table(s, l, t, cw, rh, data, hdr=1, aligns=None, size=10, merges=None, fills=None,
            bold_cols=(0,), gray_rows=()):
    """표: 헤더 #EBEFF5 볼드 가운데, 격자 #BFBFBF 0.5pt, 본문 흰색, 1열 볼드.
    cw·rh = 열 폭·행 높이(cm) 리스트. gray_rows = 합계·결론 행 번호(면 #F2F2F2).
    fills = {(r, c): RGB}. aligns = 열별 "l"/"c"/"r" 문자열(헤더는 항상 가운데)."""
    g = s.shapes.add_table(len(data), len(cw), Cm(l), Cm(t), Cm(sum(cw)), Cm(sum(rh)))
    tb = g.table; tb.first_row = False; tb.horz_banding = False
    for j, w in enumerate(cw):
        tb.columns[j].width = Cm(w)
    for i, h in enumerate(rh):
        tb.rows[i].height = Cm(h)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            c = tb.cell(i, j); _border(c)
            c.vertical_anchor = MID
            c.margin_left = Cm(0.15); c.margin_right = Cm(0.10)
            c.margin_top = Cm(0.02); c.margin_bottom = Cm(0.02)
            base = HDR if i < hdr else (GRAY if i in gray_rows else WHITE)
            c.fill.solid(); c.fill.fore_color.rgb = (fills or {}).get((i, j), base)
            if val is None:
                continue
            tf = c.text_frame; tf.word_wrap = True
            lines = val.split("\n") if isinstance(val, str) else val
            al = C if i < hdr else {"l": L, "c": C, "r": R_}[(aligns or "c" * len(cw))[j]]
            for k, ln in enumerate(lines):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph(); p.alignment = al
                for seg in _segs(ln, i < hdr or j in bold_cols or i in gray_rows):
                    col = seg[2] if len(seg) > 2 and seg[2] is not None else INK
                    _emit_runs(p, seg[0], seg[3] if len(seg) > 3 else size, seg[1], col)
    for (r0, c0), (r1, c1) in (merges or []):
        tb.cell(r0, c0).merge(tb.cell(r1, c1))
    return tb


def h_graybox(s, l, t, w, h, lines=None, size=11, pad=0.35):
    """회색 시사점 박스(#F2F2F2, 테두리 없음). lines 를 주면 줄마다 글상자 하나(빨간 밑줄 위치 맞춤용)"""
    _rect(s, l, t, w, h, fill=GRAY, line=None)
    if lines:
        lh = (h - 0.2) / len(lines)
        for k, ln in enumerate(lines):
            h_text(s, l + pad, t + 0.1 + k * lh, w - pad * 2, lh, [_segs(ln, True)], size=size, bold=True)


def h_redline(s, x, y, w):
    """빨간 밑줄 = 선 도형 #FF0000 1.5pt (글자 서식 underline 금지)"""
    _line(s, x, y, x + w, y, color=RED, w_pt=1.5)


def h_redline_center(s, cx, y, text, size, bold=True, pad=0.25):
    """가운데 정렬 글자(표 셀 값) 아래 빨간 밑줄. cx = 셀 중심 x"""
    w = tw(text, size, bold) + pad * 2
    h_redline(s, cx - w / 2, y, w)


def h_thick(s, l, t, w, h, pt=1.5):
    """굵은 검정 테두리(1.5~2.25pt): 표에서 결론이 되는 열·행 묶음을 감쌈. 영역당 1곳"""
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(l), Cm(t), Cm(w), Cm(h))
    sp.fill.background(); sp.line.color.rgb = BLK; sp.line.width = Pt(pt); sp.shadow.inherit = False
    return sp


def h_tri(s, cx, t, w=4.15, h=0.40):
    """회색 역삼각: 표 → 시사점 박스 연결"""
    sp = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Cm(cx - w / 2), Cm(t), Cm(w), Cm(h))
    sp.rotation = 180
    sp.fill.solid(); sp.fill.fore_color.rgb = GRAY
    sp.line.color.rgb = TRI_LINE; sp.line.width = Pt(0.5); sp.shadow.inherit = False
    return sp


def h_line(s, x1, y1, x2, y2, color=SUB, w_pt=0.5):
    """직선(사선 가능). 도식의 축·추세선"""
    ln = s.shapes.add_connector(1, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.shadow.inherit = False; ln.line.color.rgb = color; ln.line.width = Pt(w_pt)
    return ln


def h_dot(s, x, y, d=0.22, fill=BLK):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x - d / 2), Cm(y - d / 2), Cm(d), Cm(d))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.color.rgb = BLK; sp.line.width = Pt(1); sp.shadow.inherit = False
    return sp


def h_card(s, l, t, label, h=5.81, label_w=2.14, w=12.00):
    """카드형 장표의 카드 1개: 라벨 칸(#EBEFF5, 12pt 볼드) + 내용 칸(흰색). 내용 칸 (x, y, 폭) 반환.
    2×2 배치: x = LT / CARD_RT, y = 5.03 / 11.72"""
    _rect(s, l, t, label_w, h, fill=HDR, line=LINE, line_w=0.5)
    h_text(s, l, t, label_w, h, label.split("\n"), size=12, bold=True, align=C)
    _rect(s, l + label_w, t, w - label_w, h, fill=WHITE, line=LINE, line_w=0.5)
    return l + label_w + 0.46, t + 0.40, w - label_w - 0.9


def h_footnote(s, text):
    """좌하단 번호 각주 8pt (「1. 내용」, 본문 위첨자 번호와 대응)"""
    h_text(s, 0.76, 18.34, 20.0, 0.37, text, size=8)


def h_summary(prs, paragraphs, title="- Summary -"):
    """덱 끝 Summary 장: 좌측 제목 20pt 볼드(밑줄 없음) + ■ 12pt 볼드 문단(각 2~3줄).
    paragraphs = [문단], 문단 = [줄], 줄 = 문자열 | [(글, 볼드, 색, 크기)]. 괄호 부기는 (글, False, None, 10)"""
    s = _blank(prs)
    h_text(s, 0.46, 1.05, 11.91, 1.11, title, size=20, bold=True, anchor=TOP)
    n = len(paragraphs); pitch = (BOT - 3.17) / n
    for i, para in enumerate(paragraphs):
        y = 3.17 + pitch * i
        for k, ln in enumerate(para):
            segs = _segs(ln, True)
            if k == 0:
                segs = [("■  ", True, None, 9)] + list(segs)
            h_text(s, 1.22 if k == 0 else 1.76, y + k * 0.95, 24.5, 0.86, [segs], size=12, bold=True)
    return s
