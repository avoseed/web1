# -*- coding: utf-8 -*-
"""
zetta_ppt_standard.py  —  롯데마트 온라인/ZETTA 표준 PPT 빌더 (v4)
================================================================
'_양식__내부토의_공통양식_v3_' 파일과 '11차 Ocado 정기협의체' 산출물을
역설계하여 확정한 디자인 토큰·슬라이드 원형을 코드로 encode.

설계 원칙
  - 모든 좌표는 cm (실측 기준값). EMU 변환은 Cm() 이 처리.
  - 폰트는 맑은 고딕 3중 지정(latin·ea·cs) — v4 확정. (v3 테마 latin=Arial 을 교정)
  - 슬라이드 크롬(헤더/단위/페이지번호/각주)은 좌표 고정 → 전 페이지 정합.
  - 원형(archetype) 함수는 slide 를 반환 → 호출측에서 본문 자유 배치.
"""
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ─────────────────────────────────────────────────────────────
# 1. 디자인 토큰 (역설계 확정값)
# ─────────────────────────────────────────────────────────────
PAGE_W_CM, PAGE_H_CM = 27.52, 19.05          # A4 가로 (9906000 x 6858000 EMU)
FONT       = "맑은 고딕"                       # 단일 폰트 — 한글·영문·숫자 공통 3중 지정 (v4.1)

# 팔레트 (theme1.xml 실측)
NAVY       = RGBColor(0x00, 0x00, 0x66)       # accent1  [F계열] 제목·강조
CRIMSON    = RGBColor(0xC3, 0x0C, 0x3E)       # accent3  경고·리스크 강조
BLACK      = RGBColor(0x22, 0x22, 0x22)       # [P계열] 제목
INK        = RGBColor(0x00, 0x00, 0x00)       # 본문 기본
RED_GUIDE  = RGBColor(0xFF, 0x00, 0x00)       # [확인] 가이드(최종본 삭제 대상)
TH_PRIMARY = RGBColor(0xEA, 0xEE, 0xF6)       # 표 헤더 1차
TH_SECOND  = RGBColor(0xD8, 0xE0, 0xEC)       # 표 헤더 2차
TH_DARK    = RGBColor(0x00, 0x1E, 0x62)       # 강조 헤더(짙은 네이비)
GRID       = RGBColor(0xBF, 0xBF, 0xBF)       # 표 격자
SUB_GRAY   = RGBColor(0x59, 0x59, 0x59)       # 보조 텍스트·시사점 탭 회색
DIVIDER_BG = RGBColor(0xF2, 0xF2, 0xF2)       # 챕터 간지 회색 풀블리드(bg1 lumMod95%)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# 크롬 좌표 (실측 기반, v4.1 여백 +0.5mm 조정)  L, T, W, H  (cm)
HDR   = (1.21, 1.20, 25.14, 0.90)             # 좌상단 브래킷 헤더
LEAD  = (1.21, 2.90, 21.75, 0.65)             # 제목 하단 리드메시지 (■ 한 줄, 헤더 하단 +0.80)
UNIT  = (23.04, 2.35, 3.30, 0.60)             # 우상단 단위 (우측 여백선 정렬)
FOOT  = (1.09, 17.90, 18.00, 0.55)            # 좌하단 각주
MARGIN_L, MARGIN_R = 1.18, 1.18               # 좌우 여백(대칭, v4.1 +0.5mm)
BODY_W = PAGE_W_CM - MARGIN_L - MARGIN_R      # 본문 폭 25.16
BODY_TOP = 3.00                               # 본문 시작 y (리드메시지 미사용 시)

# 수직 간격 토큰 (정기협의체 산출물 실측 — 개체 간 간격 표준)
BODY_TOP_LEAD = 4.10                          # 본문 시작 y (리드메시지 사용 시, 리드 하단 +0.55)
GAP_SECTION   = 0.50                          # 본문 블록(소제목 묶음) 종료 → 다음 소제목 간격
GAP_HEAD      = 0.55                          # 소제목 시작 → 첫 본문 요소 시작 (피치)
LINE_H        = 0.48                          # 개조식 불릿 행간 (피치)

# 개조식 불릿 체계 (정기협의체 실측): ■ 리드(크롬) → • 주항목(볼드) → - 하위 → · 세부
BULLET_MARKS  = ("•", "-", "·")
BULLET_INDENT = 0.55                          # 하위 수준당 들여쓰기 (cm)

# 본장 2단 컬럼 프레임 (정기협의체 8·14/15 실측): 헤더 바 + 중앙 세로선 + 결론 박스
COL_GAP      = 0.66                           # 컬럼 간 간격
COL_HEADER_H = 0.75                           # 컬럼 헤더 바 높이
CONCL_H      = 0.90                           # 컬럼 하단 결론 박스 높이

# 제목 크기 (tier 별) — v4.1 교정: v3 실측(20pt 단일)로 회귀, tier 는 색상만 구분
TITLE_PT = {"P": 20, "F": 20}                 # P계열 검정 / F계열 네이비, 공통 20pt

# 표준 타이포 스케일 (v4.1, 정기협의체 실측) — 임의 크기 사용 금지
FONT_PT = {
    "title":    20,   # 브래킷 헤더 (볼드, P 검정 #222222 / F 네이비 #000066)
    "lead":     14,   # ■ 리드메시지 (볼드 검정)
    "bullet0":  12,   # • 주항목 (볼드 검정)
    "bullet1":  11,   # - 하위 (일반 검정)
    "bullet2":  10,   # · 세부 (일반 검정)
    "table":    10,   # 표 본문 (헤더: 네이비 볼드 동일 크기)
    "unit":     10,   # 우상단 단위 (볼드)
    "footnote":  9,   # 좌하단 각주
}


# ─────────────────────────────────────────────────────────────
# 2. 저수준 헬퍼
# ─────────────────────────────────────────────────────────────
def _set_font(run, size=None, bold=None, color=None, name=FONT):
    """맑은 고딕 3중 지정(latin·ea·cs) + 속성. v4 폰트 규칙."""
    f = run.font
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if color is not None:
        f.color.rgb = color
    f.name = name                              # latin
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):               # ea + cs 를 명시적으로 동일 지정
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def _txt(slide, l, t, w, h, text="", size=11, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=FONT, wrap=True):
    """텍스트 박스 1개 생성 → shape 반환. 내부 패딩 0."""
    tb = slide.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        _set_font(r, size, bold, color, name)
    return tb


def _cell_border(cell, color=None, w_pt=0.75):
    """셀 4방(lnL·lnR·lnT·lnB) 실선 격자 — 스키마 순서 보존 위해 tcPr 선두 삽입."""
    hexval = str(color) if color is not None else str(GRID)
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):     # insert(0) 역순 → 최종 L,R,T,B
        el = tcPr.find(qn(tag))
        if el is not None:
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {"w": str(int(w_pt * 12700)), "cap": "flat"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": hexval})
        fill.append(clr)
        ln.append(fill)
        tcPr.insert(0, ln)


def _rect(slide, l, t, w, h, fill=None, line=None, line_w=0.75):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(l), Cm(t), Cm(w), Cm(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    return sp


# ─────────────────────────────────────────────────────────────
# 3. 슬라이드 크롬 (전 페이지 공통)
# ─────────────────────────────────────────────────────────────
def add_header(slide, category, title, tier="P"):
    """좌상단 브래킷 헤더: [분류] 제목.  tier P=검정, F=네이비."""
    l, t, w, h = HDR
    color = BLACK if tier == "P" else NAVY
    size = TITLE_PT[tier]
    tb = slide.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
    tf = tb.text_frame; tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r1 = p.add_run(); r1.text = f"[{category}] "
    _set_font(r1, size, True, color)
    r2 = p.add_run(); r2.text = title
    _set_font(r2, size, True, color)
    return tb


def add_lead(slide, text, size=None):
    """제목 하단 리드메시지: `■ ` 접두 볼드 검정 한 줄 (v4.1, 산출물 실측 반영)."""
    l, t, w, h = LEAD
    return _txt(slide, l, t, w, h, "■ " + text,
                size=size or FONT_PT["lead"], bold=True, color=INK)


def add_bullets(slide, items, l=MARGIN_L, t=BODY_TOP, w=BODY_W,
                line_h=LINE_H, bold_top=True):
    """표준 개조식 불릿 배치 (v4.1). items: [(level, text), ...] → 종료 y 반환.

    수준별 마크·크기: 0 `•` 12pt볼드 / 1 `-` 11pt / 2 `·` 10pt (FONT_PT 스케일 고정).
    □·○ 등 이형 마크, 임의 크기 사용 금지.
    """
    sizes = (FONT_PT["bullet0"], FONT_PT["bullet1"], FONT_PT["bullet2"])
    y = t
    for level, text in items:
        lv = min(level, len(BULLET_MARKS) - 1)
        _txt(slide, l + level * BULLET_INDENT, y, w - level * BULLET_INDENT, line_h,
             f"{BULLET_MARKS[lv]} {text}", size=sizes[lv],
             bold=(level == 0 and bold_top), color=INK)
        y += line_h
    return y


def add_block_header(slide, l, t, w, text):
    """[본장 블록 소제목] 【 제목 】 — 12pt 볼드 검정 (5/15·Link 자료 실측)."""
    return _txt(slide, l, t, w, 0.55, f"【 {text} 】",
                size=FONT_PT["bullet0"], bold=True, color=INK)


def add_hline(slide, l, t, w, color=None, w_pt=0.75):
    """가로 괘선 — 구분 매트릭스(D형) 행 구분·제목 밑줄 룰."""
    ln = slide.shapes.add_connector(2, Cm(l), Cm(t), Cm(l + w), Cm(t))
    ln.line.color.rgb = color if color is not None else GRID
    ln.line.width = Pt(w_pt)
    return ln


def add_timeline(slide, l, t, w, milestones):
    """[경과형 E] 가로 타임라인: 굵은 축 + 마일스톤 점, 상단 날짜·하단 라벨.

    milestones: [(날짜, 라벨), ...] → 종료 y 반환. (OSP 경과보고·Coles 협업 실측)
    """
    axis_y = t + 0.55
    ax = slide.shapes.add_connector(2, Cm(l), Cm(axis_y), Cm(l + w), Cm(axis_y))
    ax.line.color.rgb = INK; ax.line.width = Pt(2.0)
    step = w / len(milestones)
    for i, (date, label) in enumerate(milestones):
        cx = l + step * (i + 0.5)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(cx - 0.11),
                                     Cm(axis_y - 0.11), Cm(0.22), Cm(0.22))
        dot.shadow.inherit = False
        dot.fill.solid(); dot.fill.fore_color.rgb = GRID
        dot.line.fill.background()
        _txt(slide, cx - step / 2, t, step, 0.45, date, size=FONT_PT["table"],
             bold=True, color=SUB_GRAY, align=PP_ALIGN.CENTER)
        _txt(slide, cx - step / 2, axis_y + 0.15, step, 0.55, label,
             size=FONT_PT["bullet1"], bold=True, color=INK, align=PP_ALIGN.CENTER)
    return axis_y + 0.75


def add_insight_box(slide, l, t, w, h, items, tab="주요 시사점"):
    """[시사점 박스] 회색 테두리 박스 + 좌상단 짙은 탭 라벨 (Coles 협업 실측).

    items: [(level, text), ...] — 내부는 표준 불릿 체계.
    """
    _rect(slide, l, t + 0.28, w, h, fill=WHITE, line=GRID, line_w=1.0)
    _rect(slide, l + 0.35, t, 3.0, 0.55, fill=SUB_GRAY)
    _txt(slide, l + 0.35, t, 3.0, 0.55, tab, size=FONT_PT["table"], bold=True,
         color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(slide, items, l=l + 0.55, t=t + 0.75, w=w - 1.1, line_h=0.55)


def add_col_header(slide, l, t, w, text):
    """[본장 2단 컬럼] 컬럼 헤더 바 — #EAEEF6 채움 + 네이비 볼드 가운데 (8·14/15 실측)."""
    _rect(slide, l, t, w, COL_HEADER_H, fill=TH_PRIMARY)
    _txt(slide, l, t, w, COL_HEADER_H, text, size=FONT_PT["bullet0"], bold=True,
         color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return t + COL_HEADER_H


def add_conclusion_box(slide, l, t, w, text, h=CONCL_H):
    """[본장 2단 컬럼] 컬럼 하단 결론 박스 — 네이비 채움 + 백색 볼드 가운데 (So-What)."""
    _rect(slide, l, t, w, h, fill=NAVY)
    _txt(slide, l, t, w, h, text, size=FONT_PT["bullet1"], bold=True,
         color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def add_vsep(slide, x, t, h):
    """[본장 2단 컬럼] 중앙 세로 구분선 (#BFBFBF 0.75pt)."""
    ln = slide.shapes.add_connector(2, Cm(x), Cm(t), Cm(x), Cm(t + h))
    ln.line.color.rgb = GRID
    ln.line.width = Pt(0.75)
    return ln


def add_unit(slide, text="(단위 : 억원, %)"):
    """우상단 단위 표기."""
    l, t, w, h = UNIT
    return _txt(slide, l, t, w, h, text, size=FONT_PT["unit"], bold=True,
                align=PP_ALIGN.RIGHT, color=INK)


def add_footnote(slide, text):
    """좌하단 각주 (※ ...)."""
    l, t, w, h = FOOT
    return _txt(slide, l, t, w, h, text, size=FONT_PT["footnote"], bold=False, color=INK)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])   # 빈 레이아웃


# ─────────────────────────────────────────────────────────────
# 4. 원형(Archetype) 빌더
# ─────────────────────────────────────────────────────────────
def new_deck():
    """A4 가로 빈 프레젠테이션."""
    prs = Presentation()
    prs.slide_width = Emu(9906000)
    prs.slide_height = Emu(6858000)
    return prs


def add_cover(prs, title, org="온라인사업단", date="2026. 07. 06", tag="내부 토의용"):
    """[표지] 상단 로고존 · 태그 · 대제목(밑줄) · 하단 조직/일자."""
    s = _blank(prs)
    _txt(s, 2.0, 6.4, 10.0, 0.7, tag, size=13, bold=True, color=INK)
    # 대제목 (가운데, 밑줄)
    tb = _txt(s, 2.0, 7.3, 23.5, 2.0, title, size=40, bold=True, color=INK,
              align=PP_ALIGN.CENTER)
    # 밑줄
    ln = s.shapes.add_connector(2, Cm(2.2), Cm(9.4), Cm(25.3), Cm(9.4))
    ln.line.color.rgb = INK; ln.line.width = Pt(1.5)
    _txt(s, 2.0, 14.3, 23.5, 0.7, org, size=14, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    _txt(s, 2.0, 15.0, 23.5, 0.7, date, size=13, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    return s


def add_divider(prs, num, title, items):
    """[챕터 간지] 회색 풀블리드 + 대제목 + ①②③ 하위항목."""
    s = _blank(prs)
    _rect(s, 0.0, 3.65, PAGE_W_CM, 13.9, fill=DIVIDER_BG)
    _txt(s, 3.5, 5.6, 20.5, 1.2, f"{num}. {title}", size=28, bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    circ = "①②③④⑤⑥⑦⑧⑨"
    y = 8.0
    for i, it in enumerate(items):
        _txt(s, 9.0, y, 15.0, 0.7, f"{circ[i]} {it}", size=15, bold=True, color=INK)
        y += 0.85
    return s


def add_content(prs, category, title, tier="P",
                lead=None, unit=None, footnote=None):
    """[본문 P/F계열] 크롬만 세팅한 슬라이드 반환 → 본문 자유 배치.

    lead: 제목 하단 `■ ` 리드메시지 한 줄 (선택). 사용 시 본문은 BODY_TOP_LEAD(4.10)부터.
    페이지번호 크롬은 v4.1 에서 제거.
    """
    s = _blank(prs)
    add_header(s, category, title, tier)
    if lead:
        add_lead(s, lead)
    if unit:
        add_unit(s, unit)
    if footnote:
        add_footnote(s, footnote)
    return s


def add_closing(prs, company="롯데마트 · 롯데슈퍼"):
    """[종료] No.1 GROCERY MARKET 마감 장표(간이)."""
    s = _blank(prs)
    _txt(s, 2.0, 2.0, 20.0, 2.2, "No.1", size=72, bold=True, color=CRIMSON)
    _txt(s, 2.0, 4.6, 24.0, 1.6, "GROCERY MARKET", size=44, bold=True, color=CRIMSON)
    _txt(s, 2.05, 7.0, 20.0, 0.8, "Discover a joyful food life", size=18,
         bold=False, color=RGBColor(0x80, 0x80, 0x80))
    cols = [("RE:TARGET", "경영목표"), ("RE:DESIGN", "고객경험"), ("RE:FOCUS", "일하는 방식")]
    xs = [1.5, 10.0, 18.5]
    for (en, ko), x in zip(cols, xs):
        ln = s.shapes.add_connector(2, Cm(x), Cm(13.0), Cm(x + 6.0), Cm(13.0))
        ln.line.color.rgb = INK; ln.line.width = Pt(1.0)
        tb = _txt(s, x, 13.2, 7.0, 0.7, "", size=15, bold=True, color=NAVY)
        p = tb.text_frame.paragraphs[0]
        r1 = p.add_run(); r1.text = en + "  "; _set_font(r1, 15, True, NAVY)
        r2 = p.add_run(); r2.text = ko; _set_font(r2, 10, False, INK)
    return s


# ─────────────────────────────────────────────────────────────
# 5. 재무형 표 (EAEEF6 헤더 · 전 셀 가운데 · 재무 수치열만 우측)
# ─────────────────────────────────────────────────────────────
_TBL_ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}


def add_fin_table(slide, l, t, w, h, data, col_w=None,
                  header_rows=1, header_fill=TH_PRIMARY,
                  col_align=None, font_size=None):
    """
    data: 2D list [ [row0col0, ...], ... ]  (문자열)
    header_rows: 상단 헤더 행 수 (헤더 채움 + 네이비 볼드 + 가운데)

    정렬 표준 (v4.1, 정기협의체 본장 실측): 헤더·본문 전 셀 **가운데** 고정.
    재무 수치열만 col_align 으로 우측 지정 — 예: col_align=["c","r","r"].
    셀별 임의 정렬 금지. 1열(구분)은 볼드.
    """
    rows, cols = len(data), len(data[0])
    gtbl = slide.shapes.add_table(rows, cols, Cm(l), Cm(t), Cm(w), Cm(h))
    tbl = gtbl.table
    # 기본 스타일 제거(밴딩 off)
    tbl.first_row = False; tbl.horz_banding = False
    if col_w:
        for j, cw in enumerate(col_w):
            tbl.columns[j].width = Cm(cw)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Cm(0.1); cell.margin_right = Cm(0.1)
            cell.margin_top = 0; cell.margin_bottom = 0
            _cell_border(cell)                       # 격자 #BFBFBF 0.75pt (v4.1 코드 강제)
            is_head = i < header_rows
            if is_head:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
            tf = cell.text_frame; tf.word_wrap = True
            if is_head:
                align = PP_ALIGN.CENTER
            else:
                align = _TBL_ALIGN[col_align[j] if col_align else "c"]
            # 다중행 셀도 전 행 동일 정렬 — 줄마다 문단 분리
            for k, line in enumerate(str(val).split("\n")):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = align
                r = p.add_run(); r.text = line
                _set_font(r, font_size or FONT_PT["table"], bold=is_head or j == 0,
                          color=NAVY if is_head else INK)
    return tbl
