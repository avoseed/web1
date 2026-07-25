# -*- coding: utf-8 -*-
"""
frames.py — 완성 슬라이드 템플릿 레이어 (호출 1번 = 장표 1장)
================================================================
설계 의도
  기존 `zetta_ppt_standard.py` 는 프리미티브(표·타임라인·맵…) 조립 방식이라
  매번 좌표를 손으로 계산해야 했고, 그 탓에 여백·하단 충전·정렬 규칙이 반복 재발했다.
  이 모듈은 **규칙을 코드로 강제**한다 — 함수 하나가 크롬(제목·밑줄·리드·각주·페이지번호)과
  1.5cm 액자, 하단 충전(CONTENT_BOTTOM)까지 자동 처리한 완성 장표를 만든다.

원형 출처
  F계열 = `docs/BODY_FRAMES.md` 실측 카탈로그(임원 보고 덱 원본 PPTX 측정).
  일부 원형은 공개 레퍼런스(mckinsey-pptx)의 '완성 슬라이드 함수' API 설계를 참고.

공통 규약
  - 모든 함수 시그니처: f_xxx(prs, title, lead, ..., footnote=None, page=None) -> slide
  - 좌표 cm. 본문은 BODY_L(1.50) ~ BODY_L+BODY_W(26.02), BODY_TOP_F(4.20) ~ CONTENT_BOTTOM(17.55).
  - 각주가 있으면 콘텐츠 하한이 각주 높이만큼 올라간다(각주는 항상 하단선에 붙음).
  - 결론은 리드에 선치(하단 결론 박스 없음), 파란 글자 금지(강조 셀 예외).
"""
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from zetta_ppt_standard import (
    new_deck, _blank, _txt, _rect, _line, _cell_border, _set_font,
    add_fin_table, add_matrix2x2, add_htimeline,
    PAGE_W_CM, PAGE_H_CM, FONT, MARGIN_L, MARGIN_R, MARGIN_B, BODY_W, CONTENT_BOTTOM,
    INK, BLACK, WHITE, GRID, SUB_GRAY, TH_PRIMARY, HL_FILL, HL_TEXT, CRIMSON,
)

# ── 팔레트 (BODY_FRAMES §0-A 실측) ───────────────────────────
NAVY_F   = RGBColor(0x20, 0x38, 0x64)   # 라벨·카드 헤더·표 헤더 주 채움
NAVY_M   = RGBColor(0x2E, 0x75, 0xB6)   # 명도 2단계
NAVY_L   = RGBColor(0x9D, 0xC3, 0xE6)   # 명도 3단계
RED_F    = RGBColor(0xC0, 0x00, 0x00)   # 강조 테두리·콜아웃·핵심 수치
SOFT     = RGBColor(0xF5, 0xF7, 0xFA)   # 연회색 패널
LIGHTBLU = RGBColor(0xEA, 0xF2, 0xFA)   # 연블루 배경
GRAY_L   = RGBColor(0xB7, 0xBD, 0xC6)   # 격자·보조선
REDBG    = RGBColor(0xFD, 0xF3, 0xF3)   # 강조 열 연빨강 채움

# ── 크롬 좌표 (실측 → 1.5cm 액자 환산) ───────────────────────
BODY_L      = MARGIN_L                  # 1.50
BODY_R      = PAGE_W_CM - MARGIN_R      # 26.02
TITLE_T     = 0.82
TITLE_H     = 1.05
ULINE_T     = 1.96
LEAD_T      = 3.02
LEAD_H      = 0.82
BODY_TOP_F  = 4.20                      # 본문 시작
FOOT_H      = 0.52
GAP         = 0.34                      # 블록 간 표준 간격
TL_H        = 2.48                      # 【블록헤더】+가로타임라인 총 높이(예약분)

PT = {"title": 20, "lead": 14, "sec": 12, "body": 11, "small": 10, "foot": 8.5}


# ═══════════════════════════════════════════════════════════
# 크롬
# ═══════════════════════════════════════════════════════════
def _text_w(s, pt):
    """제목 밑줄 폭 추정 — 한글 1em, 라틴/숫자 0.52em."""
    em = pt * 0.03528
    return sum(em * (1.0 if ord(c) > 0x1100 else 0.52) for c in s)


def chrome(prs, title, lead=None, footnote=None, page=None, title_align="center"):
    """제목(가운데+텍스트 폭 밑줄) · ■ 리드 · 각주 · 페이지번호를 배치.
    returns (slide, body_top, body_bottom) — 본문 가용 영역."""
    s = _blank(prs)
    align = PP_ALIGN.CENTER if title_align == "center" else PP_ALIGN.LEFT
    _txt(s, BODY_L, TITLE_T, BODY_W, TITLE_H, title, size=PT["title"],
         bold=True, color=BLACK, align=align, anchor=MSO_ANCHOR.MIDDLE)
    uw = min(_text_w(title, PT["title"]) + 0.5, BODY_W)
    ux = (PAGE_W_CM - uw) / 2 if title_align == "center" else BODY_L
    _line(s, ux, ULINE_T, ux + uw, ULINE_T, color=BLACK, w_pt=2.0)

    top = BODY_TOP_F
    if lead:
        _txt(s, BODY_L, LEAD_T, BODY_W, LEAD_H, "■ " + lead, size=PT["lead"],
             bold=True, color=BLACK, anchor=MSO_ANCHOR.MIDDLE)
    else:
        top = LEAD_T

    bottom = CONTENT_BOTTOM
    if footnote:
        _txt(s, BODY_L, CONTENT_BOTTOM - FOOT_H, BODY_W, FOOT_H, footnote,
             size=PT["foot"], color=SUB_GRAY, anchor=MSO_ANCHOR.BOTTOM)
        bottom = CONTENT_BOTTOM - FOOT_H - 0.16
    if page:
        _txt(s, BODY_R - 3.0, CONTENT_BOTTOM - FOOT_H, 3.0, FOOT_H, page,
             size=PT["foot"], color=SUB_GRAY, align=PP_ALIGN.RIGHT,
             anchor=MSO_ANCHOR.BOTTOM)
    return s, top, bottom


def _sec(s, l, t, w, text):
    """블록 소제목 — 반환: 다음 y"""
    _txt(s, l, t, w, 0.52, text, size=PT["sec"], bold=True, color=INK)
    return t + 0.62


def _blk(s, l, t, w, text):
    """【 】 블록 헤더 — 반환: 다음 y"""
    _txt(s, l, t, w, 0.52, "【 %s 】" % text, size=PT["sec"], bold=True, color=INK)
    return t + 0.62


def _bullets(s, l, t, w, items, lh=0.50, sub_lh=0.44):
    """위계 불릿: (0,'포괄') / (1,'세부'). 반환: 종료 y"""
    y = t
    for lv, tx in items:
        if lv == 0:
            _txt(s, l, y, w, lh, "• " + tx, size=PT["body"], bold=True, color=INK)
            y += lh
        else:
            _txt(s, l + 0.42, y, w - 0.42, sub_lh, "– " + tx,
                 size=PT["small"], color=INK)
            y += sub_lh
    return y


def _numbox(s, x, y, n, size=0.50, fill=None):
    """검정/네이비 번호 사각 박스"""
    _rect(s, x, y, size, size, fill=fill or BLACK)
    _txt(s, x, y, size, size, str(n), size=PT["small"], bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def _chevron(s, cx, cy, w=0.52, h=0.46, color=None):
    """회색 쉐브론(전개 방향)"""
    sp = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Cm(cx), Cm(cy), Cm(w), Cm(h))
    sp.rotation = 90
    sp.shadow.inherit = False
    sp.fill.solid(); sp.fill.fore_color.rgb = color or GRAY_L
    sp.line.fill.background()
    return sp


# ═══════════════════════════════════════════════════════════
# F계열 — 실측 카탈로그 원형
# ═══════════════════════════════════════════════════════════
def f_exec_summary(prs, title, lead, points, footnote=None, page=None):
    """[요약] ■ 항목 3~5개를 넉넉한 행간으로 — 덱 서두."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(points)
    h = (bot - top) / n
    for i, p in enumerate(points):
        y = top + h * i
        _rect(s, BODY_L, y + h * 0.12, 0.09, h * 0.66, fill=NAVY_F)
        if isinstance(p, (tuple, list)):
            head, sub = p
            _txt(s, BODY_L + 0.42, y + h * 0.10, BODY_W - 0.42, 0.60, head,
                 size=13, bold=True, color=INK)
            _txt(s, BODY_L + 0.42, y + h * 0.10 + 0.62, BODY_W - 0.42, h * 0.5, sub,
                 size=PT["body"], color=INK)
        else:
            _txt(s, BODY_L + 0.42, y + h * 0.10, BODY_W - 0.42, h * 0.7, p,
                 size=13, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    return s


def f_spec_overview(prs, title, lead, spec_rows, ops_title=None, ops_items=None,
                    bg_bullets=None, timeline=None, implication=None,
                    footnote=None, page=None, spec_title="① 도입 개요",
                    hi_rows=None):
    """[F4 스펙 개요] 현황 보고의 정본.
    좌 = 제원표(항목|값), 우 = 운영 기반 패널. 상단 배경 불릿·하단 경과 타임라인 선택.

    spec_rows : [(항목, 값), ...]          hi_rows : 강조할 행 인덱스
    ops_items : [(제목, 설명), ...]        bg_bullets : [(lv, text), ...]
    timeline  : [(시점, 내용), ...]        implication : 시사점 한 줄
    """
    s, top, bot = chrome(prs, title, lead, footnote, page)
    y = top
    if bg_bullets:
        y = _bullets(s, BODY_L, y, BODY_W, bg_bullets) + GAP

    # 하단 예약(타임라인·시사점)
    rsv = 0.0
    if timeline:
        rsv += TL_H
    if implication:
        rsv += 0.90
    main_bot = bot - rsv - (GAP if rsv else 0)

    lw = BODY_W * 0.56
    rw = BODY_W - lw - 0.60
    rx = BODY_L + lw + 0.60

    ty = _sec(s, BODY_L, y, lw, spec_title)
    data = [[k, v] for k, v in spec_rows]
    add_fin_table(s, BODY_L, ty, lw, main_bot - ty, data,
                  col_w=[lw * 0.30, lw * 0.70], header_rows=0,
                  col_align=["c", "l"], bold_cols=(0,),
                  hl_rows=hi_rows, font_size=PT["small"])

    if ops_items:
        oy = _sec(s, rx, y, rw, ops_title or "② 운영 기반")
        _rect(s, rx, oy, rw, main_bot - oy, fill=SOFT, line=GRAY_L)
        n = len(ops_items)
        ih = (main_bot - oy) / n
        for i, (head, sub) in enumerate(ops_items):
            iy = oy + ih * i + 0.16
            _numbox(s, rx + 0.26, iy + 0.02, i + 1, 0.44, fill=NAVY_F)
            _txt(s, rx + 0.84, iy, rw - 1.1, 0.46, head, size=PT["small"],
                 bold=True, color=INK)
            _txt(s, rx + 0.84, iy + 0.48, rw - 1.1, ih - 0.62, sub,
                 size=9.5, color=SUB_GRAY)

    y = main_bot + GAP
    if timeline:
        y = _blk(s, BODY_L, y, BODY_W, "확대 경과")
        add_htimeline(s, BODY_L, y, BODY_W, timeline, when_h=0.44,
                      content_h=1.05, size=PT["small"])
        y += 1.75
    if implication:
        _rect(s, BODY_L, y, 0.09, 0.72, fill=NAVY_F)
        _txt(s, BODY_L + 0.34, y, BODY_W - 0.34, 0.72,
             "시사점 · " + implication, size=PT["small"], color=INK,
             anchor=MSO_ANCHOR.MIDDLE)
    return s


def f_compare_table(prs, title, lead, header, rows, hero_col=1,
                    quote=None, timeline=None, bg_bullets=None,
                    footnote=None, page=None, sec=None):
    """[F2 강조 열 비교표] 동향·비교형의 정본.
    header : [구분, A, B, C]   rows : [(항목, [값...]), ...]
    hero_col : 값 열 중 강조 인덱스(0-base) — 빨강 테두리 + 연빨강 채움.
    """
    s, top, bot = chrome(prs, title, lead, footnote, page)
    y = top
    if bg_bullets:
        y = _bullets(s, BODY_L, y, BODY_W, bg_bullets) + GAP

    rsv = (TL_H if timeline else 0) + (0.70 if quote else 0)
    tbl_bot = bot - rsv - (GAP if rsv else 0)

    if sec:
        y = _sec(s, BODY_L, y, BODY_W, sec)
    ncol = len(header)
    kw = BODY_W * 0.17
    vw = (BODY_W - kw) / (ncol - 1)
    data = [list(header)] + [[k] + list(v) for k, v in rows]
    add_fin_table(s, BODY_L, y, BODY_W, tbl_bot - y, data,
                  col_w=[kw] + [vw] * (ncol - 1), header_rows=1,
                  hl_cols=[hero_col + 1], bold_cols=(0,), font_size=PT["small"])
    # 강조 열 빨강 테두리
    hx = BODY_L + kw + vw * hero_col
    for xx in (hx, hx + vw):
        _line(s, xx, y, xx, tbl_bot, color=RED_F, w_pt=1.75)
    _line(s, hx, y, hx + vw, y, color=RED_F, w_pt=1.75)
    _line(s, hx, tbl_bot, hx + vw, tbl_bot, color=RED_F, w_pt=1.75)

    y = tbl_bot + GAP
    if quote:
        _txt(s, BODY_L, y, BODY_W, 0.62, "“%s”" % quote, size=PT["sec"],
             bold=True, color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.70
    if timeline:
        y = _blk(s, BODY_L, y, BODY_W, "추진 경과")
        add_htimeline(s, BODY_L, y, BODY_W, timeline, when_h=0.44,
                      content_h=1.05, size=PT["small"])
    return s


def f_two_col(prs, title, lead, left_head, left_body, right_head, right_body,
              footnote=None, page=None):
    """[F6/F9 좌우 2단 근거] 각 컬럼 = 헤더 바 + 위계 불릿.
    left_body/right_body : [(lv, text), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    cw = (BODY_W - 0.70) / 2
    rx = BODY_L + cw + 0.70
    for x, head, body in ((BODY_L, left_head, left_body), (rx, right_head, right_body)):
        _rect(s, x, top, cw, 0.72, fill=TH_PRIMARY, line=GRAY_L)
        _txt(s, x, top, cw, 0.72, head, size=PT["sec"], bold=True, color=INK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _bullets(s, x + 0.20, top + 0.94, cw - 0.40, body)
    _line(s, BODY_L + cw + 0.35, top, BODY_L + cw + 0.35, bot, color=GRAY_L, w_pt=0.75)
    return s


def f_cards(prs, title, lead, cards, footnote=None, page=None, chevron=False):
    """[병렬 카드 3~4] 네이비 헤더 + 흰 본문. cards : [(헤더, [줄...]), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(cards)
    gap = 0.55
    cw = (BODY_W - gap * (n - 1)) / n
    for i, (head, lines) in enumerate(cards):
        x = BODY_L + (cw + gap) * i
        _rect(s, x, top, cw, 0.78, fill=NAVY_F)
        _txt(s, x, top, cw, 0.78, head, size=PT["sec"], bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _rect(s, x, top + 0.78, cw, bot - top - 0.78, fill=WHITE, line=GRAY_L)
        y = top + 1.02
        for ln in lines:
            if ln.startswith("-"):
                _txt(s, x + 0.55, y, cw - 0.80, 0.46, ln, size=9.5, color=SUB_GRAY)
                y += 0.46
            else:
                _txt(s, x + 0.30, y, cw - 0.55, 0.50, "● " + ln, size=PT["small"],
                     bold=True, color=INK)
                y += 0.52
        if chevron and i < n - 1:
            _chevron(s, x + cw + 0.02, top + (bot - top) / 2 - 0.23)
    return s


def f_process(prs, title, lead, steps, note=None, detail=None,
              footnote=None, page=None):
    """[E1 프로세스·게이트] 노드 → 화살표. steps : [(제목, 설명), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(steps)
    aw = 0.72
    nw = (BODY_W - aw * (n - 1)) / n
    nh = min(2.30, (bot - top) * 0.42)
    for i, (head, desc) in enumerate(steps):
        x = BODY_L + (nw + aw) * i
        _rect(s, x, top, nw, nh, fill=WHITE, line=GRAY_L)
        _rect(s, x, top, nw, 0.10, fill=NAVY_F)
        _txt(s, x + 0.20, top + 0.30, nw - 0.40, 0.52, head, size=PT["sec"],
             bold=True, color=NAVY_F, align=PP_ALIGN.CENTER)
        _txt(s, x + 0.20, top + 0.90, nw - 0.40, nh - 1.05, desc, size=9.5,
             color=INK, align=PP_ALIGN.CENTER)
        if i < n - 1:
            _txt(s, x + nw, top + nh / 2 - 0.32, aw, 0.64, "→", size=16,
                 bold=True, color=SUB_GRAY, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
    y = top + nh + GAP
    if note:
        _txt(s, BODY_L, y, BODY_W, 0.48, "※ " + note, size=9.5, color=SUB_GRAY)
        y += 0.58
    if detail:
        _bullets(s, BODY_L, y, BODY_W, detail)
    return s


def f_timeline(prs, title, lead, milestones, blocks=None, footnote=None, page=None):
    """[B2 타임라인 스파인] 마일스톤 중심 축 + 하단 분석 블록(선택)."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    th = 2.60 if blocks else (bot - top) * 0.55
    add_htimeline(s, BODY_L, top + 0.30, BODY_W, milestones,
                  when_h=0.48, content_h=1.25, size=PT["small"])
    if blocks:
        y = top + th + GAP
        n = len(blocks)
        gap = 0.55
        cw = (BODY_W - gap * (n - 1)) / n
        for i, (head, items) in enumerate(blocks):
            x = BODY_L + (cw + gap) * i
            _txt(s, x, y, cw, 0.52, head, size=PT["sec"], bold=True, color=INK)
            _line(s, x, y + 0.54, x + cw, y + 0.54, color=CRIMSON, w_pt=1.5)
            _bullets(s, x, y + 0.70, cw, items)
    return s


def f_gantt(prs, title, lead, periods, tracks, note=None, footnote=None, page=None):
    """[B3 멀티트랙 로드맵] periods : ['1Q','2Q',...]
    tracks : [(워크스트림, start_idx, span, 설명), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    lw = BODY_W * 0.24
    lane_x = BODY_L + lw
    lane_w = BODY_W - lw
    pw = lane_w / len(periods)
    for i, p in enumerate(periods):
        _txt(s, lane_x + pw * i, top, pw, 0.50, p, size=PT["small"], bold=True,
             color=INK, align=PP_ALIGN.CENTER)
    _line(s, lane_x, top + 0.54, lane_x + lane_w, top + 0.54, color=GRAY_L)
    y0 = top + 0.72
    n = len(tracks)
    rh = min(1.35, (bot - y0 - (0.60 if note else 0)) / n)
    shades = [NAVY_F, NAVY_M, NAVY_L]
    for i, (name, st, span, desc) in enumerate(tracks):
        y = y0 + rh * i
        _txt(s, BODY_L, y, lw - 0.25, rh * 0.55, name, size=PT["small"],
             bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        _rect(s, lane_x, y + rh * 0.14, lane_w, rh * 0.44, fill=SOFT)
        _rect(s, lane_x + pw * st, y + rh * 0.14, pw * span, rh * 0.44,
              fill=shades[i % 3])
        if desc:
            _txt(s, BODY_L, y + rh * 0.60, BODY_W, rh * 0.36, "– " + desc,
                 size=9.5, color=SUB_GRAY)
    if note:
        _txt(s, BODY_L, bot - 0.52, BODY_W, 0.48, "※ " + note, size=9.5, color=SUB_GRAY)
    return s


def f_waterfall(prs, title, lead, base, steps, final, groups=None,
                unit=None, footnote=None, page=None):
    """[F12 워터폴 + 드라이버] 목표 기여 분해 — 덱 종착 장표.
    base : (라벨, 값)   steps : [(라벨, 증분), ...]   final : (라벨, 값)
    groups : [(그룹명, [세부, 값], ...)] 상단 미니표(선택)
    """
    s, top, bot = chrome(prs, title, lead, footnote, page)
    y = top
    if unit:                                   # 단위는 본문 우상단(그룹표 위)
        _txt(s, BODY_R - 6.0, y - 0.46, 6.0, 0.42, unit, size=9.5,
             color=SUB_GRAY, align=PP_ALIGN.RIGHT)
    if groups:
        n = len(groups)
        gap = 0.55
        gw = (BODY_W - gap * (n - 1)) / n
        gh = 1.70
        for i, (gname, items) in enumerate(groups):
            x = BODY_L + (gw + gap) * i
            _numbox(s, x, y, i + 1, 0.46, fill=BLACK)
            _txt(s, x + 0.58, y, gw - 0.58, 0.46, gname, size=PT["small"],
                 bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
            data = [[k, v] for k, v in items]
            add_fin_table(s, x, y + 0.56, gw, gh - 0.56, data,
                          col_w=[gw * 0.58, gw * 0.42], header_rows=0,
                          col_align=["l", "c"], font_size=9.5)
        y += gh + GAP

    ch_bot = bot - 0.75
    ch_top = y
    ch_h = ch_bot - ch_top
    cats = [base] + steps + [final]
    n = len(cats)
    cw = BODY_W / n
    vals = [base[1]] + [v for _, v in steps] + [final[1]]
    peak = max(base[1] + sum(v for _, v in steps), final[1])
    scale = ch_h / (peak * 1.12) if peak else 1
    _line(s, BODY_L, ch_bot, BODY_R, ch_bot, color=INK, w_pt=1.0)
    run = base[1]
    for i, (lab, val) in enumerate(cats):
        x = BODY_L + cw * i + cw * 0.18
        bw = cw * 0.64
        if i == 0:
            h = base[1] * scale
            _rect(s, x, ch_bot - h, bw, h, fill=GRAY_L)
            top_y, show = ch_bot - h, "{:,}".format(base[1])
        elif i == n - 1:
            h = final[1] * scale
            _rect(s, x, ch_bot - h, bw, h, fill=BLACK)
            top_y, show = ch_bot - h, "{:,}".format(final[1])
        else:
            h = abs(val) * scale
            y0 = ch_bot - (run + max(val, 0)) * scale
            _rect(s, x, y0, bw, h, fill=NAVY_M if val >= 0 else RED_F)
            run += val
            top_y, show = y0, ("+{:,}".format(val) if val >= 0
                               else "▲{:,}".format(abs(val)))
        _txt(s, x - cw * 0.15, top_y - 0.52, bw + cw * 0.30, 0.48, show,
             size=PT["small"], bold=True, color=INK, align=PP_ALIGN.CENTER)
        _txt(s, BODY_L + cw * i, ch_bot + 0.10, cw, 0.60, lab, size=9.5,
             color=INK, align=PP_ALIGN.CENTER)
    return s


def f_funnel(prs, title, lead, levels, note=None, footnote=None, page=None):
    """[D5 퍼널/TAM-SAM-SOM] levels : [(라벨, 값문자열), ...] 위→아래 축소."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(levels)
    avail = bot - top - (0.60 if note else 0)
    lh = min(1.55, avail / n * 0.82)
    gap = (avail - lh * n) / max(n - 1, 1)
    shades = [NAVY_L, NAVY_M, NAVY_F, NAVY_F]
    for i, (lab, val) in enumerate(levels):
        wpct = 1.0 - i * (0.52 / max(n - 1, 1))
        w = BODY_W * wpct
        x = BODY_L + (BODY_W - w) / 2
        yy = top + (lh + gap) * i
        _rect(s, x, yy, w, lh, fill=shades[min(i, 3)])
        _txt(s, x, yy, w, lh, "%s   %s" % (lab, val), size=PT["sec"], bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if note:
        _txt(s, BODY_L, bot - 0.52, BODY_W, 0.48, "※ " + note, size=9.5, color=SUB_GRAY)
    return s


def f_pyramid(prs, title, lead, tiers, note=None, footnote=None, page=None):
    """[D1 피라미드] tiers : [(라벨, 설명), ...] 상위=소수·하위=다수."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(tiers)
    avail = bot - top - (0.60 if note else 0)
    th = avail / n * 0.80
    gap = (avail - th * n) / max(n - 1, 1)
    shades = [NAVY_F, NAVY_M, NAVY_L, NAVY_L]
    for i, (lab, desc) in enumerate(tiers):
        wpct = 0.42 + i * (0.58 / max(n - 1, 1))
        w = BODY_W * wpct
        x = BODY_L + (BODY_W - w) / 2
        yy = top + (th + gap) * i
        _rect(s, x, yy, w, th, fill=shades[min(i, 3)])
        _txt(s, x, yy, w, th * 0.55, lab, size=PT["sec"], bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
        if desc:
            _txt(s, x, yy + th * 0.55, w, th * 0.42, desc, size=9.5, color=WHITE,
                 align=PP_ALIGN.CENTER)
    if note:
        _txt(s, BODY_L, bot - 0.52, BODY_W, 0.48, "※ " + note, size=9.5, color=SUB_GRAY)
    return s


def f_matrix(prs, title, lead, x_axis, y_axis, cells, star=None, dim=None,
             note=None, footnote=None, page=None):
    """[C1 2×2 매트릭스] cells : [[좌상, 우상], [좌하, 우하]]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    h = bot - top - (0.62 if note else 0)      # 하단까지 충전(§3-4-0)
    w = min(BODY_W, h * 1.85)
    x = BODY_L + (BODY_W - w) / 2
    add_matrix2x2(s, x, top, w, h, x_axis, y_axis, cells, star=star, dim=dim)
    if note:
        _txt(s, BODY_L, top + h + 0.14, BODY_W, 0.48, "※ " + note,
             size=9.5, color=SUB_GRAY)
    return s


def f_before_after(prs, title, lead, before, after, bg_bullets=None,
                   steps=None, footnote=None, page=None,
                   before_label="As-Is · 현재", after_label="To-Be · 전환 후"):
    """[A3 As-Is → To-Be] before/after : [(항목, 값), ...] 또는 [줄, ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    y = top
    if bg_bullets:
        y = _bullets(s, BODY_L, y, BODY_W, bg_bullets) + GAP
    rsv = (TL_H - 0.10) if steps else 0
    box_bot = bot - rsv - (GAP if rsv else 0)
    aw = 1.15
    cw = (BODY_W - aw) / 2
    for x, lab, body, accent in ((BODY_L, before_label, before, INK),
                                 (BODY_L + cw + aw, after_label, after, RED_F)):
        _txt(s, x, y, cw, 0.58, lab, size=PT["sec"], bold=True, color=accent,
             align=PP_ALIGN.CENTER)
        _line(s, x, y + 0.62, x + cw, y + 0.62, color=accent, w_pt=1.75)
        _rect(s, x, y + 0.72, cw, box_bot - y - 0.72, fill=WHITE, line=GRAY_L)
        yy = y + 0.98
        for it in body:
            if isinstance(it, (tuple, list)):
                k, v = it
                _txt(s, x + 0.26, yy, cw * 0.30, 0.50, k, size=9.5, bold=True,
                     color=SUB_GRAY)
                _txt(s, x + 0.26 + cw * 0.30, yy, cw * 0.64, 0.50, v,
                     size=PT["small"], color=INK)
            else:
                _txt(s, x + 0.26, yy, cw - 0.50, 0.50, "• " + it,
                     size=PT["small"], color=INK)
            yy += 0.56
    _txt(s, BODY_L + cw, y + (box_bot - y) / 2 - 0.35, aw, 0.70, "▶",
         size=17, bold=True, color=GRAY_L, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    if steps:
        yy = box_bot + GAP
        yy = _blk(s, BODY_L, yy, BODY_W, "실행 단계")
        add_htimeline(s, BODY_L, yy, BODY_W, steps, when_h=0.44,
                      content_h=0.95, size=PT["small"])
    return s


def f_assessment(prs, title, lead, criteria, options, best=None,
                 legend="● 충족  ◐ 부분  ○ 미흡", footnote=None, page=None):
    """[A5 평가 매트릭스(하비볼)] options : [(옵션명, [기호...], 종합), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    header = ["구분"] + list(criteria) + ["종합"]
    data = [header] + [[nm] + list(marks) + [tot] for nm, marks, tot in options]
    kw = BODY_W * 0.20
    rest = (BODY_W - kw) / (len(header) - 1)
    hl = [best + 1] if best is not None else None
    add_fin_table(s, BODY_L, top, BODY_W, bot - top - 0.60, data,
                  col_w=[kw] + [rest] * (len(header) - 1), header_rows=1,
                  hl_rows=hl, bold_cols=(0,), font_size=PT["small"])
    _txt(s, BODY_L, bot - 0.52, BODY_W, 0.48, "※ " + legend, size=9.5, color=SUB_GRAY)
    return s


def f_kpi_tiles(prs, title, lead, tiles, body_bullets=None, footnote=None, page=None):
    """[KPI 타일] tiles : [(라벨, 값, 부가), ...] — 실적 대시보드 전용(남용 금지)."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    n = len(tiles)
    gap = 0.50
    tw = (BODY_W - gap * (n - 1)) / n
    th = 2.35
    for i, t in enumerate(tiles):
        lab, val = t[0], t[1]
        sub = t[2] if len(t) > 2 else None
        x = BODY_L + (tw + gap) * i
        _rect(s, x, top, tw, th, fill=WHITE, line=GRAY_L)
        _rect(s, x, top, tw, 0.09, fill=NAVY_F)
        _txt(s, x, top + 0.30, tw, 0.46, lab, size=PT["small"], color=SUB_GRAY,
             align=PP_ALIGN.CENTER)
        _txt(s, x, top + 0.80, tw, 0.92, val, size=19, bold=True, color=NAVY_F,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            _txt(s, x, top + 1.76, tw, 0.46, sub, size=9.5, color=SUB_GRAY,
                 align=PP_ALIGN.CENTER)
    if body_bullets:
        _bullets(s, BODY_L, top + th + GAP, BODY_W, body_bullets)
    return s


def f_issue_tree(prs, title, lead, root, branches, note=None, footnote=None, page=None):
    """[D2 이슈/로직 트리] branches : [(가지, [세부...]), ...]"""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    avail = bot - top - (0.60 if note else 0)
    rw = BODY_W * 0.24
    bw = BODY_W * 0.30
    lw = BODY_W - rw - bw - 1.20
    _rect(s, BODY_L, top + avail * 0.32, rw, avail * 0.36, fill=NAVY_F)
    _txt(s, BODY_L, top + avail * 0.32, rw, avail * 0.36, root, size=PT["sec"],
         bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    n = len(branches)
    bh = avail / n
    bx = BODY_L + rw + 0.60
    for i, (br, subs) in enumerate(branches):
        y = top + bh * i + bh * 0.10
        _rect(s, bx, y, bw, bh * 0.80, fill=WHITE, line=NAVY_F)
        _txt(s, bx + 0.16, y, bw - 0.32, bh * 0.80, br, size=PT["small"],
             bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        _line(s, BODY_L + rw, top + avail * 0.5, bx, y + bh * 0.40, color=GRAY_L)
        sy = y
        for sub in subs:
            _txt(s, bx + bw + 0.60, sy, lw, 0.48, "– " + sub, size=9.5, color=INK)
            sy += 0.50
        _line(s, bx + bw, y + bh * 0.40, bx + bw + 0.45, y + bh * 0.40, color=GRAY_L)
    if note:
        _txt(s, BODY_L, bot - 0.52, BODY_W, 0.48, "※ " + note, size=9.5, color=SUB_GRAY)
    return s


def f_bigstat(prs, title, lead, stat, caption=None, support=None,
              footnote=None, page=None):
    """[대형 숫자 강조] 단일 수치로 메시지를 못박는 장표."""
    s, top, bot = chrome(prs, title, lead, footnote, page)
    h = bot - top
    _txt(s, BODY_L, top + h * 0.16, BODY_W, h * 0.42, stat, size=54, bold=True,
         color=NAVY_F, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if caption:
        _txt(s, BODY_L, top + h * 0.60, BODY_W, 0.80, caption, size=PT["lead"],
             bold=True, color=INK, align=PP_ALIGN.CENTER)
    if support:
        _bullets(s, BODY_L + BODY_W * 0.16, top + h * 0.74, BODY_W * 0.68, support)
    return s


# 카탈로그 — 에이전트가 프레임을 고를 때 참조
CATALOG = {
    "f_exec_summary":  "요약 — ■ 항목 3~5개(덱 서두·결론 요약)",
    "f_spec_overview": "F4 현황 개요 — 제원표 + 운영 기반(현황 보고 정본)",
    "f_compare_table": "F2 강조 열 비교표 + 인용 + 경과(동향·비교형 정본)",
    "f_two_col":       "F6/F9 좌우 2단 근거 — 두 갈래 병렬 논증",
    "f_cards":         "병렬 카드 3~4 — 독립 축·기둥",
    "f_process":       "E1 프로세스·게이트 — 흐름·단계",
    "f_timeline":      "B2 타임라인 스파인 — 경과·마일스톤",
    "f_gantt":         "B3 멀티트랙 로드맵 — 과제 × 기간",
    "f_waterfall":     "F12 워터폴 + 드라이버 — 목표 기여 분해(종착 장표)",
    "f_funnel":        "D5 퍼널/TAM-SAM-SOM — 단계 축소",
    "f_pyramid":       "D1 피라미드 — 계층·전략 체계",
    "f_matrix":        "C1 2×2 매트릭스 — 위치·우선순위",
    "f_before_after":  "A3 As-Is → To-Be — 상태 전환",
    "f_assessment":    "A5 평가 매트릭스(하비볼) — 옵션 × 기준",
    "f_kpi_tiles":     "KPI 타일 — 실적 대시보드 전용(남용 금지)",
    "f_issue_tree":    "D2 이슈/로직 트리 — MECE 분해",
    "f_bigstat":       "대형 숫자 강조 — 단일 수치 메시지",
}
