# -*- coding: utf-8 -*-
"""롯데마트 캄보디아 확장 기획안(18매) — 상세 레이아웃 명세(유형 A~E) 정밀 구현.

공통 그리드: 좌우 여백 1.2(폭 25.12) · 헤더 y0.7~1.6 · 리딩 밴드 y1.8~2.9(#F2F4F8)
· 본문 y3.3~16.9(70%+ 충전) · 각주 좌하단 8pt · 페이지번호 우하단 n/N.
제목: P계열(표지·1요약·14제언) 검정 볼드 / F계열 네이비 20pt.
제약: ASK/그룹전략/교육용/베트남손익 금지 · 검증 수치만 · ※검증 필요 유지 · 개조식 명사 종결.
출처: Worldometer/UN · IMF · 6Wresearch · US Commerce(trade.gov) · Khmer Times/MoC
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, _rect, _set_font, _cell_border

# ── 팔레트(명세 지정) ──
BAND = RGBColor(0xF2, 0xF4, 0xF8)
CARD_BD = RGBColor(0xD9, 0xDE, 0xE8)
STRIPE = RGBColor(0xF7, 0xF8, 0xFB)
HL_ROW = RGBColor(0xFD, 0xF2, 0xF5)
G6 = RGBColor(0x66, 0x66, 0x66)
G8 = RGBColor(0x88, 0x88, 0x88)
NAVY30 = RGBColor(179, 179, 209)
NAVY60 = RGBColor(102, 102, 163)
CW = 25.12                                     # 콘텐츠 폭
X0, X1 = 1.2, 26.32

N_BODY = 17
_pg = [0]


def header(s, cat, title, tier="F"):
    color = INK if tier == "P" else NAVY
    size = 28 if tier == "P" else 20
    tb = s.shapes.add_textbox(Cm(X0), Cm(0.55), Cm(CW), Cm(1.15))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r1 = p.add_run(); r1.text = f"[{cat}] "; _set_font(r1, size, True, color)
    r2 = p.add_run(); r2.text = title; _set_font(r2, size, True, color)


def band(s, text):
    _rect(s, X0, 1.8, CW, 1.1, fill=BAND)
    _txt(s, X0 + 0.4, 1.8, CW - 0.8, 1.1, text, size=14, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)


def footer(s, note=None):
    _pg[0] += 1
    if note:
        _txt(s, X0, 17.3, 18.5, 0.9, note, size=8, color=G8)
    _txt(s, 22.0, 17.35, 4.32, 0.6, f"{_pg[0]} / {N_BODY}", size=9, color=G6,
         align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, label, body, lcolor=NAVY):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.shadow.inherit = False
    sp.adjustments[0] = 0.05
    sp.fill.solid(); sp.fill.fore_color.rgb = WHITE
    sp.line.color.rgb = CARD_BD; sp.line.width = Pt(1.0)
    _txt(s, x + 0.4, y + 0.35, w - 0.8, 0.7, label, size=12, bold=True, color=lcolor)
    _txt(s, x + 0.4, y + 1.15, w - 0.8, h - 1.5, body, size=12, color=INK,
         anchor=MSO_ANCHOR.MIDDLE)


def strip(s, x, y, w, text, fill=BAND, tcolor=NAVY, size=13, bold=True,
          align=PP_ALIGN.CENTER, h=1.1):
    _rect(s, x, y, w, h, fill=fill)
    _txt(s, x + 0.4, y, w - 0.8, h, text, size=size, bold=bold, color=tcolor,
         align=align, anchor=MSO_ANCHOR.MIDDLE)


def sub_band(s, x, y, w, text):
    _rect(s, x, y, w, 0.9, fill=NAVY)
    _txt(s, x, y, w, 0.9, text, size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def ctable(s, x, y, col_w, data, aligns, hl_row=None, header_h=1.0, body_h=1.3):
    rows, cols = len(data), len(data[0])
    total = header_h + (rows - 1) * body_h
    gt = s.shapes.add_table(rows, cols, Cm(x), Cm(y), Cm(sum(col_w)), Cm(total))
    t = gt.table; t.first_row = False; t.horz_banding = False
    for j, cw in enumerate(col_w):
        t.columns[j].width = Cm(cw)
    t.rows[0].height = Cm(header_h)
    for i in range(1, rows):
        t.rows[i].height = Cm(body_h)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            c = t.cell(i, j); c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = Cm(0.25); c.margin_right = Cm(0.15)
            c.margin_top = 0; c.margin_bottom = 0
            _cell_border(c)
            if i == 0:
                fill = NAVY
            elif hl_row is not None and i == hl_row:
                fill = HL_ROW
            elif i % 2 == 0:
                fill = STRIPE
            else:
                fill = WHITE
            c.fill.solid(); c.fill.fore_color.rgb = fill
            tf = c.text_frame; tf.word_wrap = True
            al = PP_ALIGN.CENTER if i == 0 else aligns[j]
            for k, line in enumerate(str(val).split("\n")):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = al
                r = p.add_run(); r.text = line
                _set_font(r, 12 if i == 0 else 11, i == 0, WHITE if i == 0 else INK)
    return y + total


def concl(s, y, text):
    _txt(s, X0, y, CW, 0.9, text, size=12, bold=True, color=CRIMSON,
         anchor=MSO_ANCHOR.MIDDLE)


def flowbox(s, x, y, w, h, text, fill=NAVY, tcolor=WHITE, line=None, lw=1.0, size=12):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.shadow.inherit = False; sp.adjustments[0] = 0.06
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is not None:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    else:
        sp.line.color.rgb = fill
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Cm(0.1))
    for i, ln in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = ln; _set_font(r, size, True, tcolor)
    return sp


def arrow(s, x1, y1, x2, y2, color=G6, w_pt=2.5):
    ln = s.shapes.add_connector(2, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.shadow.inherit = False
    ln.line.color.rgb = color; ln.line.width = Pt(w_pt)
    el = ln.line._get_or_add_ln()
    el.append(el.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return ln


SRC = "※ 출처 : Worldometer/UN · IMF · 6Wresearch · US Commerce(trade.gov) · Khmer Times/MoC"
prs = new_deck()

# ══ 표지 (유형 E) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
_txt(s, 1.5, 7.0, 24.5, 1.6, "롯데마트 동남아 유통플랫폼 확장 기획 : 캄보디아 진입",
     size=30, bold=True, color=INK, align=PP_ALIGN.CENTER)
_txt(s, 1.5, 9.2, 24.5, 1.0, "베트남 검증 역량 기반 Asset-Light 성장 방안",
     size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
_txt(s, 1.5, 15.5, 24.5, 0.8, "롯데마트   /   2026. 07.", size=11, color=G6,
     align=PP_ALIGN.CENTER)

# ══ 1 · 요약 (유형 A 2×2 + 스트립, P계열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "요약", "Executive Summary", tier="P")
band(s, "베트남 검증 역량 기반, 캄보디아 Asset-Light 진입으로 저자본·조기수익 실현")
card(s, 1.2, 3.5, 12.2, 6.0, "배경",
     "국내 유통 성숙·동남아 소비 고성장\n→ 신규 프론티어 확보 필요")
card(s, 13.9, 3.5, 12.2, 6.0, "대상",
     "메콩 3국 중 저위험·고성장·베트남 연계\n+ 개방적 외자 규제 = 캄보디아")
card(s, 1.2, 9.85, 12.2, 5.6, "방식",
     "출점 없이 B2B·PB·프랜차이즈·물류 수익화\n한국 상품·신선 특화 차별 포지션")
card(s, 13.9, 9.85, 12.2, 5.6, "기대효과",
     "저투자·조기 현금흐름 확보 후\n유통 플랫폼 단계 발전")
strip(s, 1.2, 15.65, CW,
      "관리 원칙 — 단계별 Gate 검증 후 확장 : Test → Scale → Platform")
footer(s)

# ══ 2 · 당위성① (유형 A 1×3) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "당위성①", "왜 해외 확장인가")
band(s, "내수 유통 성숙·동남아 소비 고성장 → 신규 프론티어 확보 필요")
c3 = [("내수 한계", "국내 대형마트 시장 성숙·성장 둔화로\n외연 확장 필요\n\n(※근거 수치 보강 필요)"),
      ("외부 기회", "동남아 중산층 확대·소비 고성장의\n구조적 기회"),
      ("진입 방식", "출점형 대규모 투자 부담\n→ 저자본 진입 방식 요구")]
for (lab, bod), x in zip(c3, [1.2, 9.7, 18.2]):
    card(s, x, 3.5, 8.0, 12.9, lab, bod)
footer(s)

# ══ 3 · 당위성② (유형 A 1×3 + 스트립) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "당위성②", "베트남 모델의 한계")
band(s, "성장은 지속되나 출점형 모델은 구조적 확장 제약 보유")
c3 = [("출점 여력", "베트남 진출(2008) 이후 15개점\n남부 집중 → 신규 출점 여력 제한"),
      ("자본 구조", "출점형 모델은 부지·건축·초기 투자로\n점포당 대규모 자본 소요·회수 장기화"),
      ("경쟁 환경", "AEON·Central 등 경쟁 심화로\n신규 출점 기대수익 하락")]
for (lab, bod), x in zip(c3, [1.2, 9.7, 18.2]):
    card(s, x, 3.5, 8.0, 10.2, lab, bod)
strip(s, 1.2, 14.1, CW,
      "매출 성장 지속(2024년 +20%)이 입증하듯 제약은 역량 아닌 자산 모델 "
      "→ 프론티어 시장에는 저자본 진입 대안 필요",
      fill=HL_ROW, tcolor=CRIMSON, size=12)
footer(s)

# ══ 4 · 당위성③ (유형 C 표 6열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "당위성③", "진출국 선정")
band(s, "메콩 3국 중 캄보디아가 저위험·고성장·시너지 동시 충족")
d = [["국가", "시장성", "리스크", "경쟁강도", "베트남 시너지", "종합"],
     ["캄보디아", "高", "低", "低", "高", "★★★★"],
     ["라오스", "低", "中", "低", "中", "★★"],
     ["미얀마", "高", "高", "中", "低", "★"]]
yb = ctable(s, X0, 3.6, [3.5, 4.2, 4.2, 4.2, 4.2, 4.82], d,
            aligns=[PP_ALIGN.CENTER] * 6, hl_row=1, header_h=1.2, body_h=2.9)
_txt(s, X0, yb + 0.4, CW, 2.4,
     "· 캄보디아 = 달러경제·개방 외자정책\n"
     "· 미얀마 = 정치·규제 리스크 과다\n"
     "· 라오스 = 규모·성장성 제약",
     size=12, color=INK, anchor=MSO_ANCHOR.TOP)
footer(s, "※ 등급은 인구·GDP 성장·규제·물류 연결성 기반 (부록A)")

# ══ 5 · 적합성① (유형 B 좌우 2분할) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "적합성①", "시장 매력도 및 진입 환경")
band(s, "젊은 고성장 소비시장 + 개방적 진입 규제 — 이중 호조건")
LX, RX, BW = 1.2, 13.9, 12.2
_rect(s, LX, 3.5, BW, 12.8, fill=WHITE, line=CARD_BD, line_w=1.0)
_rect(s, RX, 3.5, BW, 12.8, fill=WHITE, line=CARD_BD, line_w=1.0)
sub_band(s, LX, 3.5, BW, "시장 매력도")
sub_band(s, RX, 3.5, BW, "진입 규제 환경")
add_bullets(s, [
    (1, "인구 약 1,740만·중위연령 26.4세·생산연령 64% → 젊은 소비기반"),
    (1, "1인당 GDP 약 2,760달러(2025)·중산층 확대"),
    (1, "소매시장 CAGR 약 7~9% 고성장"),
    (1, "부분 달러라이제이션(외화예금 80%+)으로 환위험 상대적 완화"),
], l=LX + 0.4, t=4.9, w=BW - 0.8, line_h=1.45)
add_bullets(s, [
    (1, "수입·유통 외국인 참여 제한 최소"),
    (1, "상무부 브랜드 독점 수입권 제도"),
    (1, "100% 외자·JV·디스트리뷰터/프랜차이지 지정 가능"),
    (1, "→ Asset-Light 진입의 제도적 기반"),
    (1, "(※프랜차이즈 세부 법제는 향후 과제)"),
], l=RX + 0.4, t=4.9, w=BW - 0.8, line_h=1.45)
footer(s, "※ 출처 : Worldometer/UN · IMF · 6Wresearch · US Commerce(trade.gov)")

# ══ 6 · 적합성② (유형 D 피라미드 + 병렬 불릿) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "적합성②", "유통구조 — 구조적 틈새")
band(s, "현대유통 태동기 — 플레이어는 존재하나 통합 공급 역량은 부재")
cx = 7.3
for w, yy, txt, fill, tc in [
        (6.0, 3.9, "E-commerce (초기·17.8억 달러)", NAVY30, INK),
        (9.5, 7.6, "현대유통 (도시 집중·초기)", NAVY60, WHITE),
        (13.0, 11.3, "전통시장 (대부분)", NAVY, WHITE)]:
    sp = s.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Cm(cx - w / 2), Cm(yy), Cm(w), Cm(3.4))
    sp.shadow.inherit = False
    sp.fill.solid(); sp.fill.fore_color.rgb = fill; sp.line.color.rgb = fill
    tf = sp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = txt; _set_font(r, 12, True, tc)
add_bullets(s, [
    (1, "현대화는 프놈펜·씨엠립 등 도시 집중"),
    (1, "기존 플레이어 : AEON(몰)·Makro(도매)·"),
    (2, "Lucky·Chip Mong·Thai Huot(슈퍼)"),
    (1, "수입 의존·공급망 파편화"),
    (2, "→ 소싱-물류-품질 통합 공급 사업자 부재"),
], l=15.5, t=4.6, w=10.8, line_h=1.75)
strip(s, 15.5, 13.5, 10.8,
      "'빈 시장' 아닌 '통합 공급 역량의 틈새'", fill=HL_ROW, tcolor=CRIMSON, size=13, h=1.5)
footer(s)

# ══ 7 · 적합성③ (유형 C 표 3열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "적합성③", "Right to Win")
band(s, "기존 플레이어와 정면 경쟁 없는 차별화 포지션 확보 가능")
d = [["플레이어", "포맷", "우리와의 관계"],
     ["Makro", "범용 캐시앤캐리·도매", "범용 도매 vs 한국 상품·PB·신선 특화 — 카테고리 비중첩"],
     ["AEON", "자산형 몰·GMS", "출점 경쟁 아닌 무출점 공급자 — 잠재 B2B 고객화"],
     ["Lucky·Chip Mong·Thai Huot", "로컬 슈퍼", "경쟁 아닌 핵심 B2B 공급 대상"]]
yb = ctable(s, X0, 3.6, [5.0, 7.0, 13.12], d,
            aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
            header_h=1.1, body_h=2.9)
concl(s, yb + 0.5,
      "차별화 3종 = ① K-상품·PB 독점 소싱  ② 신선 콜드체인 표준  ③ 베트남 연계 크로스보더 물류")
_txt(s, X0, yb + 1.4, CW, 0.8,
     "— 기존 소매사업자는 경쟁자가 아닌 B2B 고객 후보", size=12, bold=True, color=CRIMSON)
footer(s, "※ 각 사 점포수·포맷 세부는 현지 조사 확정 (부록C)")

# ══ 8 · 적합성④ (유형 C 표 2열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "적합성④", "기회 정의")
band(s, "캄보디아 Pain Point ↔ 롯데 Capability 정합")
d = [["시장 현황 (Pain)", "롯데 강점 (Capability)"],
     ["공급망 비효율·파편화", "SCM · Hub & Spoke 물류"],
     ["신선 품질 불균일", "품질관리 표준화"],
     ["상품 다양성 부족", "글로벌 소싱 · PB"],
     ["운영 역량 미숙", "매장 운영 매뉴얼"]]
yb = ctable(s, X0, 3.6, [12.56, 12.56], d,
            aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT], header_h=1.1, body_h=2.6)
concl(s, yb + 0.5, "결론 — 역량 적합성(Capability-Fit) 높은 시장")
footer(s)

# ══ 9 · 적합성⑤ (유형 A 1×3 + 스트립) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "적합성⑤", "우리의 무기 — 크로스보더 공급 네트워크")
band(s, "베트남 남부 인프라를 캄보디아로 연장하는 국경 연계 공급망")
c3 = [("공급망 플랫폼", "베트남 남부 15개점 소싱·물류\n→ 육로 국경(목바이/바벳)\n→ 프놈펜"),
      ("소싱 파워", "PB 개발·글로벌 직소싱 역량"),
      ("운영 표준화", "신선·델리·매장 운영 매뉴얼 이식")]
for (lab, bod), x in zip(c3, [1.2, 9.7, 18.2]):
    card(s, x, 3.5, 8.0, 10.5, lab, bod)
strip(s, 1.2, 14.3, CW,
      "재정의 — SCM → Cross-border Supply Network  /  PB → Profit Engine")
footer(s, "※ 국경 통관 리드타임·관세는 검증 필요 (부록C) — 타당성 조사 1순위")

# ══ 10 · 수익모델 (유형 D 돈흐름 플로우) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "수익모델", "Asset-Light 구조")
band(s, "출점 없이 B2B·PB·로열티·물류로 수익 — 저투자·고마진")
flowbox(s, 1.2, 6.0, 5.2, 4.0, "베트남 남부\n소싱·물류", fill=NAVY)
_txt(s, 6.5, 7.3, 1.7, 0.5, "국경 연계", size=10, color=G6, align=PP_ALIGN.CENTER)
arrow(s, 6.5, 8.0, 8.15, 8.0)
flowbox(s, 8.2, 6.0, 5.6, 4.0, "B2B 도매\n— 초기 현금흐름", fill=WHITE, tcolor=NAVY,
        line=CRIMSON, lw=2.0)
for label, fill, yy in [("PB 고마진", NAVY, 4.2), ("프랜차이즈 로열티", NAVY, 7.9),
                        ("물류 수익", NAVY, 11.6)]:
    arrow(s, 13.9, 8.0, 16.35, yy + 1.6)
    flowbox(s, 16.4, yy, 9.9, 3.2, label, fill=fill)
_txt(s, 1.2, 13.5, 14.5, 1.0,
     "Revenue Stream 4종 = 상품공급 마진 · PB · 로열티 · 물류",
     size=12, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
strip(s, 1.2, 15.4, CW, "Low Investment · High Margin · Fast Expansion", size=13)
footer(s)

# ══ 11 · 실행① (유형 C 표 3열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "실행①", "단계별 진입 및 Gate 기준")
band(s, "단계별 검증 통과 시에만 확장 — Test → Scale → Platform")
d = [["단계", "내용", "전환 Gate (가정치)"],
     ["Step1  B2B", "상품 공급 저위험 진입·현금흐름", "월 공급액·활성 파트너 N개·목표 마진율 달성"],
     ["Step2  Franchise", "마스터 프랜차이즈 무자본 확장", "가맹 점포 수·점당 매출·로열티 안정화"],
     ["Step3  Platform", "유통 + 물류 생태계 완성", "— (장기 비전)"]]
yb = ctable(s, X0, 3.6, [4.5, 10.0, 10.62], d,
            aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
            header_h=1.1, body_h=2.9)
concl(s, yb + 0.5, "Gate 미달 시 확장 보류·모델 수정 — 손실 하방 차단 구조")
footer(s, "※ Gate 수치는 파일럿 설계 시 확정 (부록B)")

# ══ 12 · 실행② (유형 C 표 2열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "실행②", "실행 로드맵")
band(s, "개시 12개월 전 착수, KPI 기반 관리")
d = [["시점", "과제"],
     ["M-12 (개시 12개월 전)", "프로젝트 착수·시장 정밀조사"],
     ["M-10", "물류 인프라·국경 통관 설계"],
     ["M-8", "파트너 발굴·실사"],
     ["M-3", "상품 박람회·소싱 확정"],
     ["M-0", "B2B 사업 개시"]]
yb = ctable(s, X0, 3.6, [5.5, 19.62], d,
            aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT], hl_row=5, header_h=1.1, body_h=2.15)
concl(s, yb + 0.5, "KPI — 매출·취급 SKU·확보 파트너 수·공급 리드타임")
footer(s)

# ══ 13 · 리스크 (유형 C 표 2열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "리스크", "대응 체계")
band(s, "관리 가능한 리스크 구조")
d = [["리스크", "대응"],
     ["파트너 리스크", "복수 파트너·사전 실사"],
     ["물류·통관 리스크", "국경 통관 사전 검증(M-10)·현지 3PL 병행"],
     ["경쟁 대응 (Makro 등 도매 확장)", "K-상품·PB·신선 특화 차별화"],
     ["환율·통화", "달러 결제 기본·정부 리엘화 정책 모니터링 (완전 달러화 아님)"],
     ["프랜차이즈 법제 불확실성", "현지 법률 검토 선행"]]
ctable(s, X0, 3.6, [8.0, 17.12], d,
       aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT], header_h=1.1, body_h=2.35)
footer(s)

# ══ 14 · 제언 (유형 A 좌2+우1+스트립, P계열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "제언", "방향 및 향후 과제", tier="P")
band(s, "저자본 파일럿 우선 추진, Gate 검증 후 단계 확장")
card(s, 1.2, 3.5, 12.2, 6.0, "제언",
     "B2B 파일럿 저위험 진입\nGate 통과 시 프랜차이즈·플랫폼 확장")
card(s, 1.2, 10.2, 12.2, 6.0, "기대효과",
     "저투자 조기 현금흐름\n베트남 역량 재활용·통합 공급 틈새 선점")
card(s, 13.9, 3.5, 12.2, 9.5, "향후 검토과제",
     "① 경쟁사 점포·포맷 실사\n\n② 국경 물류 리드타임·관세 타당성\n\n"
     "③ 프랜차이즈 법제\n\n④ 재무 추정 정교화")
strip(s, 13.9, 13.4, 12.2,
      "다음 단계 — 하반기 내 현지 정밀조사 착수 → 연내 파일럿 실행계획 수립",
      size=12)
footer(s)

# ══ 부록A (유형 C 표 3열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "부록A", "시장 데이터 — 캄보디아")
band(s, "캄보디아 주요 시장·규제 지표 원출처 정리")
d = [["항목", "값", "출처"],
     ["인구", "약 1,740만", "Worldometer/UN"],
     ["중위연령", "26.4세", "UN"],
     ["생산연령 비중", "64%", "UN"],
     ["1인당 GDP", "약 2,760달러 (2025)", "IMF"],
     ["소매 성장률", "CAGR 약 7~9%", "6Wresearch"],
     ["E-commerce", "약 17.8억 달러 (2025)·초기", "Khmer Times/MoC"],
     ["통화", "부분 달러화 (외화예금 80%+·리엘 병용)", "IMF"],
     ["규제", "외국인 유통 제한 최소·독점 수입권·100% 외자/JV/프랜차이지", "US Commerce(trade.gov)"]]
ctable(s, X0, 3.4, [5.0, 12.0, 8.12], d,
       aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT], header_h=1.1, body_h=1.55)
footer(s)

# ══ 부록B (유형 C 표 — 가정치·예시) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "부록B", "재무 추정 — Phase별 예시 프레임")
band(s, "가정치(예시) — 실제 수치는 재무 추정 정교화 후 확정")
d = [["Phase", "주요 투자", "매출 구성", "회수 관점", "Gate 연계"],
     ["Step1  B2B", "가정치 설정 대상", "상품공급 마진", "가정치 설정 대상", "월 공급액·마진율"],
     ["Step2  Franchise", "가정치 설정 대상", "로열티 + 공급 마진", "가정치 설정 대상", "가맹 수·로열티"],
     ["Step3  Platform", "가정치 설정 대상", "플랫폼·물류 수익", "가정치 설정 대상", "장기 비전"]]
ctable(s, X0, 3.6, [4.6, 5.2, 5.5, 5.0, 4.82], d,
       aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT],
       header_h=1.1, body_h=3.05)
footer(s, "※ 본 표는 가정치(예시) — 수치 셀은 파일럿 설계 시 확정, 임의 수치 생성 금지")

# ══ 부록C (유형 C 표 2열) ══
s = prs.slides.add_slide(prs.slide_layouts[6])
header(s, "부록C", "검증 필요 항목 (타당성 조사 체크리스트)")
band(s, "타당성 조사 1순위 항목 — 진입 전 확정 필요")
d = [["항목", "확인 내용"],
     ["경쟁사 점포수·포맷", "Makro·AEON 등 점포망·B2B 커버리지 실사"],
     ["국경 물류", "목바이/바벳 통관 리드타임·관세·콜드체인 가능성"],
     ["프랜차이즈 법제", "가맹 관련 현지 법규·로열티 송금 규정"],
     ["국내 시장 근거", "대형마트 성숙 판단 근거 수치 확보"]]
ctable(s, X0, 3.6, [9.0, 16.12], d,
       aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT], header_h=1.1, body_h=2.85)
footer(s)

out = os.path.join(os.path.dirname(__file__), "..", "output", "캄보디아_확장기획.pptx")
prs.save(out)
print("saved", os.path.normpath(out), "| slides:", len(prs.slides._sldIdLst))
