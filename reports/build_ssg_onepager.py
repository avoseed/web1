# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 경쟁사 동향 1매 — 소량 퀵커머스 → 全 상품 장보기 채널.

본질: 취급 상품 소량 1만 종 → 마트 全 상품 15만 종 확대 · 점포(PP센터) 온라인 가동률 제고.
3섹션: ① 배경 → ② 서비스 개요 및 상품 차별점(7행 비교표, 장표 중심) → ③ 확산 로드맵(최종 요소).
표준 정합: 결론은 리드에 선치 · 하단 네이비 박스 없음(로드맵이 본문 종결) · 全 상품 통일.
출처: 비즈워치·머니투데이·한국경제·뉴스핌·비즈니스포스트('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, _set_font, MSO_ANCHOR

FN = "※ 출처 : 비즈워치 · 머니투데이 · 한국경제 · 뉴스핌 · 비즈니스포스트 ('26.7.8~10)"

prs = new_deck()
# ══ 제목(■ 제목 전용) ══
s = add_content(prs, None, "이마트, 점포 기반 '2시간 배송' 도입", tier="P")
# ══ 리드(볼드 1줄·검정) — 핵심 대비 구간만 크림슨 강조 ══
lb = s.shapes.add_textbox(Cm(LEAD[0]), Cm(LEAD[1]), Cm(24.8), Cm(LEAD[3]))
ltf = lb.text_frame; ltf.word_wrap = True
for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
    setattr(ltf, m, 0)
lp = ltf.paragraphs[0]; lp.alignment = PP_ALIGN.LEFT
for txt, col in [("취급 상품 ", INK),
                 ("소량 1만 종 → 마트 全 상품 15만 종 확대", CRIMSON),
                 (" — 점포(PP센터) 온라인 가동률 제고 목적", INK)]:
    r = lp.add_run(); r.text = txt; _set_font(r, FONT_PT["lead"], True, col)

# ══ ① 도입 배경 ══
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "'바로퀵'(1시간 배송) 매출 197% 증가('26.6월) → 퀵커머스 대용량 상품 수요 확대"),
    (1, "B마트 등 업계 대용량 품목 취급 확대 → 소량 1시간(이륜차)·대용량 2시간(사륜차) 이원화 대응"),
], l=MARGIN_L, t=4.02, w=BODY_W, line_h=0.56)

# ══ ② 서비스 개요 및 상품 차별점 (7행 비교표 — 장표 중심) ══
y = add_bullets(s, [
    (0, "② 서비스 개요 및 상품 차별점 — 소량 퀵커머스에서 全 상품 장보기로"),
], l=MARGIN_L, t=y + 0.16, w=BODY_W, line_h=0.54)
TBL_TOP = y + 0.06
H_TBL = 6.9
tbl = [
    ["구분", "'바로퀵' (기존)", "★ '2시간 배송' (신규)"],
    ["배송", "1H 이내 (반경 3km) · 이륜차", "2H 이내 · 사륜차 · 20시 주문 마감"],
    ["취급", "약 1만 종 (소포장·즉석식품 중심)", "약 15만 종 (마트 全 상품) — 대용량·신선 전면 개방"],
    ["신선·포장", "일반 포장", "육류·생선·과일 매장 직접 집품 · 전용 보냉 파우치 콜드체인"],
    ["출고", "매장 일부 품목", "PP센터 집품·포장 → 즉시 출고"],
    ["고객·용도", "2030 · 소용량 (그로서리 90%)", "본 장보기 수요 겨냥"],
    ["조건", "별도 배송비 체계", "무료배송 4만 원 이상 (쓱배송 동일)"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, BODY_W, H_TBL, tbl,
              col_w=[3.4, 9.4, BODY_W - 12.8], header_rows=1,
              hl_cols=[2], bold_rows=[2], col_align=["c", "l", "l"])
tfn_y = TBL_TOP + H_TBL + 0.05
_txt(s, MARGIN_L, tfn_y, BODY_W, 0.40,
     "※ 예약배송(주간·새벽·트레이더스) 기존 운영 유지, 병행 이용 가능",
     size=FONT_PT["footnote"], color=SUB_GRAY)

# ══ ③ 확산 로드맵 (가로 타임라인 — 본문 최종 요소, 연말 마디 강조) ══
y = add_bullets(s, [(0, "③ 확산 로드맵")], l=MARGIN_L, t=tfn_y + 0.44, w=BODY_W, line_h=0.54)
TL_TOP = y + 0.06
tl_bot = add_htimeline(s, MARGIN_L, TL_TOP, BODY_W, [
    ("'26.7월", "양재·하남점 시범"),
    ("8월", "서울권 확대(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
], when_h=0.45, content_h=0.58, size=FONT_PT["bullet2"])
# 종착점(연말 마디) 강조 — 크림슨 채움 점
axis_y = TL_TOP + 0.45 + 0.12
cx = MARGIN_L + (BODY_W / 4) * 3.5
dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(cx - 0.17), Cm(axis_y - 0.17), Cm(0.34), Cm(0.34))
dot.shadow.inherit = False
dot.fill.solid(); dot.fill.fore_color.rgb = CRIMSON; dot.line.color.rgb = CRIMSON

# ══ 각주(콘텐츠 직하) ══
_txt(s, MARGIN_L, tl_bot + 0.28, BODY_W, 0.5, FN, size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
