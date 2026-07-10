# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

정기협의체 레이아웃 프레임워크 적용: 좌측 레이블 레일(배경→핵심→체계→로드맵)로
주제를 구획하고 우측에 불릿·키메시지 박스·비교표·단계 박스를 행 단위 배치.
출처: 비즈워치 「[인사이드 스토리] SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt

SRC = "※ 출처 : 비즈워치 「SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)"
BOX_GAP = 0.30


def row_bullets(slide, t, items):
    """레일 행 콘텐츠: 하위 불릿(-) 목록."""
    y = t + 0.12
    for text in items:
        _txt(slide, CONTENT_L, y, CONTENT_W, LINE_H, "- " + text,
             size=FONT_PT["bullet1"], color=INK)
        y += 0.55


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·PP센터 가동률 제고",
                footnote=SRC)

y = BODY_TOP_LEAD

# 1행. 배경 (Why)
H1 = 1.45
add_label_row(s, y, H1, "배경")
row_bullets(s, y, [
    "이륜차 퀵커머스는 소량 한정 → 3~4인 가구 대용량 장보기 공백",
    "바로퀵 매출 197%↑(6월, 1월比) → 온디맨드 수요 실증",
])
y += H1 + ROW_GAP

# 2행. 핵심 (What) — 키메시지 박스 3개
H2 = 1.95
add_label_row(s, y, H2, "핵심")
kw = (CONTENT_W - 2 * BOX_GAP) / 3
add_frame_box(s, CONTENT_L, y, kw, H2,
              caption="배송 속도", title="결제 후 2시간 내", sub="20시 주문 마감")
add_frame_box(s, CONTENT_L + kw + BOX_GAP, y, kw, H2,
              caption="운송·적재", title="사륜차 대용량", sub="중량 상품 즉시 배송")
add_frame_box(s, CONTENT_L + 2 * (kw + BOX_GAP), y, kw, H2,
              caption="상품 구색", title="약 15만 종", sub="바로퀵 1만 · B마트 2만")
y += H2 + ROW_GAP

# 3행. 체계 (How) — 배송 서비스 비교표
H3 = 5.00
add_label_row(s, y, H3, "서비스\n체계")
data = [
    ["구분", "배송 속도", "운송 수단", "취급 품목", "특징"],
    ["주간배송", "예약배송 (4~5시간)", "사륜차", "점포 상품", "지역별 2~5개 구간"],
    ["새벽배송", "새벽 시간대", "사륜차", "네오 상품", "네오 한정 (PP센터 불가)"],
    ["트레이더스배송", "예약배송", "사륜차", "트레이더스 상품", "—"],
    ["바로퀵", "1시간 내 (3km)", "이륜차", "약 1만 종", "소량·1~2인 가구"],
    ["2시간 배송 (신규)", "2시간 내 (20시 마감)", "사륜차", "약 15만 종", "대용량 즉시배송"],
]
add_fin_table(s, CONTENT_L, y, CONTENT_W, H3, data,
              col_w=[3.6, 4.6, 2.4, 3.3, CONTENT_W - 13.9], header_rows=1)
y += H3 + ROW_GAP

# 4행. 로드맵 (Plan) — 단계 박스 (현재 단계 네이비 강조)
H4 = 1.95
add_label_row(s, y, H4, "로드맵")
pw = (CONTENT_W - 3 * BOX_GAP) / 4
phases = [
    ("'26. 7월", "양재·하남 시범", "7. 9. 개시", True),
    ("'26. 8월", "서울 확대", "주요 지역", False),
    ("'26. 9월", "전국 확대", "쓱배송 권역", False),
    ("'26. 연말", "50여 개 점포", "운영 목표", False),
]
for i, (cap, title, sub, active) in enumerate(phases):
    add_frame_box(s, CONTENT_L + i * (pw + BOX_GAP), y, pw, H4,
                  caption=cap, title=title, sub=sub, active=active)
y += H4 + ROW_GAP

# 규제 참고 (크림슨)
_txt(s, MARGIN_L, y + 0.15, BODY_W, LINE_H,
     "※ 유통산업발전법 : 영업 0~10시 제한·월 2회 휴업 → PP센터 새벽배송 불가 (규제 완화 시 확장 여지)",
     size=FONT_PT["table"], bold=False, color=CRIMSON)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out, "slides:", len(prs.slides._sldIdLst), "| last y:", round(y, 2))
