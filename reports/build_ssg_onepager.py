# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

P계열 본문 1매. 리드메시지 → 도입 배경 → 시범 운영 개요 → 서비스 비교 →
품목 비교·확대 로드맵 순의 논리 구성. 간격·타이포·정렬·불릿은 표준 토큰 상속.
출처: 비즈워치 「[인사이드 스토리] SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt

SRC = "※ 출처 : 비즈워치 「SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)"
COL_GAP = 0.66
HALF_W = (BODY_W - COL_GAP) / 2
COL2_L = MARGIN_L + HALF_W + COL_GAP
SECTION_H = 0.55            # 소제목 텍스트박스 높이


def section(slide, l, t, w, text):
    """소제목 — 표준 불릿 체계 주항목(•)."""
    return _txt(slide, l, t, w, SECTION_H, "• " + text,
                size=FONT_PT["bullet0"], bold=True, color=BLACK)


def bullet(slide, l, t, w, text):
    """표준 불릿 체계 하위항목(-)."""
    return _txt(slide, l + BULLET_INDENT, t, w - BULLET_INDENT, LINE_H,
                "- " + text, size=FONT_PT["bullet1"], color=INK)


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·PP센터 가동률 제고",
                footnote=SRC)

y = BODY_TOP_LEAD                          # 본문 커서

# 1. 도입 배경
section(s, MARGIN_L, y, BODY_W, "도입 배경")
bullet(s, MARGIN_L, y + GAP_HEAD, BODY_W,
       "이륜차 퀵커머스는 소량 한정 → 3~4인 가구 대용량 장보기 공백")
bullet(s, MARGIN_L, y + GAP_HEAD + LINE_H, BODY_W,
       "바로퀵 매출 197%↑(6월, 1월比) → 온디맨드 수요 실증")
y += GAP_HEAD + 2 * LINE_H + GAP_SECTION

# 2. 시범 운영 개요
section(s, MARGIN_L, y, BODY_W, "시범 운영 개요 ('26. 7. 9. 개시)")
bullet(s, MARGIN_L, y + GAP_HEAD, BODY_W,
       "20시까지 주문 → 결제 후 2시간 내 배송, 예약배송 병행 가능")
bullet(s, MARGIN_L, y + GAP_HEAD + LINE_H, BODY_W,
       "사륜차로 대용량·중량 상품 즉시 배송 (약 15만 종 기반)")
y += GAP_HEAD + 2 * LINE_H + GAP_SECTION

# 3. 배송 서비스 비교
MAIN_TBL_H = 4.50
section(s, MARGIN_L, y, BODY_W, "이마트·SSG닷컴 배송 서비스 비교")
data = [
    ["구분", "배송 속도", "운송 수단", "취급 품목", "특징"],
    ["주간배송", "예약배송 (4~5시간)", "사륜차", "점포 상품", "지역별 2~5개 구간"],
    ["새벽배송", "새벽 시간대", "사륜차", "네오 상품", "네오 한정 (PP센터 불가)"],
    ["트레이더스배송", "예약배송", "사륜차", "트레이더스 상품", "—"],
    ["바로퀵", "1시간 내 (3km)", "이륜차", "약 1만 종", "소량·1~2인 가구"],
    ["2시간 배송 (신규)", "2시간 내 (20시 마감)", "사륜차", "약 15만 종", "대용량 즉시배송"],
]
add_fin_table(s, MARGIN_L, y + GAP_HEAD, BODY_W, MAIN_TBL_H, data,
              col_w=[3.9, 5.3, 2.5, 3.6, 9.86], header_rows=1)
y += GAP_HEAD + MAIN_TBL_H + GAP_SECTION

# 4. 하단 2단 (취급 품목 수 비교 / 확대 로드맵)
SUB_TBL_H = 2.60
section(s, MARGIN_L, y, HALF_W, "취급 품목 수 비교")
data = [
    ["구분", "품목 수", "비고"],
    ["이마트 점포", "약 150,000", "2시간 배송 기반"],
    ["바로퀵", "약 10,000", "이륜차 적재 한계"],
    ["B마트", "약 20,000", "퀵커머스 1위"],
]
add_fin_table(s, MARGIN_L, y + GAP_HEAD, HALF_W, SUB_TBL_H, data,
              col_w=[3.5, 3.6, HALF_W - 7.1], header_rows=1,
              header_fill=TH_SECOND, col_align=["c", "r", "c"])

section(s, COL2_L, y, HALF_W, "확대 로드맵")
data = [
    ["시기", "계획"],
    ["'26. 7월", "양재·하남점 시범 운영"],
    ["'26. 8월", "서울 주요 지역 확대"],
    ["'26. 9월 → 연말", "전국 확대 → 50여 개 점포"],
]
add_fin_table(s, COL2_L, y + GAP_HEAD, HALF_W, SUB_TBL_H, data,
              col_w=[3.9, HALF_W - 3.9], header_rows=1,
              header_fill=TH_SECOND)
y += GAP_HEAD + SUB_TBL_H + GAP_SECTION

# 규제 참고
_txt(s, MARGIN_L, y, BODY_W, LINE_H,
     "※ 유통산업발전법 : 영업 0~10시 제한·월 2회 휴업 → PP센터 새벽배송 불가 (규제 완화 시 확장 여지)",
     size=FONT_PT["table"], bold=False, color=CRIMSON)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out, "slides:", len(prs.slides._sldIdLst), "| last y:", round(y, 2))
