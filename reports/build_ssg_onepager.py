# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

본장 프레임워크 유형 A(SPEC 3-1, 정기협의체 8·14/15 실측) 적용:
■ 리드 → 컬럼 헤더 바 2개 + 중앙 세로선 → 컬럼 내 • 불릿·표 → 네이비 결론 박스.
좌: 도입 개요(Why/What) | 우: 배송 체계·확대 계획(How/Plan).
출처: 비즈워치 「[인사이드 스토리] SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *

SRC = "※ 출처 : 비즈워치 「SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)"
HALF_W = (BODY_W - COL_GAP) / 2
COL2_L = MARGIN_L + HALF_W + COL_GAP
VSEP_X = MARGIN_L + HALF_W + COL_GAP / 2


def block(slide, l, t, w, head, subs):
    """• 주항목 + - 하위 불릿 블록 → 종료 y 반환."""
    items = [(0, head)] + [(1, s) for s in subs]
    return add_bullets(slide, items, l=l, t=t, w=w)


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·PP센터 가동률 제고",
                footnote=SRC)

# ── 컬럼 헤더 바 + 중앙 세로 구분선 ──
Y0 = BODY_TOP_LEAD
add_col_header(s, MARGIN_L, Y0, HALF_W, "2시간 배송 도입 개요")
add_col_header(s, COL2_L, Y0, HALF_W, "배송 체계 및 확대 계획")
y_body = Y0 + COL_HEADER_H + 0.35

# ── 좌 컬럼 : Why / What ──
y = y_body
y = block(s, MARGIN_L, y, HALF_W, "도입 배경", [
    "이륜차 퀵커머스 소량 한정 → 대용량 공백",
    "바로퀵 매출 197%↑(6월) → 수요 실증",
])
y += 0.55
y = block(s, MARGIN_L, y, HALF_W, "시범 운영 ('26. 7. 9. 개시)", [
    "20시 주문 마감 → 결제 후 2시간 내 배송",
    "예약배송 병행, 사륜차 대용량 대응",
])
y += 0.55
data = [
    ["구분", "품목 수", "비고"],
    ["이마트 점포", "약 150,000", "2시간 배송 기반"],
    ["바로퀵", "약 10,000", "이륜차 적재 한계"],
    ["B마트", "약 20,000", "퀵커머스 1위"],
]
add_fin_table(s, MARGIN_L, y, HALF_W, 3.0, data,
              col_w=[3.4, 3.4, HALF_W - 6.8], header_rows=1,
              col_align=["c", "r", "c"])

# ── 우 컬럼 : How / Plan ──
y = y_body
data = [
    ["구분", "속도 · 수단", "특징"],
    ["주간배송", "예약 4~5시간 · 사륜", "지역별 2~5개 구간"],
    ["새벽배송", "새벽 · 사륜", "네오 한정 (PP센터 불가)"],
    ["트레이더스배송", "예약 · 사륜", "—"],
    ["바로퀵", "1시간 · 이륜", "소량, 약 1만 종"],
    ["2시간 배송 (신규)", "2시간 · 사륜", "대용량, 약 15만 종"],
]
add_fin_table(s, COL2_L, y, HALF_W, 4.4, data,
              col_w=[3.4, 3.9, HALF_W - 7.3], header_rows=1)
y += 4.4 + 0.55
y = block(s, COL2_L, y, HALF_W, "확대 로드맵", [
    "7월 양재·하남 → 8월 서울 → 9월 전국",
    "연말 50여 개 점포 운영 목표",
])
y += 0.55
y = block(s, COL2_L, y, HALF_W, "규제 이슈", [
    "유통법 영업 제한 → PP센터 새벽 불가",
    "규제 완화 시 새벽배송 확장 여지",
])

# ── 컬럼 결론 박스 (So-What) + 세로선 ──
CONCL_Y = y + 0.55
add_conclusion_box(s, MARGIN_L, CONCL_Y, HALF_W,
                   "사륜차 대용량 즉시배송으로 퀵커머스 공백 흡수")
add_conclusion_box(s, COL2_L, CONCL_Y, HALF_W,
                   "연말 50여 개 점포 — 전국 즉시배송 기반 구축")
add_vsep(s, VSEP_X, Y0, CONCL_Y + CONCL_H - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out, "slides:", len(prs.slides._sldIdLst), "| concl y:", round(CONCL_Y, 2))
