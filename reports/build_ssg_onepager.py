# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

논리 흐름이 프레임워크를 결정(SPEC 3-1): 위→아래·좌→우 읽기 순서 = 설명 순서.
좌 컬럼(현황·배경 : 문제) → 우 컬럼(추진 방안 : 해법) → 하단 결론(So-What).
정기협의체 15/15 백업 장표의 2단 서사 구조 실측 반영.
출처: 비즈워치 「[인사이드 스토리] SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 · B마트 2만 종      ※ 출처 : 비즈워치 ('26. 7. 10.)"
HALF_W = (BODY_W - COL_GAP) / 2
COL2_L = MARGIN_L + HALF_W + COL_GAP
VSEP_X = MARGIN_L + HALF_W + COL_GAP / 2
PITCH = 0.76                 # 불릿 행간 (수직 균형 확대 적용)
GAP_BLK = 0.50               # 블록 간 간격


def col_title(slide, l, t, text):
    """컬럼 제목 — 가운데 볼드 + 굵은 밑줄 룰 (15/15 실측)."""
    _txt(slide, l, t, HALF_W, 0.55, text, size=FONT_PT["bullet0"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(slide, l, t + 0.62, HALF_W, color=INK, w_pt=1.5)
    return t + 0.90


def block(slide, l, t, head, subs):
    """● 주항목 + - 하위 블록 → 종료 y."""
    items = [(0, head)] + [(1, s) for s in subs]
    return add_bullets(slide, items, l=l, t=t, w=HALF_W, line_h=PITCH)


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·PP센터 가동률 제고",
                footnote=SRC)

Y0 = BODY_TOP_LEAD

# ══ 좌 컬럼 : 현황 및 도입 배경 (문제 정의) ══
y = col_title(s, MARGIN_L, Y0, "현황 및 도입 배경")
y = block(s, MARGIN_L, y, "퀵커머스 수요 급증", [
    "이륜차 퀵커머스 소량 한정 → 대용량 퀵커머스 서비스 공백",
    "바로퀵 매출 197%↑(6월) → 수요 실증",
])
y += GAP_BLK
y = block(s, MARGIN_L, y, "현행 배송 체계", [])
data = [
    ["구분", "속도 · 수단", "특징"],
    ["주간배송", "예약 4~5시간 · 사륜", "지역별 2~5개 구간"],
    ["새벽배송", "새벽 · 사륜", "네오 한정"],
    ["트레이더스배송", "예약 · 사륜", "—"],
    ["바로퀵", "1시간 · 이륜", "소량, 약 1만 종"],
]
add_fin_table(s, MARGIN_L, y, HALF_W, 4.8, data,
              col_w=[3.2, 4.2, HALF_W - 7.4], header_rows=1)
y += 4.8 + GAP_BLK
block(s, MARGIN_L, y, "규제 제약", [
    "유통산업발전법 → PP센터 새벽배송 불가",
])

# ══ 우 컬럼 : 추진 방안 (해법) ══
y = col_title(s, COL2_L, Y0, "추진 방안 : '2시간 배송'")
y = block(s, COL2_L, y, "시범 운영 ('26. 7. 9.~)", [
    "20시 주문 마감 → 2시간 내 배송",
    "기존 예약배송 병행 선택 가능",
    "사륜차 활용, 대용량·중량 대응",
])
y += GAP_BLK
y = block(s, COL2_L, y, "기대 효과", [
    "대용량 즉시배송 서비스 공백 흡수",
    "점포 15만 종 구색 → 즉시배송 확장",
    "PP센터 낮 시간대 가동률 제고",
])
y += GAP_BLK
y = block(s, COL2_L, y, "확대 로드맵", [])
add_timeline(s, COL2_L, y + 0.05, HALF_W, [
    ("'26. 7월", "양재·하남"),
    ("8월", "서울 확대"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포"),
])

# ══ 중앙 세로선 + 전폭 결론 박스 (하단 앵커) ══
add_vsep(s, VSEP_X, Y0, 15.45 - Y0)
add_conclusion_box(s, MARGIN_L, 16.30, BODY_W,
                   "사륜차 대용량 즉시배송으로 퀵커머스 공백 흡수 — 연말 50여 개 점포, 전국 즉시배송 기반 구축",
                   h=1.00)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
