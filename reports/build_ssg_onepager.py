# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

본장 프레임워크 유형 B(SPEC 3-1, 정기협의체 12/15 실측) 적용:
■ 리드 → 상단 전폭 비교표(근거) → 하단 균등 3단(밑줄 단제목 + • 불릿)
→ 전폭 네이비 결론 박스. 좌우 대칭 구조.
출처: 비즈워치 「[인사이드 스토리] SSG닷컴, '2시간 실험'에 나선 진짜 이유는」('26. 7. 10.)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 · B마트 2만 종      ※ 출처 : 비즈워치 ('26. 7. 10.)"
N_COL = 3
COL_W = (BODY_W - (N_COL - 1) * COL_GAP) / N_COL      # 균등 3단 폭


PANEL_PITCH = 0.66      # 하단 3단 불릿 행간 (수직 균형 위해 표준 LINE_H 확대 적용)


def panel(slide, idx, title, items, t):
    """하단 해설 단: 가운데 단제목 + 밑줄 룰 + 표준 불릿 (12/15 실측)."""
    l = MARGIN_L + idx * (COL_W + COL_GAP)
    _txt(slide, l, t, COL_W, 0.55, title, size=FONT_PT["bullet0"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(slide, l, t + 0.62, COL_W, color=INK, w_pt=1.0)
    return add_bullets(slide, items, l=l, t=t + 0.90, w=COL_W, line_h=PANEL_PITCH)


from pptx.enum.text import PP_ALIGN

prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·PP센터 가동률 제고",
                footnote=SRC)

# ── 상단 전폭 : 배송 서비스 비교표 (근거) ──
# 수직 균형: 본문이 각주 직전(~17.3)까지 페이지를 채우도록 행고·간격 비례 확대
Y0 = BODY_TOP_LEAD
TBL_H = 7.20
data = [
    ["구분", "배송 속도", "운송 수단", "취급 품목", "특징"],
    ["주간배송", "예약배송 (4~5시간)", "사륜차", "점포 상품", "지역별 2~5개 구간"],
    ["새벽배송", "새벽 시간대", "사륜차", "네오 상품", "네오 한정 (PP센터 불가)"],
    ["트레이더스배송", "예약배송", "사륜차", "트레이더스 상품", "—"],
    ["바로퀵", "1시간 내 (3km)", "이륜차", "약 1만 종", "소량·1~2인 가구"],
    ["2시간 배송 (신규)", "2시간 내 (20시 마감)", "사륜차", "약 15만 종", "대용량 즉시배송"],
]
add_fin_table(s, MARGIN_L, Y0, BODY_W, TBL_H, data,
              col_w=[3.9, 5.3, 2.5, 3.6, BODY_W - 15.3], header_rows=1)

# ── 하단 균등 3단 : 배경 → 시범 운영 → 확대 계획 ──
PY = Y0 + TBL_H + 0.75
ends = [
    panel(s, 0, "도입 배경", [
        (0, "대용량 장보기 공백"),
        (1, "이륜차 퀵커머스 소량 한정"),
        (0, "온디맨드 수요 실증"),
        (1, "바로퀵 매출 197%↑ (6월)"),
    ], PY),
    panel(s, 1, "시범 운영 ('26. 7. 9.~)", [
        (0, "2시간 내 배송"),
        (1, "20시 주문 마감, 예약배송 병행"),
        (0, "사륜차 대용량 대응"),
        (1, "점포 15만 종 기반 즉시배송"),
    ], PY),
    panel(s, 2, "확대 로드맵", [
        (0, "전국 확대"),
        (1, "8월 서울 → 9월 전국"),
        (1, "연말 50여 개 점포 목표"),
        (0, "규제 변수 : PP센터 새벽 불가"),
    ], PY),
]
# 단 사이 세로 구분선
P_BOT = max(ends)
for i in range(1, N_COL):
    x = MARGIN_L + i * (COL_W + COL_GAP) - COL_GAP / 2
    add_vsep(s, x, PY + 0.1, P_BOT - PY - 0.1)

# ── 전폭 결론 박스 (So-What) — 하단 앵커 (각주 상단 정렬) ──
CONCL_Y = 16.30
add_conclusion_box(s, MARGIN_L, CONCL_Y, BODY_W,
                   "사륜차 대용량 즉시배송으로 퀵커머스 공백 흡수 — 연말 50여 개 점포, 전국 즉시배송 기반 구축",
                   h=1.00)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out, "| panel bottom:", round(P_BOT, 2))
