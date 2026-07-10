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
PITCH = 0.62                 # 불릿 행간 (수직 균형 조정)
GAP_BLK = 0.40               # 블록 간 간격 (좌 컬럼)
GAP_BLK_R = 0.35             # 블록 간 간격 (우 컬럼)


def col_title(slide, l, t, text):
    """컬럼 제목 — 가운데 볼드 + 굵은 밑줄 룰 (15/15 실측)."""
    _txt(slide, l, t, HALF_W, 0.55, text, size=FONT_PT["bullet0"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(slide, l, t + 0.62, HALF_W, color=INK, w_pt=1.5)
    return t + 1.15          # 룰 하단 +0.53 여백 (제목 아래는 넉넉히)


def block(slide, l, t, head, subs):
    """● 주항목 + - 하위 블록 → 종료 y."""
    items = [(0, head)] + [(1, s) for s in subs]
    return add_bullets(slide, items, l=l, t=t, w=HALF_W, line_h=PITCH)


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·점포 물류기지 가동률 제고",
                footnote=SRC)

Y0 = BODY_TOP_LEAD

# ══ 좌 컬럼 : 현황 및 도입 배경 (문제 정의) ══
y = col_title(s, MARGIN_L, Y0, "현황 및 도입 배경")
y = block(s, MARGIN_L, y, "퀵커머스 수요 급증", [
    "1시간 오토바이 배송(바로퀵) 매출 197%↑(6월)",
    "B마트 등 경쟁사도 대용량 품목 확대 추세",
])
y += 0.30
y = block(s, MARGIN_L, y, "현행 배송 체계", [])
data = [
    ["구분", "배송 형태", "피킹·패킹 거점", "운송 수단", "취급 규모"],
    ["주간배송", "예약", "'네오' +\n이마트 점포", "사륜차", "대용량 가능"],
    ["새벽배송", "예약 (새벽)", "'네오' 한정\n(점포 불가)", "사륜차", "대용량 가능"],
    ["트레이더스배송", "예약", "트레이더스 점포", "사륜차", "대용량 가능"],
    ["바로퀵", "즉시\n(1시간)", "이마트 점포\n(반경 3km)", "이륜차", "소량 (1만 종)"],
]
add_fin_table(s, MARGIN_L, y, HALF_W, 5.0, data,
              col_w=[2.75, 2.30, 3.05, 1.60, HALF_W - 9.70], header_rows=1)
y += 5.0 + 0.30
_txt(s, MARGIN_L + 0.20, y, HALF_W - 0.20, LINE_H,
     "→ 사륜 차량 2시간 배송 통한, 서비스 공백 해소 (대용량/즉시)",
     size=FONT_PT["bullet1"], bold=False, color=INK)
y += PITCH + 0.30
block(s, MARGIN_L, y, "규제 제약", [
    "유통산업발전법 → 점포 물류거점(PP센터) 새벽 불가",
])

# ══ 우 컬럼 : 추진 방안 (해법) ══
y = col_title(s, COL2_L, Y0, "추진 방안 : '2시간 배송'")
y = block(s, COL2_L, y, "시범 운영 ('26. 7. 9.~)", [
    "20시까지 주문 접수 → 결제 후 2시간 내 배송 완료",
    "기존 예약배송 병행 선택 가능",
    "쓱배송 사륜차로 대용량·중량 상품 대응",
])
y += GAP_BLK_R
y = block(s, COL2_L, y, "서비스 세분화", [
    "1시간 바로퀵(소량) / 2시간(대용량) 이원화",
    "점포 15만 종 구색 기반 (바로퀵의 15배)",
])
y += GAP_BLK_R
y = block(s, COL2_L, y, "기대 효과", [
    "대용량 즉시배송 서비스 공백 흡수",
    "160여 PP센터 즉시배송 기지화·낮 가동률 제고",
    "규제 완화 시 새벽배송 확장 여지",
])
y += GAP_BLK_R
y = block(s, COL2_L, y, "확대 로드맵", [])
add_timeline(s, COL2_L, y + 0.05, HALF_W, [
    ("'26. 7월", "양재·하남"),
    ("8월", "서울 확대"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포"),
])

# ══ 중앙 세로선 + 전폭 결론 박스 (하단 앵커) ══
add_vsep(s, VSEP_X, Y0, 15.60 - Y0)
add_conclusion_box(s, MARGIN_L, 15.85, BODY_W,   # 본문 하단 +0.45 밀착
                   "사륜차 대용량 즉시배송으로 퀵커머스 공백 흡수 — 연말 50여 개 점포, 전국 즉시배송 기반 구축",
                   h=1.00)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
