# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

논리 흐름이 프레임워크를 결정(SPEC 3-1): 좌 컬럼(현황·배경) → 우 컬럼(추진 방안),
우측 하단 확대 로드맵 타임라인으로 마감. '26.7.10 사용자 수정본 실측 반영:
- 표: 상위 구분(예약배송/퀵커머스) 셀 병합 + 신규 행 연블루 강조
- 결론 박스 제거 → So-What 은 로드맵 하위 불릿으로 통합
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
PITCH = 0.62                 # 불릿 행간


def col_title(slide, l, t, text):
    """컬럼 제목 — 가운데 볼드 14pt + 굵은 밑줄 룰."""
    _txt(slide, l, t, HALF_W, 0.60, text, size=FONT_PT["lead"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(slide, l, t + 0.79, HALF_W, color=INK, w_pt=1.5)


def block(slide, l, t, head, subs):
    """● 주항목 + - 하위 블록 → 종료 y."""
    items = [(0, head)] + [(1, s) for s in subs]
    return add_bullets(slide, items, l=l, t=t, w=HALF_W, line_h=PITCH)


prs = new_deck()
s = add_content(prs, "이마트 퀵커머스", "SSG닷컴 '2시간 배송' 도입 전략",
                tier="P",
                lead="양재·하남점 '2시간 배송' 시범 도입 — 대용량 장보기 공략·점포 물류기지 가동률 제고",
                footnote=SRC)

Y0 = 4.18

# ══ 좌 컬럼 : 현황 및 도입 배경 ══
col_title(s, MARGIN_L, Y0, "현황 및 도입 배경")
block(s, MARGIN_L, 5.45, "퀵커머스 수요 급증", [
    "1시간 오토바이 배송(바로퀵) 매출 197%↑ ('26.6월)",
    "B마트 등 경쟁사도 대용량 품목 확대 추세",
])
block(s, MARGIN_L, 7.83, "이마트 배송 체계", [])
data = [
    ["구분", "서비스명", "피킹·패킹 거점", "운송 수단", "취급 규모"],
    ["예약배송", "주간배송", "센터+점포", "사륜차", "대용량"],
    [None, "새벽배송", "센터", None, None],
    [None, "트레이더스 배송", "점포", None, None],
    ["퀵커머스", "바로퀵 (1시간)", "점포 (반경 3km)", "이륜차", "소량 (1만 종)"],
    [None, "신규 (2시간)", "점포", "사륜차", "대용량"],
]
add_fin_table(s, MARGIN_L, 8.84, HALF_W, 5.34, data,
              col_w=[1.92, 3.13, 2.77, 1.88, 2.55], header_rows=1,
              merges=[(1, 3, 0), (1, 3, 3), (1, 3, 4), (4, 5, 0)],
              hl_rows=[5], bold_cols=(0, 1))
_txt(s, MARGIN_L + 0.47, 14.53, HALF_W - 0.47, LINE_H,
     "→ 사륜 차량 2시간 배송 통한, 서비스 공백 해소 (대용량/즉시)",
     size=FONT_PT["bullet1"], bold=True, color=INK)
block(s, MARGIN_L, 15.63, "규제 제약", [
    "유통산업발전법 → 점포 물류거점(PP센터) 새벽 불가",
])

# ══ 우 컬럼 : 추진 방안 ══
col_title(s, COL2_L, Y0, "추진 방안 : '2시간 배송'")
block(s, COL2_L, 5.45, "시범 운영 ('26. 7. 9.~)", [
    "20시까지 주문 접수 → 결제 후 2시간 내 배송 완료",
    "기존 예약배송 병행 선택 가능",
    "쓱배송 사륜차로 대용량·중량 상품 대응",
])
block(s, COL2_L, 8.50, "서비스 세분화", [
    "1시간 바로퀵(소량) / 2시간(대용량) 이원화",
    "점포 15만 종 구색 기반 (바로퀵의 15배)",
])
block(s, COL2_L, 10.85, "기대 효과", [
    "대용량 즉시배송 서비스 공백 흡수",
    "160여 PP센터 즉시배송 기지화·낮 가동률 제고",
    "규제 완화 시 새벽배송 확장 여지",
])
block(s, COL2_L, 14.02, "확대 로드맵", [
    "연말 50여 개 점포, 전국 즉시배송 기반 구축",
])
add_timeline(s, COL2_L, 15.62, HALF_W, [
    ("'26. 7월", "양재·하남점\n도입"),
    ("8월", "서울 지역\n확대"),
    ("9월", "전국 확대"),
    ("연말", "약 50개 점포\n운영"),
])

# ══ 중앙 세로 구분선 ══
add_vsep(s, VSEP_X, Y0, 17.30 - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
