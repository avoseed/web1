# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 보고서 — ZETTA 표준 PPT 빌더(v4.1) 기반.

'26.7.10 최종 본문 반영:
- 제목 [동향] 분류, 부제(리드)
- 좌 「현황 및 배경」: 수요 급증 → 서비스 이원화 → 이마트 배송 체계(표, 병합·강조)
- 우 「서비스 세부 내용」: 운영 방침 → 확산 로드맵(타임라인 + 리드 불릿)
불릿 체계(•/-)·표 양식·괄호 서식·좌우 하단 정렬 표준 유지.
출처: 비즈워치·머니투데이·한국경제('26. 7. 8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, _rect
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 종      ※ 출처 : 비즈워치 · 머니투데이 · 한국경제 ('26. 7. 8~10)"
HALF_W = (BODY_W - COL_GAP) / 2
COL2_L = MARGIN_L + HALF_W + COL_GAP
VSEP_X = MARGIN_L + HALF_W + COL_GAP / 2
PITCH = 0.62
BOTTOM = 16.55          # 좌·우 컬럼 공통 하단선


def col_title(slide, l, t, text):
    """컬럼 제목 — 가운데 볼드 14pt + 굵은 밑줄 룰."""
    _txt(slide, l, t, HALF_W, 0.60, text, size=FONT_PT["lead"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(slide, l, t + 0.79, HALF_W, color=INK, w_pt=1.5)


def block(slide, l, t, head, subs, pitch=PITCH):
    """● 주항목 + - 하위 블록 → 종료 y."""
    items = [(0, head)] + [(1, x) for x in subs]
    return add_bullets(slide, items, l=l, t=t, w=HALF_W, line_h=pitch)


prs = new_deck()
s = add_content(prs, "동향", "이마트 퀵커머스 '2시간 배송' 서비스 도입",
                tier="P",
                lead="양재·하남점 시범 도입 — 점포 물류거점화로 대용량 즉시배송 공백 공략",
                footnote=SRC)

Y0 = 4.18

# ══ 좌 컬럼 : 현황 및 배경 ══
col_title(s, MARGIN_L, Y0, "현황 및 배경")
y = block(s, MARGIN_L, 5.45, "퀵커머스 수요 급증", [
    "1시간 오토바이 배송(바로퀵) 매출 197%↑ ('26.6월) (미검증)",
    "(배경) B마트 등 즉시배송 시장 전반 대용량 품목 확대 추세",
])
y = block(s, MARGIN_L, y + 0.30, "퀵커머스 서비스 이원화", [
    "기존 바로퀵, 소량·즉석식 한정 → 대용량 즉시배송 수요 미충족",
    "1시간 배송(소량) + 2시간 배송(대용량) 이원화",
    "점포 15만 종 구색 기반 (바로퀵 1만 종 대비 15배)",
])
y = block(s, MARGIN_L, y + 0.30, "이마트 배송 체계", [])
TBL_TOP = y + 0.05
TBL_W = 11.85          # < HALF_W, 제목 밑줄 안쪽 + 중앙 세로선 여백
data = [
    ["구분", "서비스명", "피킹·패킹 거점", "운송 수단", "취급 규모"],
    ["예약배송", "주간배송", "센터+점포", "사륜차", "대용량"],
    [None, "새벽배송", "센터", None, None],
    [None, "트레이더스 배송", "점포", None, None],
    ["퀵커머스", "1시간 배송(바로퀵)", "점포(반경 3km)", "이륜차", "소량(1만 종)"],
    [None, "2시간 배송", "점포", "사륜차", "대용량"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, TBL_W, BOTTOM - TBL_TOP, data,
              col_w=[1.85, 3.00, 2.65, 1.80, 2.55], header_rows=1,
              merges=[(1, 3, 0), (1, 3, 3), (1, 3, 4), (4, 5, 0)],
              hl_rows=[5], bold_cols=(0, 1))

# ══ 우 컬럼 : 서비스 세부 내용 ══
col_title(s, COL2_L, Y0, "서비스 세부 내용")
block(s, COL2_L, 5.45, "운영 방침 ('26. 7. 9.~)", [
    "20시까지 주문 접수 → 2시간 내 배송 완료",
    "쓱배송 사륜차 활용 → 대용량·중량 상품 대응",
    "기존 예약배송 병행 선택 가능",
    "(리드) 기존 점포·인력 재활용형 서비스 → 신규 자산 투자 최소",
], pitch=0.90)

block(s, COL2_L, 10.30, "확산 로드맵", [])
# 로드맵 타임라인 (프레임 박스)
_rect(s, COL2_L, 11.00, HALF_W, 3.87, fill=WHITE, line=GRID, line_w=1.0)
add_timeline(s, COL2_L + 0.35, 11.88, HALF_W - 0.70, [
    ("'26. 7월", "양재·하남"),
    ("8월", "서울\n(월계·가든5·신도림)"),
    ("9월", "전국"),
    ("연말", "50여 점포"),
])
# 로드맵 하단 리드 불릿 (• 주항목)
add_bullets(s, [
    (0, "(리드) 점포를 물류거점으로 재정의, 유휴 PP센터 가동률 활용"),
    (1, "→ 연말 50여 점 본격 확산 (전국 즉시배송 기반 구축)"),
], l=COL2_L, t=15.37, w=HALF_W, line_h=0.78)

# ══ 중앙 세로 구분선 (하단 정렬선까지) ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
