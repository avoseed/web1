# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 동향 원페이저 — 확정 빌드 사양.

가드레일: ① 순한글만(한자 금지) ② 표는 2×2 포지셔닝 맵 하나뿐(라인업 표 금지)
③ 맵은 3×3 표 객체(파편화 금지) ④ 오버플로 금지.
레이아웃: 헤드라인(전폭) + 좌 45%(도입 배경 + 2×2 맵) / 우 55%(서비스 개요 + 확산 계획).
출처: 비즈워치·머니투데이·한국경제('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 종      ※ 출처 : 비즈워치 · 머니투데이 · 한국경제 ('26.7.8~10)"

# 좌 45% (배경 + 맵) / 우 55% (개요 + 확산)
LW = (BODY_W - COL_GAP) * 0.45
RW = (BODY_W - COL_GAP) * 0.55
COL2_L = MARGIN_L + LW + COL_GAP
VSEP_X = MARGIN_L + LW + COL_GAP / 2
Y0 = 4.18
BOTTOM = 16.35
LEAD_TOP = 15.95

prs = new_deck()
# 헤드라인 = 결론 선치 (순한글: '물류거점화')
s = add_content(prs, "동향", "이마트, 점포 물류거점화로 대용량 즉시배송 선점",
                tier="P",
                lead="'2시간 배송' 양재·하남 시범(7.9~) → 연말 50여 점 · 유휴 PP센터 재활용 저비용 모델",
                footnote=SRC)

# ══════════ 좌 45% : ① 도입 배경 + 2×2 맵 ══════════
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "바로퀵 매출 197%↑ ('26.6월) (미검증)"),
    (1, "(배경) B마트 등 즉시배송 시장 대용량 품목 확대"),
    (1, "기존 라인업 '빠름 × 대용량' 부재"),
], l=MARGIN_L, t=Y0, w=LW, line_h=0.64)

# 2×2 포지셔닝 맵 (3×3 표 객체, 좌측 하단 채움) — 유일한 표
MAP_TOP = y + 0.35
add_matrix2x2(s, MARGIN_L, MAP_TOP, LW, BOTTOM - MAP_TOP,
              col_labels=["소량", "대용량"],
              row_labels=["빠름\n(당일 1~2H)", "예약\n(익일~)"],
              cells=[["바로퀵 (1H)", "★ 2시간 배송 (신규)"],
                     ["—", "주간·새벽·트레이더스"]],
              star=(0, 1), corner="속도＼물량")

# ══════════ 우 55% : ② 서비스 개요 + ③ 확산 계획 ══════════
y = add_bullets(s, [
    (0, "② 서비스 개요"),
    (1, "거점 : 유휴 점포 PP센터 (반경 내 즉시 출고)"),
    (1, "운송 : 쓱배송 사륜차 → 대용량·중량 대응"),
    (1, "운영 : 20시 마감 → 2시간 내 배송 완료 (예약배송 병행)"),
    (1, "규모 : 약 15만 종 (바로퀵 1만 종 대비 15배)"),
], l=COL2_L, t=Y0, w=RW, line_h=0.66)

y = add_bullets(s, [(0, "③ 확산 계획")], l=COL2_L, t=y + 0.30, w=RW, line_h=0.64)
VT_TOP = y + 0.10
add_vtimeline(s, COL2_L, VT_TOP, RW, (LEAD_TOP - 0.15) - VT_TOP, [
    ("'26.7월", "양재·하남점 도입"),
    ("8월", "서울 확대 (월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
])
add_bullets(s, [
    (0, "(리드) 유휴 PP센터 재활용 → 신규 투자 최소 · 저비용 급속 확산"),
], l=COL2_L, t=LEAD_TOP, w=RW, line_h=0.64)

# ══ 중앙 세로 구분선 ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
