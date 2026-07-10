# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 — ZETTA 표준 빌더(v4.1) · 문제·해결 대응형(v3).

지배 논리: 헤드라인=결론 선치 + 3막(도입 배경 → 서비스 개요 → 확산 계획).
핵심 장치: 좌(기존 라인업의 공백=문제) ↔ 우(신규 서비스의 채움=해결) 대응.
앵커 시각물: ② 서비스 개요 중심의 2×2 포지셔닝 맵(속도×물량, ★ 신규).
레이아웃: 헤드라인(전폭) + 좌 40%(배경·라인업 표) / 우 60%(맵·스펙·로드맵).
섹션 제목은 실무 명사형(도입 배경·서비스 개요·확산 계획)만.
출처: 비즈워치·머니투데이·한국경제('26. 7. 8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 종      ※ 출처 : 비즈워치 · 머니투데이 · 한국경제 ('26. 7. 8~10)"

# 좌 40% (배경·문제) / 우 60% (서비스·해결)
LW = (BODY_W - COL_GAP) * 0.40          # 9.80
RW = (BODY_W - COL_GAP) * 0.60          # 14.70
COL2_L = MARGIN_L + LW + COL_GAP
VSEP_X = MARGIN_L + LW + COL_GAP / 2
Y0 = 4.18
BOTTOM = 16.45


prs = new_deck()
# 헤드라인 = 결론(So What) 선치 + 부제(시점·규모·성격)
s = add_content(prs, "동향", "이마트, 점포 물류거점化로 대용량 즉시배송 선점",
                tier="P",
                lead="'2시간 배송' 양재·하남 시범(7.9~) → 연말 50여 점 · 유휴 PP센터 재활용 저비용 모델",
                footnote=SRC)

# ══════════ 좌 40% : ① 도입 배경 (문제) ══════════
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "바로퀵 매출 197%↑ ('26.6월) (미검증)"),
    (1, "(배경) B마트 등 대용량 품목 확대"),
    (1, "기존 라인업 '빠름 × 대용량' 부재"),
], l=MARGIN_L, t=Y0, w=LW, line_h=0.64)

# 기존 라인업 표 — 공백 드러냄 (속도·물량 축, 빠름+대용량 부재 가시화)
lu = [
    ["서비스명", "속도", "물량"],
    ["주간배송", "예약", "대용량"],
    ["새벽배송", "예약(새벽)", "대용량"],
    ["트레이더스", "예약", "대용량"],
    ["바로퀵", "빠름(1H)", "소량"],
    ["✗ 공백", "빠름", "대용량"],
]
LU_TOP = y + 0.35
add_fin_table(s, MARGIN_L, LU_TOP, LW, BOTTOM - LU_TOP, lu,
              col_w=[3.30, 3.20, LW - 6.50], header_rows=1, hl_rows=[5])

# ══════════ 우 60% : ② 서비스 개요 (해결) + ③ 확산 계획 ══════════
y = add_bullets(s, [(0, "② 서비스 개요")], l=COL2_L, t=Y0, w=RW, line_h=0.64)

# 앵커 — 2×2 포지셔닝 맵 (속도 × 물량, 빈 사분면 ★ 신규)
MAP_TOP = y + 0.15
add_matrix2x2(s, COL2_L, MAP_TOP, RW, 4.60,
              col_labels=["소량", "대용량"],
              row_labels=["빠름\n(당일 1~2H)", "예약\n(익일~)"],
              cells=[["바로퀵 (1H)", "★ 2시간 배송 (신규)"],
                     ["—", "주간·새벽·트레이더스"]],
              star=(0, 1), corner="속도 ↓ / 물량 →")

# 핵심 스펙 4종 (신규 서비스가 무엇을 가능케 하나)
y = add_bullets(s, [
    (1, "거점 유휴 점포 · 운송 사륜차 · 운영 20시 마감 2시간 · 규모 15만 종"),
], l=COL2_L, t=MAP_TOP + 4.60 + 0.15, w=RW, line_h=0.62)

# ③ 확산 계획 — 세로 타임라인으로 우측 세로 공간 채움
y = add_bullets(s, [(0, "③ 확산 계획")], l=COL2_L, t=y + 0.28, w=RW, line_h=0.64)
vt_bottom = add_vtimeline(s, COL2_L, y + 0.10, RW, 4.35, [
    ("'26. 7월", "양재·하남점 도입"),
    ("8월", "서울 확대 (월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
])
# 본질 리드 1줄 (So What 재확인, 별도 시사점 블록 없음)
add_bullets(s, [
    (0, "(리드) 유휴 PP센터 재활용 → 신규 투자 최소 · 저비용 급속 확산"),
], l=COL2_L, t=vt_bottom + 0.15, w=RW, line_h=0.64)

# ══ 중앙 세로 구분선 ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
