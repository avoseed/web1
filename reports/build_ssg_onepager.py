# -*- coding: utf-8 -*-
"""SSG닷컴 '2시간 배송' 원페이지 — ZETTA 표준 빌더(v4.1) · 6:4 비대칭 레이아웃.

레이아웃(경쟁사 동향 6:4형): 좌 60% 현황·구조 3블록 / 우 40% 전개 2블록.
- 좌 ① 수요·배경 ② 서비스 구조 + 신·구 비교표(성격 차이) ③ 전체 체계 표(라인업 위치)
- 우 ① 운영 방침 ② 확산 로드맵(세로 타임라인) + So What 리드 1줄
두 표는 관점 분리: ② '왜 나눴나(성격)' / ③ '어디에 있나(위치)'.
출처: 비즈워치·머니투데이·한국경제('26. 7. 8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt
from pptx.enum.text import PP_ALIGN

SRC = "※ 취급 품목 : 이마트 15만 · 바로퀵 1만 종      ※ 출처 : 비즈워치 · 머니투데이 · 한국경제 ('26. 7. 8~10)"

# ── 6:4 비대칭 2단 ──
LW = (BODY_W - COL_GAP) * 0.60          # 좌 14.70
RW = (BODY_W - COL_GAP) * 0.40          # 우 9.80
COL2_L = MARGIN_L + LW + COL_GAP        # 우 컬럼 좌측
VSEP_X = MARGIN_L + LW + COL_GAP / 2
Y0 = 4.18
BOTTOM = 16.75
PITCH = 0.62


def col_title(l, t, w, text):
    _txt(s, l, t, w, 0.60, text, size=FONT_PT["lead"], bold=True,
         color=INK, align=PP_ALIGN.CENTER)
    add_hline(s, l, t + 0.79, w, color=INK, w_pt=1.5)


def lblock(t, head, subs, pitch=PITCH):
    items = [(0, head)] + [(1, x) for x in subs]
    return add_bullets(s, items, l=MARGIN_L, t=t, w=LW, line_h=pitch)


def rblock(t, head, subs, pitch=PITCH):
    items = [(0, head)] + [(1, x) for x in subs]
    return add_bullets(s, items, l=COL2_L, t=t, w=RW, line_h=pitch)


prs = new_deck()
s = add_content(prs, "동향", "이마트 퀵커머스 '2시간 배송' 서비스 도입",
                tier="P",
                lead="양재·하남점 시범 도입 — 점포 물류거점화로 대용량 즉시배송 공백 공략",
                footnote=SRC)

# ══════════ 좌 60% : 현황 · 구조 ══════════
col_title(MARGIN_L, Y0, LW, "현황 및 구조")

# ① 수요 / 배경
y = lblock(5.45, "퀵커머스 수요 급증", [
    "1시간 오토바이 배송(바로퀵) 매출 197%↑ ('26.6월) (미검증)",
    "(배경) B마트 등 즉시배송 시장 전반 대용량 품목 확대 추세",
])

# ② 서비스 구조 + 신·구 비교표 (성격 차이 : 왜 나눴나)
y = lblock(y + 0.28, "퀵커머스 서비스 이원화", [
    "기존 바로퀵, 소량·즉석식 한정 → 대용량 즉시배송 공백",
])
cmp_data = [
    ["구분", "1시간 배송(바로퀵)", "2시간 배송(신규)"],
    ["운송 수단", "이륜차", "사륜차"],
    ["취급 품목", "약 1만 종", "약 15만 종"],
    ["주 용도", "소량·즉석식", "대용량 장보기"],
]
add_fin_table(s, MARGIN_L, y + 0.10, LW - 0.55, 3.05, cmp_data,
              col_w=[3.05, 5.55, LW - 0.55 - 8.60], header_rows=1,
              hl_cols=[2], bold_cols=(0,))

# ③ 전체 체계 표 (라인업 위치 : 어디에 있나)
y3 = y + 0.10 + 3.05 + 0.40
lblock(y3, "이마트 배송 라인업", [])
sys_data = [
    ["구분", "서비스명", "피킹·패킹 거점", "취급 규모"],
    ["예약배송", "주간배송", "센터+점포", "대용량"],
    [None, "새벽배송", "센터", None],
    [None, "트레이더스 배송", "점포", None],
    ["퀵커머스", "1시간 배송(바로퀵)", "점포(반경 3km)", "소량"],
    [None, "2시간 배송(신규)", "점포", "대용량"],
]
TBL3_TOP = y3 + 0.72
add_fin_table(s, MARGIN_L, TBL3_TOP, LW - 0.55, BOTTOM - TBL3_TOP, sys_data,
              col_w=[2.10, 4.60, 4.30, LW - 0.55 - 11.00], header_rows=1,
              merges=[(1, 3, 0), (1, 3, 3), (4, 5, 0)],
              hl_rows=[5], bold_cols=(0, 1))

# ══════════ 우 40% : 전개 ══════════
col_title(COL2_L, Y0, RW, "전개 방향")

# ① 운영 방침
rblock(5.45, "운영 방침 ('26. 7. 9.~)", [
    "20시 마감 → 2시간 내 배송 완료",
    "쓱배송 사륜차 → 대용량·중량 대응",
    "기존 예약배송 병행 선택 가능",
    "점포·인력 재활용 → 신규 투자 최소",
], pitch=0.70)

# ② 확산 로드맵 (세로 타임라인)
rblock(9.30, "확산 로드맵", [])
vt_bottom = add_vtimeline(s, COL2_L, 10.05, RW, 5.20, [
    ("'26. 7월", "양재·하남점 도입"),
    ("8월", "서울 확대\n(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
])
# So What 리드 1줄 (별도 블록 없이 로드맵 하단 마감)
add_bullets(s, [
    (0, "(리드) 점포를 물류거점으로 재정의"),
    (1, "→ 유휴 PP센터 가동률 활용, 전국 즉시배송 기반 구축"),
], l=COL2_L, t=vt_bottom + 0.25, w=RW, line_h=0.70)

# ══ 중앙 세로 구분선 ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
