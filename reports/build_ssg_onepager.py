# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 동향 원페이저 — 정기협의체 본장형(6:4 경쟁사 동향, SPEC 3-2).

본장 규율: 컬럼 헤더 바 + 담백한 그리드 표 + 중앙 세로선 + 세로 타임라인 + 네이비 결론.
정밀 기준(정기협의체): ① 객체 끝선 정렬(좌 전 객체 동일 폭 LW) ② 객체간 일정 간격(GB)
③ 표는 중앙 세로선과 좌우 대칭 여백 확보(침범 금지) ④ 표 행 높이 조밀(밀도).
좌 60% 「현황 및 배경」: ① 수요 급증 → ② 서비스 이원화 → ③ 배송 체계 표(그룹 병합·2시간 강조).
우 40% 「서비스 세부 내용」: ① 운영 방침 → ② 확산 로드맵(세로 타임라인) → So What 네이비 결론.
좌 표 하단 = 우 결론 박스 하단 동일선.
출처: 비즈워치·머니투데이·한국경제('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR
from pptx.util import Cm

SRC = ("※ 취급 품목: 이마트 15만 · 바로퀵 1만 종      "
       "※ 출처: 비즈워치 · 머니투데이 · 한국경제 ('26.7.8~10)")

# 좌 60% / 우 40% — 컬럼 간격 GAP 넓혀 표 끝선과 중앙 세로선 좌우 대칭 여백(0.55) 확보
GAP = 1.10
LW = (BODY_W - GAP) * 0.60
RW = (BODY_W - GAP) * 0.40
COL2_L = MARGIN_L + LW + GAP
VSEP_X = MARGIN_L + LW + GAP / 2
Y0 = 4.15
GB = 0.28                              # 객체간 일정 간격
LN = 0.50                              # 불릿 행 피치
BOTTOM = 15.15                         # 좌 표 하단 = 우 결론 박스 하단 동일선

prs = new_deck()
# ══ 헤드라인(전폭) = 결론 선치 ══
s = add_content(prs, "동향", "이마트, 점포 물류거점화로 대용량 즉시배송 선점", tier="P",
                lead="'2시간 배송' 양재·하남 시범(7.9~) → 연말 50여 점 · 유휴 PP센터 재활용 저비용 모델",
                footnote=SRC)

# ══════════ 좌 60% : 「현황 및 배경」 ══════════
add_col_header(s, MARGIN_L, Y0, LW, "현황 및 배경")
y = Y0 + COL_HEADER_H + 0.30
y = add_bullets(s, [
    (0, "① 수요 급증"),
    (1, "바로퀵 매출 197%↑ ('26.6월) (미검증)"),
    (1, "B마트 등 즉시배송 시장, 대용량 품목 확대"),
], l=MARGIN_L, t=y, w=LW, line_h=LN)
y = add_bullets(s, [
    (0, "② 서비스 이원화"),
    (1, "바로퀵은 소량·즉석식 한정 → 대용량 미충족"),
    (1, "1시간(소량) / 2시간(대용량) 이원화, 15만 종 구색"),
], l=MARGIN_L, t=y + GB, w=LW, line_h=LN)
add_bullets(s, [(0, "③ 이마트 배송 체계")], l=MARGIN_L, t=y + GB, w=LW, line_h=0.52)
TBL_TOP = y + GB + 0.60
tbl_data = [
    ["구분", "서비스", "시간 · 형태", "운송", "취급"],
    ["예약배송", "주간배송", "구간당 4~5H", "사륜차", "점포 상품"],
    [None, "새벽배송", "새벽 (PP 불가)", "사륜차", "네오 상품"],
    [None, "트레이더스배송", "예약", "사륜차", "트레이더스"],
    ["퀵커머스", "바로퀵", "1H (반경 3km)", "이륜차", "약 1만 종"],
    [None, "★ 2시간 배송 (신규)", "2H (20시 마감)", "사륜차", "약 15만 종"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, LW, BOTTOM - TBL_TOP, tbl_data,
              col_w=[1.6, 4.0, 3.1, 1.4, LW - 10.1], header_rows=1,
              merges=[(1, 3, 0), (4, 5, 0)], hl_rows=[5], font_size=9)

# ══════════ 우 40% : 「서비스 세부 내용」 ══════════
add_col_header(s, COL2_L, Y0, RW, "서비스 세부 내용")
y = Y0 + COL_HEADER_H + 0.30
y = add_bullets(s, [
    (0, "① 운영 방침"),
    (1, "20시 마감 → 2시간 내 배송 완료"),
    (1, "쓱배송 사륜차 (대용량·중량 대응)"),
    (1, "기존 예약배송 병행"),
    (1, "유휴 PP센터 재활용형 서비스"),
], l=COL2_L, t=y, w=RW, line_h=0.52)
add_bullets(s, [(0, "② 확산 로드맵")], l=COL2_L, t=y + GB, w=RW, line_h=0.52)
VT_TOP = y + GB + 0.60
CONCL_TOP = BOTTOM - 1.15
add_vtimeline(s, COL2_L, VT_TOP, RW, (CONCL_TOP - 0.20) - VT_TOP + 0.30, [
    ("'26.7월", "양재·하남점 도입"),
    ("8월", "서울 확대\n(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
], max_gap=1.85)
add_conclusion_box(s, COL2_L, CONCL_TOP, RW,
                   "유휴 PP센터 재활용 → 신규 투자 최소 · 저비용 급속 확산", h=1.15)

# ══ 중앙 세로 구분선 ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
