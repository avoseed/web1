# -*- coding: utf-8 -*-
"""[스타터 템플릿] 6:4 본장형 원페이저 — 복사해서 내용만 채우세요.

사용법:  cp _template_onepager.py build_<주제>.py  →  아래 «채우기» 부분만 수정  →  python3 build_<주제>.py
프레임워크: 정기협의체 6:4 본장형(SPEC 3-2). 좌 60% 현황·구조 / 우 40% 전개.
정밀 기하(내장): ① 좌 전 객체 동일 폭 LW 끝선 정렬 ② 객체 간격 GB 통일
③ 표 끝선-중앙 세로선 대칭 여백(GAP=1.10 로 확보, 침범 금지) ④ 하단까지 채움 + 각주 직하.
가드레일: 순한글만(한자 금지) · 파란 글자 금지 · 부호 앞 공백 금지 · 오버플로 금지.
빌드 후 docs/BUILDER_GUIDE.md §5 QA 필수.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR

# ── «채우기 1» 각주(출처·취급) ─────────────────────────────────────────
SRC = "※ 출처: (매체·일자)"

# ── 레이아웃 파라미터 (보통 그대로 사용) ──────────────────────────────
GAP = 1.10                             # 컬럼 간격 — 표 끝선과 중앙 세로선 대칭 여백(0.55) 확보
LW = (BODY_W - GAP) * 0.60             # 좌 60%
RW = (BODY_W - GAP) * 0.40             # 우 40%
COL2_L = MARGIN_L + LW + GAP
VSEP_X = MARGIN_L + LW + GAP / 2
Y0 = 4.15                              # 본문 시작(리드 하단)
GB = 0.32                              # 객체 간 일정 간격
LN = 0.58                              # 불릿 행 피치(하단까지 분산 채움; 컴팩트 원하면 0.50)
HG = 0.55                              # 소제목 행 피치
BOTTOM = 17.00                         # 좌 표 하단 = 우 결론 박스 하단(하단까지 채움 모드)
FN_Y = BOTTOM + 0.48                   # 각주는 콘텐츠 직하(하단 여백 방지)
#  ※ 컴팩트 모드로 바꾸려면: BOTTOM 을 콘텐츠에 맞춰 낮추고 FN_Y=BOTTOM+0.45 유지.

prs = new_deck()
# ── «채우기 2» 헤드라인(결론 선치) + 리드 ─────────────────────────────
s = add_content(prs, "동향", "«명사형 결론 한 줄 제목»", tier="P",
                lead="«신설 시점 · 규모 · 핵심 성격 요약 한 줄»")

# ══════════ 좌 60% : 「현황 및 배경」 ══════════
add_col_header(s, MARGIN_L, Y0, LW, "현황 및 배경")
y = Y0 + COL_HEADER_H + GB
# «채우기 3» 현황 불릿 (0=•주항목 / 1=-하위). 하단까지 차도록 소제목 2개 + 각 3줄 권장.
y = add_bullets(s, [
    (0, "① 소제목 A"),
    (1, "핵심 지표·근거 1"),
    (1, "핵심 지표·근거 2"),
    (1, "핵심 지표·근거 3"),
], l=MARGIN_L, t=y, w=LW, line_h=LN)
y = add_bullets(s, [
    (0, "② 소제목 B"),
    (1, "핵심 지표·근거 1"),
    (1, "핵심 지표·근거 2"),
    (1, "핵심 지표·근거 3"),
], l=MARGIN_L, t=y + GB, w=LW, line_h=LN)
add_bullets(s, [(0, "③ 표 소제목")], l=MARGIN_L, t=y + GB, w=LW, line_h=HG)
TBL_TOP = y + GB + HG + 0.06
# «채우기 4» 담백한 그리드 표 — 병합 연속 셀은 None, 강조 행은 hl_rows.
tbl_data = [
    ["구분", "열2", "열3", "열4", "열5"],
    ["그룹A", "행1", "-", "-", "-"],
    [None, "행2", "-", "-", "-"],
    ["그룹B", "행3", "-", "-", "-"],
    [None, "★ 신규(강조)", "-", "-", "-"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, LW, BOTTOM - TBL_TOP, tbl_data,
              col_w=[1.6, 4.0, 3.1, 1.4, LW - 10.1], header_rows=1,
              merges=[(1, 2, 0), (3, 4, 0)], hl_rows=[4], font_size=9)

# ══════════ 우 40% : 「서비스 세부 내용」 ══════════
add_col_header(s, COL2_L, Y0, RW, "서비스 세부 내용")
y = Y0 + COL_HEADER_H + GB
# «채우기 5» 전개 불릿
y = add_bullets(s, [
    (0, "① 소제목"),
    (1, "항목 1"),
    (1, "항목 2"),
    (1, "항목 3"),
    (1, "항목 4"),
], l=COL2_L, t=y, w=RW, line_h=LN)
add_bullets(s, [(0, "② 로드맵")], l=COL2_L, t=y + GB, w=RW, line_h=HG)
VT_TOP = y + GB + HG + 0.06
CONCL_TOP = BOTTOM - 1.15
# «채우기 6» 세로 타임라인(시점, 내용). 긴 내용은 \n 으로 2줄.
add_vtimeline(s, COL2_L, VT_TOP, RW, (CONCL_TOP - 0.22) - VT_TOP + 0.30, [
    ("시점1", "내용1"),
    ("시점2", "내용2\n(부가)"),
    ("시점3", "내용3"),
    ("시점4", "내용4"),
], max_gap=2.4)
# «채우기 7» So What 네이비 결론 한 줄
add_conclusion_box(s, COL2_L, CONCL_TOP, RW, "«So What 결론 한 줄»", h=1.15)

# ══ 중앙 세로 구분선 + 각주(콘텐츠 직하) ══
add_vsep(s, VSEP_X, Y0, BOTTOM - Y0)
_txt(s, MARGIN_L, FN_Y, BODY_W, 0.55, SRC, size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "out_onepager.pptx")
prs.save(out)
print("saved", out)
