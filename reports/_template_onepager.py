# -*- coding: utf-8 -*-
"""[스타터 템플릿] 전폭 상하 구조 원페이저 — 복사해서 내용만 채우세요.

사용법:  cp _template_onepager.py build_<주제>.py  →  «채우기» 부분만 수정  →  python3 build_<주제>.py
프레임워크: 전폭 상하 구조(SPEC 3-2 정본, 근거 표 중심). 6:4 좌우 2단은 폐기.
전폭 밴드 적층: ① 도입 배경 → ② 전폭 근거 표(중심) → ③ 가로 타임라인 → 전폭 네이비 결론.
정밀 기하(내장): 전 객체 동일 폭 BODY_W 끝선 정렬 · 일정 간격 GB · 하단까지 채움 · 각주 직하.
가드레일: 순한글만 · 파란 글자 금지 · 부호 앞 공백 금지 · 오버플로 금지.
빌드 후 docs/BUILDER_GUIDE.md §5 QA 필수.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR

# ── «채우기 1» 각주(출처·취급) ─────────────────────────────────────────
SRC = "※ 출처: (매체·일자)"

Y0 = 4.15
GB = 0.30                              # 객체 간 일정 간격
HG = 0.55                              # 소제목 피치

prs = new_deck()
# ── «채우기 2» 헤드라인(결론 선치) + 리드 ─────────────────────────────
s = add_content(prs, "동향", "«명사형 결론 한 줄 제목»", tier="P",
                lead="«신설 시점 · 규모 · 핵심 성격 요약 한 줄»")

# ══ ① 도입 배경 (전폭 불릿) ══   «채우기 3»
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "핵심 지표·인과 1 → 함의"),
    (1, "핵심 지표·인과 2 → 공백·문제 명시"),
    (1, "핵심 지표·인과 3 → 대응 방향"),
], l=MARGIN_L, t=Y0, w=BODY_W, line_h=0.54)

# ══ ② 근거 표 (전폭 — 장표 중심) ══   «채우기 4»
add_bullets(s, [(0, "② 표 소제목")], l=MARGIN_L, t=y + GB, w=BODY_W, line_h=HG)
TBL_TOP = y + GB + HG + 0.06
TBL_BOT = 13.35                        # 표 하단(중심 객체 크게); 이후 타임라인·결론이 하단 채움
tbl_data = [
    ["구분", "열2", "열3", "열4", "열5", "열6"],
    ["그룹A", "행1", "-", "-", "-", "-"],
    [None, "행2", "-", "-", "-", "-"],
    [None, "행3", "-", "-", "-", "-"],
    ["그룹B", "행4", "-", "-", "-", "-"],
    [None, "★ 신규(강조)", "-", "-", "-", "-"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, BODY_W, TBL_BOT - TBL_TOP, tbl_data,
              col_w=[2.6, 5.4, 2.2, 5.2, 2.4, BODY_W - 17.8], header_rows=1,
              merges=[(1, 3, 0), (4, 5, 0)], hl_rows=[5], bold_cols=(0, 1))

# ══ ③ 확산 로드맵 (전폭 가로 타임라인) ══   «채우기 5»
add_bullets(s, [(0, "③ 로드맵")], l=MARGIN_L, t=TBL_BOT + GB, w=BODY_W, line_h=HG)
add_htimeline(s, MARGIN_L, TBL_BOT + GB + HG + 0.05, BODY_W, [
    ("시점1", "내용1"),
    ("시점2", "내용2\n(부가)"),
    ("시점3", "내용3"),
    ("시점4", "내용4"),
])

# ══ So What (전폭 네이비 결론) + 각주 직하 ══   «채우기 6»
CONCL_TOP = 16.00
add_conclusion_box(s, MARGIN_L, CONCL_TOP, BODY_W, "«So What 결론 한 줄»", h=0.95)
_txt(s, MARGIN_L, CONCL_TOP + 0.95 + 0.30, BODY_W, 0.55, SRC,
     size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "out_onepager.pptx")
prs.save(out)
print("saved", out)
