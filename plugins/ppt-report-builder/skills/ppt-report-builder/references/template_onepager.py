# -*- coding: utf-8 -*-
"""[스타터 템플릿] 원페이저 — 지배 프레임 1개 + 정성 보조. 복사해서 내용만 채우세요.

사용법:  cp _template_onepager.py build_<주제>.py  →  «채우기» 부분만 수정  →  python3 build_<주제>.py

원페이저 철칙(FRAMEWORKS §2-1): **지배 프레임(hero) 1개 + 정성 보조.**
- 이 템플릿의 hero = ② 전폭 근거 표. 결론(So-What)은 리드에 선치 → ① 배경(정성 불릿) →
  ② hero 표(중심) → ③ 실행(가로 타임라인은 **보조·경량**, 생략 가능).
- **한 장에 도식 2종 초과 금지**(표+타임라인이 상한). 플로우+매트릭스+As-Is/To-Be 몽타주 금지(비인간적).
- hero 를 바꾸려면: As-Is/To-Be(전환)·2×2 맵(위치)·추이 차트(변화)로 ②를 교체하고 나머지는 불릿·리본으로.
- 논리가 프레임 하나로 안 담기면 원페이저가 아니라 **덱** — 마디를 슬라이드로(§1 스토리라인).

정밀 기하(내장): 전 객체 동일 폭 BODY_W 끝선 정렬 · 일정 간격 GB · 하단까지 채움 · 각주 직하.
가드레일: 순한글만 · 파란 글자 금지 · 도형 채움색 선택(무채색 OK)·색은 강조만 · 부호 앞 공백 금지 ·
공란(〔조사〕/__)이 화면 지배 금지(질적 내용 충실) · 오버플로 금지. 빌드 후 docs/BUILDER_GUIDE.md §5 QA 필수.
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
tl_bot = add_htimeline(s, MARGIN_L, TBL_BOT + GB + HG + 0.05, BODY_W, [
    ("시점1", "내용1"),
    ("시점2", "내용2\n(부가)"),
    ("시점3", "내용3"),
    ("시점4", "내용4"),
])
# ※ 결론(So-What)은 리드메시지에 선치 — 하단 네이비 결론 박스는 비권장(반복 금지).
#   로드맵이 본문 최종 요소. 하단은 요소 높이·간격으로 채운다.

# ══ 각주(콘텐츠 직하) ══   «채우기 6»
_txt(s, MARGIN_L, tl_bot + 0.30, BODY_W, 0.55, SRC,
     size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "out_onepager.pptx")
prs.save(out)
print("saved", out)
