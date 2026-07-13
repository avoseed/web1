# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 동향 원페이저 — 전폭 상하 구조(근거 표 중심, SPEC 3-1 B형).

6:4 좌우 2단 폐기 → 전폭 밴드 적층: ① 도입 배경 → ② 배송 체계 표(전폭 중심) → ③ 로드맵 + 결론.
정밀 기하: 전 객체 동일 폭(BODY_W) 끝선 정렬 · 일정 간격(GB) · 하단까지 채움 · 각주 직하.
가드레일: 순한글만 · 파란 글자 금지 · 부호 앞 공백 금지 · 오버플로 금지.
출처: 비즈워치·머니투데이·한국경제('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR

SRC = ("※ 취급 품목: 이마트 15만 · 바로퀵 1만 종      "
       "※ 출처: 비즈워치 · 머니투데이 · 한국경제 ('26.7.8~10)")

Y0 = 4.15
GB = 0.30                              # 객체 간 일정 간격
HG = 0.55                              # 소제목 피치

prs = new_deck()
# ══ 헤드라인(전폭) = 결론 선치 ══
s = add_content(prs, "동향", "이마트, 점포 물류거점화로 대용량 즉시배송 선점", tier="P",
                lead="'2시간 배송' 양재·하남 시범(7.9~) → 연말 50여 점 · 유휴 PP센터 재활용 저비용 모델")

# ══ ① 도입 배경 (전폭) ══
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "바로퀵 매출 197%↑ ('26.6월) (미검증) → 온디맨드 대용량 수요 실증"),
    (1, "B마트 등 즉시배송 시장, 대용량 품목 확대 → 소량 한정의 공백"),
    (1, "1시간(소량·이륜차) / 2시간(대용량·사륜차) 이원화로 대응"),
], l=MARGIN_L, t=Y0, w=BODY_W, line_h=0.54)

# ══ ② 이마트 배송 체계 (전폭 근거 표 — 중심) ══
add_bullets(s, [(0, "② 이마트 배송 체계")], l=MARGIN_L, t=y + GB, w=BODY_W, line_h=HG)
TBL_TOP = y + GB + HG + 0.06
TBL_BOT = 13.35
tbl_data = [
    ["구분", "서비스", "형태", "시간", "운송", "취급"],
    ["예약배송", "주간배송", "예약", "구간당 4~5H", "사륜차", "점포 상품"],
    [None, "새벽배송", "예약", "새벽 (PP 불가)", "사륜차", "네오 상품"],
    [None, "트레이더스배송", "예약", "익일", "사륜차", "트레이더스"],
    ["퀵커머스", "바로퀵", "즉시", "1H (반경 3km)", "이륜차", "약 1만 종"],
    [None, "★ 2시간 배송 (신규)", "즉시", "2H (20시 마감)", "사륜차", "약 15만 종"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, BODY_W, TBL_BOT - TBL_TOP, tbl_data,
              col_w=[2.6, 5.4, 2.2, 5.2, 2.4, BODY_W - 17.8], header_rows=1,
              merges=[(1, 3, 0), (4, 5, 0)], hl_rows=[5], bold_cols=(0, 1))

# ══ ③ 확산 로드맵 (전폭 가로 타임라인) ══
add_bullets(s, [(0, "③ 확산 로드맵")], l=MARGIN_L, t=TBL_BOT + GB, w=BODY_W, line_h=HG)
add_htimeline(s, MARGIN_L, TBL_BOT + GB + HG + 0.05, BODY_W, [
    ("'26.7월", "양재·하남점\n도입"),
    ("8월", "서울 확대\n(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포\n운영"),
])

# ══ So What (전폭 네이비 결론) + 각주 직하 ══
CONCL_TOP = 16.00
add_conclusion_box(s, MARGIN_L, CONCL_TOP, BODY_W,
                   "유휴 PP센터 재활용 → 신규 투자 최소 · 저비용으로 대용량 즉시배송 급속 확산",
                   h=0.95)
_txt(s, MARGIN_L, CONCL_TOP + 0.95 + 0.30, BODY_W, 0.55, SRC,
     size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
