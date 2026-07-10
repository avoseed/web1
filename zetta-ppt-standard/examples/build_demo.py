# -*- coding: utf-8 -*-
"""표준 빌더 검증용 데모 덱: 표지→간지→P계열(표)→F계열→종료"""
from zetta_ppt_standard import *

prs = new_deck()
TOTAL = 5

# 1. 표지
add_cover(prs, "11차 Ocado Project 정기협의체",
          org="온라인사업단", date="2026. 07. 06", tag="내부 토의용")

# 2. 챕터 간지
add_divider(prs, 1, "26년 상반기 경영실적",
            ["경영실적 종합현황", "손익계산서(PL)", "투자현황"])

# 3. P계열 본문 + 재무형 표
s = add_content(prs, "경영실적①", "경영실적 종합현황", 3, TOTAL, tier="P",
                unit="(단위 : 억원, %)",
                footnote="※ ISF+CFC 통합 기준, 전년 동기 대비")
data = [
    ["구분", "'25년", "'26년", "증감", "증감률"],
    ["총매출", "1,677", "1,815", "+138", "+8.2"],
    ["순매출", "1,195", "1,289", "+93", "+7.8"],
    ["매출이익", "299", "315", "+16", "+5.3"],
    ["판매관리비", "558", "612", "+54", "+9.7"],
    ["영업이익", "▲260", "▲298", "▲38", "▲14.8"],
]
add_fin_table(s, MARGIN_L, BODY_TOP, 25.26, 8.0, data,
              col_w=[7.26, 4.5, 4.5, 4.5, 4.5], header_rows=1,
              col_align=["c", "r", "r", "r", "r"])

# 4. F계열 백업/참고
s = add_content(prs, "Back-up", "주문확대 계획 12월 7,400건 달성", 4, TOTAL, tier="F",
                footnote="※ 일 주문건수 기준, '26.12월 Target")
data2 = [
    ["구분", "핵심 과제", "Target"],
    ["배송서비스 고도화", "새벽배송 신규 오픈 · 주간 Slot 확대", "+1,870건"],
    ["내/외부 협업", "그룹사 협업 · 카카오 커머스 제휴", "+1,930건"],
    ["마케팅 강화", "런칭 마케팅 · CRM 고도화", "+600건"],
]
add_fin_table(s, MARGIN_L, BODY_TOP, 25.26, 4.0, data2,
              col_w=[6.0, 14.26, 5.0], header_rows=1,
              header_fill=TH_SECOND)

# 5. 종료
add_closing(prs)

prs.save("demo.pptx")
print("saved demo.pptx  slides:", len(prs.slides._sldIdLst))
