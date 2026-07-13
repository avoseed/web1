# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 경쟁사 동향 1매 — 퀵커머스의 장보기 채널화.

본질 메시지: 속도 경쟁이 아니라 취급 상품 소량 1만 종 → 마트 전 상품 15만 종 확대,
즉시배송이 본 장보기 채널로 전환. 4섹션 상하 적층: ① 배경 → ② 비교표 → ③ 취급 차별점(핵심)
→ ④ 로드맵 → 하단 강조 박스. 순한글만(한자 금지) · 부호 앞 공백 금지 · 오버플로 금지.
출처: 비즈워치·머니투데이·한국경제·뉴스핌·비즈니스포스트('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR

FN = "※ 출처 : 비즈워치 · 머니투데이 · 한국경제 · 뉴스핌 · 비즈니스포스트 ('26.7.8~10)"

prs = new_deck()
# ══ 제목(■) + 리드(강조) ══
s = add_content(prs, None, "이마트, 점포 기반 '2시간 배송' 도입 — 퀵커머스의 장보기 채널화",
                tier="P",
                lead="취급 상품 소량 1만 종 → 마트 전 상품 15만 종 확대, 연말 50여 개 점포 전개")

# ══ ① 도입 배경 ══
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "'바로퀵'(1시간 배송) 매출 197% 증가('26.6월) → 퀵커머스 대용량 상품 수요 확대"),
    (1, "B마트 등 업계 대용량 품목 취급 확대 → 소량 1시간(이륜차)·대용량 2시간(사륜차) 이원화 대응"),
], l=MARGIN_L, t=4.05, w=BODY_W, line_h=0.56)

# ══ ② 서비스 개요 (비교 표) ══
add_bullets(s, [(0, "② 서비스 개요")], l=MARGIN_L, t=y + 0.14, w=BODY_W, line_h=0.5)
TBL_TOP = y + 0.14 + 0.52
tbl = [
    ["구분", "'바로퀵' (기존)", "★ '2시간 배송' (신규)"],
    ["배송", "1H 이내 (반경 3km) · 이륜차", "2H 이내 · 사륜차 · 20시 주문 마감"],
    ["취급", "약 1만 종 (소포장·즉석식품 중심)", "약 15만 종 (이마트 전 상품)"],
    ["출고", "매장 일부 품목", "PP센터 집품·포장 → 즉시 출고"],
    ["조건", "별도 배송비 체계", "무료배송 4만 원 이상 (쓱배송 동일)"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, BODY_W, 3.45, tbl,
              col_w=[3.2, 10.6, BODY_W - 13.8], header_rows=1,
              hl_cols=[2], hl_rows=[2], col_align=["c", "l", "l"])
_txt(s, MARGIN_L, TBL_TOP + 3.52, BODY_W, 0.45,
     "※ 예약배송(주간·새벽·트레이더스) 기존 운영 유지, 병행 이용 가능",
     size=FONT_PT["footnote"], color=SUB_GRAY)

# ══ ③ 취급 상품 차별점 (핵심 — 강조 헤더 바) ══
Y3 = TBL_TOP + 4.15
add_col_header(s, MARGIN_L, Y3, BODY_W, "③ 취급 상품 차별점 — 소량 퀵커머스에서 전 상품 장보기로")
add_bullets(s, [
    (1, "마트 전 상품 대상 : 기존 퀵커머스 미취급 대용량·신선 카테고리 전면 개방"),
    (1, "신선식품(육류·생선·과일) 중심 소구 — 매장 재고 직접 집품으로 선도 유지"),
    (1, "냉장·냉동 전용 보냉 파우치 포장 후 매장 즉시 출고 — 콜드체인 체계 적용"),
    (1, "기존 '바로퀵' 고객층(2030·소용량·그로서리 90%)과 구분되는 본 장보기 수요 겨냥"),
], l=MARGIN_L, t=Y3 + COL_HEADER_H + 0.18, w=BODY_W, line_h=0.62)

# ══ ④ 확산 로드맵 (가로 타임라인, 높이 최소화) ══
Y4 = Y3 + COL_HEADER_H + 0.18 + 4 * 0.62 + 0.14
add_bullets(s, [(0, "④ 확산 로드맵")], l=MARGIN_L, t=Y4, w=BODY_W, line_h=0.5)
add_htimeline(s, MARGIN_L, Y4 + 0.52, BODY_W, [
    ("'26.7월", "양재·하남점\n시범 도입"),
    ("8월", "서울권 확대\n(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 개\n점포 운영"),
], when_h=0.42, content_h=0.95)

# ══ 하단 강조 박스(전폭 네이비) + 각주 ══
BOX_TOP = 16.15
add_conclusion_box(s, MARGIN_L, BOX_TOP, BODY_W,
                   "소량 즉시배송 → 전 상품 장보기 배송 전환, 점포(PP센터) 온라인 가동률 제고",
                   h=0.95)
_txt(s, MARGIN_L, BOX_TOP + 0.95 + 0.22, BODY_W, 0.5, FN,
     size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
