# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 경쟁사 동향 1매 — 퀵커머스의 장보기 채널化.

본질: 취급 상품 소량 1만 종 → 마트 全 상품 15만 종 확대. 4섹션 상하 적층:
① 배경 → ② 비교표 → ③ 취급 차별점(핵심) → ④ 가로 타임라인 → 하단 강조 박스.
표기: 이 장표는 사용자 지시로 全·化 한자 표기 사용(순한글 가드레일 예외).
강조 역할 분리: 신규 열=음영 / 취급 행=볼드만. 리드 '1만→15만' = 표 취급 행 일치.
출처: 비즈워치·머니투데이·한국경제·뉴스핌·비즈니스포스트('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, MSO_ANCHOR

FN = "※ 출처 : 비즈워치 · 머니투데이 · 한국경제 · 뉴스핌 · 비즈니스포스트 ('26.7.8~10)"

prs = new_deck()
# ══ 제목(■ 은 제목 전용) ══
s = add_content(prs, None, "이마트, 점포 기반 '2시간 배송' 도입 — 퀵커머스의 장보기 채널化",
                tier="P")
# 리드(기호 없는 볼드 1줄·검정 — 파란 글자 금지) — v3 문구 반영
_txt(s, LEAD[0], LEAD[1], 24.6, LEAD[3],
     "취급 상품 소량 1만 종 → 마트 全 상품 15만 종 확대 — 점포(PP센터) 온라인 가동률 제고 목적",
     size=FONT_PT["lead"], bold=True, color=INK)

# ══ ① 도입 배경 ══
y = add_bullets(s, [
    (0, "① 도입 배경"),
    (1, "'바로퀵'(1시간 배송) 매출 197% 증가('26.6월) → 퀵커머스 대용량 상품 수요 확대"),
    (1, "B마트 등 업계 대용량 품목 취급 확대 → 소량 1시간(이륜차)·대용량 2시간(사륜차) 이원화 대응"),
], l=MARGIN_L, t=4.05, w=BODY_W, line_h=0.56)

# ══ ② 서비스 개요 (비교 표) — 신규 열 음영 / 취급 행 볼드만 ══
y = add_bullets(s, [(0, "② 서비스 개요")], l=MARGIN_L, t=y + 0.12, w=BODY_W, line_h=0.50)
TBL_TOP = y + 0.02
H_TBL = 3.40
tbl = [
    ["구분", "'바로퀵' (기존)", "★ '2시간 배송' (신규)"],
    ["배송", "1H 이내 (반경 3km) · 이륜차", "2H 이내 · 사륜차 · 20시 주문 마감"],
    ["취급", "약 1만 종 (소포장·즉석식품 중심)", "약 15만 종 (이마트 全 상품)"],
    ["출고", "매장 일부 품목", "PP센터 집품·포장 → 즉시 출고"],
    ["조건", "별도 배송비 체계", "무료배송 4만 원 이상 (쓱배송 동일)"],
]
add_fin_table(s, MARGIN_L, TBL_TOP, BODY_W, H_TBL, tbl,
              col_w=[3.2, 10.6, BODY_W - 13.8], header_rows=1,
              hl_cols=[2], bold_rows=[2], col_align=["c", "l", "l"])
tfn_y = TBL_TOP + H_TBL + 0.04
_txt(s, MARGIN_L, tfn_y, BODY_W, 0.40,
     "※ 예약배송(주간·새벽·트레이더스) 기존 운영 유지, 병행 이용 가능",
     size=FONT_PT["footnote"], color=SUB_GRAY)

# ══ ③ 취급 상품 차별점 (핵심) — 좌측 개조식 헤더(검정·음영 밴드 없음) ══
y = add_bullets(s, [
    (0, "③ 취급 상품 차별점 — 소량 퀵커머스에서 全 상품 장보기로"),
], l=MARGIN_L, t=tfn_y + 0.55, w=BODY_W, line_h=0.54)
y = add_bullets(s, [
    (1, "마트 全 상품 대상 : 기존 퀵커머스 미취급 대용량·신선 카테고리 전면 개방"),
    (1, "신선식품(육류·생선·과일) 중심 소구 — 매장 재고 직접 집품으로 선도 유지"),
    (1, "냉장·냉동 전용 보냉 파우치 포장 후 매장 즉시 출고 — 콜드체인 체계 적용"),
    (1, "기존 '바로퀵' 고객층(2030·소용량·그로서리 90%)과 구분되는 본 장보기 수요 겨냥"),
], l=MARGIN_L, t=y + 0.10, w=BODY_W, line_h=0.70)

# ══ ④ 확산 로드맵 (가로 타임라인, 1줄 라벨) — v3 구조: 하단 박스 없음 ══
y = add_bullets(s, [(0, "④ 확산 로드맵")], l=MARGIN_L, t=y + 0.24, w=BODY_W, line_h=0.54)
tl_bot = add_htimeline(s, MARGIN_L, y + 0.16, BODY_W, [
    ("'26.7월", "양재·하남점 시범"),
    ("8월", "서울권 확대(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포 운영"),
], when_h=0.45, content_h=0.55, size=FONT_PT["bullet2"])

# ══ 각주(콘텐츠 직하) ══
_txt(s, MARGIN_L, tl_bot + 0.30, BODY_W, 0.5, FN,
     size=FONT_PT["footnote"], color=INK)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
