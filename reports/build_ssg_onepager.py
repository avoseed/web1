# -*- coding: utf-8 -*-
"""이마트 '2시간 배송' 동향 원페이저 — 확정 빌드 사양(상하 밴드 구조).

가드레일: ① 순한글만(한자 금지) ② 표는 2×2 포지셔닝 맵 하나뿐(라인업 표 금지)
③ 맵은 3×3 표 객체(파편화 금지) ④ 오버플로 금지 ⑤ 상하 밴드 적층(② 밴드만 좌우 분할,
전체 좌우 2단 금지) ⑥ 로드맵은 가로 타임라인 ⑦ 빈 셀 '해당 없음'(회색) ⑧ 부호 앞 공백 금지.
논리: ① 배경(왜) → ② 개요+맵(무엇) → ③ 확산(어떻게).
출처: 비즈워치·머니투데이·한국경제('26.7.8~10)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *

SRC = ("※ 취급 품목: 이마트 15만 · 바로퀵 1만 종      "
       "※ 출처: 비즈워치 · 머니투데이 · 한국경제 ('26.7.8~10)")

# ── ② 밴드 좌우 분할 (좌 55% 스펙 / 우 45% 맵) — 이 밴드 안에서만 분할 ──
B2_GAP = 0.50
LW2 = (BODY_W - B2_GAP) * 0.55
RW2 = (BODY_W - B2_GAP) * 0.45
COL2R_L = MARGIN_L + LW2 + B2_GAP
VSEP2_X = MARGIN_L + LW2 + B2_GAP / 2

prs = new_deck()
# ══ 헤드라인(전폭) = 결론 선치 ══
s = add_content(prs, "동향", "이마트, 점포 물류거점화로 대용량 즉시배송 선점", tier="P",
                lead="'2시간 배송' 양재·하남 시범(7.9~) → 연말 50여 점 · 유휴 PP센터 재활용 저비용 모델",
                footnote=SRC)

# ══ ① 도입 배경 (전폭 밴드) — 3단 가로 흐름, 마지막에 공백 명시 ══
add_bullets(s, [(0, "① 도입 배경")], l=MARGIN_L, t=4.15, w=BODY_W, line_h=0.5)
add_flow_row(s, MARGIN_L, 4.72, BODY_W, [
    "바로퀵 매출 197%↑ ('26.6월) (미검증)",
    "(배경) B마트 등 즉시배송 시장 대용량 품목 확대",
    "기존 라인업, '빠름 × 대용량' 부재",
], h=1.45)

# ══ ② 서비스 개요 (유일한 좌우 분할 밴드) ══
add_bullets(s, [(0, "② 서비스 개요")], l=MARGIN_L, t=6.70, w=BODY_W, line_h=0.5)
B2_TOP = 7.30
B2_BOT = 12.65
# 좌 55% — 스펙 4종
add_bullets(s, [
    (1, "거점: 유휴 점포 PP센터 (반경 내 즉시 출고)"),
    (1, "운송: 쓱배송 사륜차 → 대용량·중량 대응"),
    (1, "운영: 20시 마감 → 2시간 내 배송 완료 (기존 예약배송 병행)"),
    (1, "규모: 약 15만 종 (바로퀵 1만 종 대비 15배)"),
], l=MARGIN_L, t=B2_TOP + 0.10, w=LW2, line_h=0.95)
# 우 45% — 2×2 포지셔닝 맵(앵커, 유일한 표). 빈 사분면 '해당 없음'(§0.7)
add_matrix2x2(s, COL2R_L, B2_TOP, RW2, B2_BOT - B2_TOP,
              col_labels=["소량", "대용량"],
              row_labels=["빠름\n(당일 1~2H)", "예약\n(익일~)"],
              cells=[["바로퀵 (1H)", "★ 2시간 배송 (신규)"],
                     [None, "주간·새벽·트레이더스"]],
              star=(0, 1), corner="속도＼물량", dim=(1, 0))
# ② 밴드 내부 좌우 구분선(이 밴드 한정)
add_vsep(s, VSEP2_X, B2_TOP, B2_BOT - B2_TOP)

# ══ ③ 확산 계획 (전폭 밴드) — 가로 타임라인 4노드 + 리드 ══
add_bullets(s, [(0, "③ 확산 계획")], l=MARGIN_L, t=13.05, w=BODY_W, line_h=0.5)
y = add_htimeline(s, MARGIN_L, 13.65, BODY_W, [
    ("'26.7월", "양재·하남점\n도입"),
    ("8월", "서울 확대\n(월계·가든5·신도림)"),
    ("9월", "전국 확대"),
    ("연말", "50여 점포\n운영"),
])
add_bullets(s, [
    (0, "(리드) 유휴 PP센터 재활용 → 신규 투자 최소 · 저비용 급속 확산"),
], l=MARGIN_L, t=y + 0.20, w=BODY_W, line_h=0.6)

out = os.path.join(os.path.dirname(__file__), "ssg-quickcommerce-onepager.pptx")
prs.save(out)
print("saved", out)
