# -*- coding: utf-8 -*-
"""롯데마트 동남아 유통플랫폼 확장 기획 : 캄보디아 진입 — 임원 보고용 기획안(18매).

ZETTA v4 표준 기반(맑은고딕 3중지정·브래킷 헤더·리딩 밴드·EAEEF6 표 헤더·각주).
제약: 승인/ASK 표현 금지(제언·기대효과·향후 과제로 마감) · 그룹 공식전략/교육용 문구 금지
      · 베트남 손익 수치 금지(자본구조·출점여력·회수기간 논거) · 검증된 수치만 · ※검증 필요 표기 유지.
구성: 표지 + 본문 14 + 부록 3. 본문에 페이지번호 n/N.
출처: Worldometer/UN · IMF · 6Wresearch · US Commerce(trade.gov) · Khmer Times/MoC
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *
from zetta_ppt_standard import _txt, _rect, _line, _set_font

N_BODY = 17                                    # 본문 페이지 수(표지 제외)
_pg = [0]


def pageno(s):
    _pg[0] += 1
    _txt(s, 23.6, 17.95, 3.4, 0.5, f"{_pg[0]} / {N_BODY}",
         size=9, color=SUB_GRAY, align=PP_ALIGN.RIGHT)


def bl(s, items, t=BODY_TOP_LEAD, w=BODY_W, line_h=0.62):
    return add_bullets(s, items, l=MARGIN_L, t=t, w=w, line_h=line_h)


def dbox(s, l, t, w, h, text, fill=NAVY, tcolor=WHITE, size=11, bold=True,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    sp = s.shapes.add_shape(shape, Cm(l), Cm(t), Cm(w), Cm(h))
    sp.shadow.inherit = False
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.color.rgb = fill
    tf = sp.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Cm(0.08))
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = line; _set_font(r, size, bold, tcolor)
    return sp


def darrow(s, x1, y1, x2, y2, color=SUB_GRAY, w_pt=2.25):
    ln = s.shapes.add_connector(2, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    ln.shadow.inherit = False
    ln.line.color.rgb = color; ln.line.width = Pt(w_pt)
    lnEl = ln.line._get_or_add_ln()
    lnEl.append(lnEl.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return ln


SRC_MKT = "※ 출처 : Worldometer/UN · IMF · 6Wresearch · US Commerce(trade.gov) · Khmer Times/MoC"

prs = new_deck()

# ══════════ 표지 ══════════
s = add_cover(prs, "롯데마트 동남아 유통플랫폼 확장 기획 : 캄보디아 진입",
              org="롯데마트", date="2026. 07.", tag="사업 기획안")
_txt(s, 2.0, 9.9, 23.5, 0.9, "베트남 검증 역량 기반 Asset-Light 성장 방안",
     size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# ══════════ 1 · 요약 ══════════
s = add_content(prs, "요약", "Executive Summary", tier="P",
                lead="베트남 검증 역량 기반, 캄보디아 Asset-Light 진입으로 저자본·조기수익 실현")
bl(s, [
    (0, "(배경) 국내 유통 성숙·동남아 소비 고성장 → 신규 프론티어 확보 필요"),
    (0, "(대상) 메콩 3국 중 저위험·고성장·베트남 연계 + 개방적 외자 규제 = 캄보디아"),
    (0, "(방식) 출점 없이 B2B·PB·프랜차이즈·물류 수익화, 한국 상품·신선 특화 차별 포지션"),
    (0, "(관리) 단계별 Gate 검증 후 확장 — Test → Scale → Platform"),
    (0, "(기대효과) 저투자·조기 현금흐름 후 유통 플랫폼 단계 발전"),
], line_h=0.90)
pageno(s)

# ══════════ 2 · 당위성① ══════════
s = add_content(prs, "당위성①", "왜 해외 확장인가", tier="F",
                lead="내수 유통 성숙·동남아 소비 고성장 → 신규 프론티어 확보 필요")
bl(s, [
    (0, "국내 대형마트 시장 성숙·성장 둔화로 외연 확장 필요 (※근거 수치 보강 필요)"),
    (0, "동남아 중산층 확대·소비 고성장의 구조적 기회"),
    (0, "출점형 대규모 투자 부담 → 저자본 진입 방식 요구"),
], line_h=0.95)
pageno(s)

# ══════════ 3 · 당위성② ══════════
s = add_content(prs, "당위성②", "베트남 모델의 한계", tier="F",
                lead="성장은 지속되나 출점형 모델은 구조적 확장 제약 보유")
bl(s, [
    (0, "베트남 진출(2008) 이후 15개점, 남부 집중 → 신규 출점 여력 제한"),
    (0, "출점형 모델은 부지·건축·초기 투자로 점포당 대규모 자본 소요·회수 장기화"),
    (0, "경쟁 심화(AEON·Central 등)로 신규 출점 기대수익 하락"),
    (0, "매출 성장 지속(2024년 +20%)이 입증하듯 제약은 역량 아닌 자산 모델"),
    (1, "시사점 : 프론티어 시장에는 저자본 진입 대안 필요"),
], line_h=0.78)
pageno(s)

# ══════════ 4 · 당위성③ (표) ══════════
s = add_content(prs, "당위성③", "진출국 선정", tier="F",
                lead="메콩 3국 중 캄보디아가 저위험·고성장·시너지 동시 충족",
                footnote="※ 등급은 인구·GDP 성장·규제·물류 연결성 기반 (부록A)")
data = [
    ["국가", "시장성", "리스크", "경쟁강도", "베트남 시너지", "종합"],
    ["캄보디아", "高", "低", "低", "高", "★★★★"],
    ["라오스", "低", "中", "低", "中", "★★"],
    ["미얀마", "高", "高", "中", "低", "★"],
]
add_fin_table(s, MARGIN_L, 4.30, BODY_W, 3.7, data,
              col_w=[4.16, 3.5, 3.5, 3.5, 5.0, 5.5], header_rows=1, hl_rows=[1])
bl(s, [
    (0, "캄보디아 = 달러경제·개방 외자정책"),
    (0, "미얀마 = 정치·규제 리스크 과다"),
    (0, "라오스 = 규모·성장성 제약"),
], t=8.6, line_h=0.85)
pageno(s)

# ══════════ 5 · 적합성① (좌우 2블록) ══════════
s = add_content(prs, "적합성①", "시장 매력도 및 진입 환경", tier="F",
                lead="젊은 고성장 소비시장 + 개방적 진입 규제 — 이중 호조건",
                footnote=SRC_MKT)
HALF = (BODY_W - 0.8) / 2
RCOL = MARGIN_L + HALF + 0.8
add_col_header(s, MARGIN_L, 4.25, HALF, "시장 매력도")
add_bullets(s, [
    (1, "인구 1,740만 · 중위연령 26.4세 · 생산연령 64%"),
    (1, "1인당 GDP 약 2,760달러 (2025)"),
    (1, "소매 성장률 CAGR 약 7~9%"),
    (1, "부분 달러화(외화예금 80%+)로 환위험 완화"),
], l=MARGIN_L, t=5.25, w=HALF, line_h=1.05)
add_col_header(s, RCOL, 4.25, HALF, "진입 규제")
add_bullets(s, [
    (1, "수입·유통 외국인 참여 제한 최소"),
    (1, "상무부 브랜드 독점 수입권 제도"),
    (1, "100% 외자·JV·프랜차이지 지정 가능"),
    (1, "→ Asset-Light 진입의 제도적 기반"),
    (1, "(※프랜차이즈 세부 법제는 향후 과제)"),
], l=RCOL, t=5.25, w=HALF, line_h=1.05)
add_vsep(s, MARGIN_L + HALF + 0.4, 4.25, 10.5)
pageno(s)

# ══════════ 6 · 적합성② (피라미드) ══════════
s = add_content(prs, "적합성②", "유통구조 — 구조적 틈새", tier="F",
                lead="현대유통 태동기 — 플레이어는 존재하나 통합 공급 역량은 부재")
# 피라미드(상→하, 폭 감소): 전통시장 → 현대유통 → E-commerce
dbox(s, 1.7, 5.0, 11.2, 1.7, "전통시장\n(대부분)", fill=SUB_GRAY, size=13, shape=MSO_SHAPE.RECTANGLE)
dbox(s, 3.5, 6.9, 7.6, 1.7, "현대유통\n(도시 집중·초기)", fill=NAVY, size=13, shape=MSO_SHAPE.RECTANGLE)
dbox(s, 5.0, 8.8, 4.6, 1.7, "E-commerce\n(초기)", fill=CRIMSON, size=12, shape=MSO_SHAPE.RECTANGLE)
add_bullets(s, [
    (1, "현대화는 프놈펜·씨엠립 등 도시 집중"),
    (1, "기존 플레이어 존재 : AEON(몰)·Makro(도매)·Lucky·Chip Mong·Thai Huot(슈퍼)"),
    (1, "E-commerce 약 17.8억 달러·로컬 플랫폼 부재"),
    (1, "수입 의존·공급망 파편화 → 소싱-물류-품질 통합 공급 사업자 부재"),
    (0, "결론 : '빈 시장' 아닌 '통합 공급 역량의 틈새'"),
], l=13.6, t=5.0, w=BODY_W - 12.4, line_h=1.15)
pageno(s)

# ══════════ 7 · 적합성③ (표) ══════════
s = add_content(prs, "적합성③", "Right to Win", tier="F",
                lead="기존 플레이어와 정면 경쟁 없는 차별화 포지션 확보 가능",
                footnote="※ 각 사 점포수·포맷 세부는 현지 조사 확정 (부록C)")
data = [
    ["플레이어", "포맷", "우리와의 관계"],
    ["Makro", "범용 캐시앤캐리·도매", "범용 도매 vs 한국 상품·PB·신선 특화 — 카테고리 비중첩"],
    ["AEON", "자산형 몰·GMS", "출점 경쟁 아닌 무출점 공급자 — 잠재 B2B 고객화"],
    ["Lucky·Chip Mong·Thai Huot", "로컬 슈퍼", "경쟁 아닌 핵심 B2B 공급 대상"],
]
add_fin_table(s, MARGIN_L, 4.30, BODY_W, 4.3, data,
              col_w=[5.4, 5.2, BODY_W - 10.6], header_rows=1,
              col_align=["l", "l", "l"])
bl(s, [
    (0, "차별화 3종 : ① K-상품·PB 독점 소싱  ② 신선 콜드체인 표준  ③ 베트남 연계 크로스보더 물류"),
    (1, "핵심 : 기존 소매사업자 = B2B 고객 후보"),
], t=9.1, line_h=0.85)
pageno(s)

# ══════════ 8 · 적합성④ (매핑표) ══════════
s = add_content(prs, "적합성④", "기회 정의", tier="F",
                lead="캄보디아 Pain Point ↔ 롯데 Capability 정합")
data = [
    ["캄보디아 Pain Point", "롯데 Capability"],
    ["공급망 비효율·파편화", "SCM · Hub & Spoke"],
    ["신선 품질 불균일", "품질관리 표준화"],
    ["상품 다양성 부족", "글로벌 소싱 · PB"],
    ["운영 역량 미숙", "매장 운영 매뉴얼"],
]
add_fin_table(s, MARGIN_L, 4.40, BODY_W, 5.0, data,
              col_w=[BODY_W / 2, BODY_W / 2], header_rows=1)
bl(s, [(0, "결론 : Capability-Fit 높은 시장")], t=10.0, line_h=0.8)
pageno(s)

# ══════════ 9 · 적합성⑤ ══════════
s = add_content(prs, "적합성⑤", "우리의 무기 — 크로스보더 공급 네트워크", tier="F",
                lead="베트남 남부 인프라를 캄보디아로 연장하는 국경 연계 공급망",
                footnote="※ 국경 통관 리드타임·관세는 검증 필요 (부록C) — 타당성 조사 1순위")
bl(s, [
    (0, "공급망 : 베트남 남부 15개점 소싱·물류 → 육로 국경(목바이/바벳) → 프놈펜"),
    (0, "소싱 : PB 개발 · 글로벌 직소싱"),
    (0, "운영 : 신선·델리·매장 매뉴얼 이식"),
    (0, "재정의 : SCM → Cross-border Supply Network / PB → Profit Engine"),
], line_h=1.0)
pageno(s)

# ══════════ 10 · 수익모델 (돈 흐름) ══════════
s = add_content(prs, "수익모델", "Asset-Light 구조", tier="F",
                lead="출점 없이 B2B·PB·로열티·물류로 수익 — 저투자·고마진")
dbox(s, 1.3, 6.2, 5.4, 1.9, "베트남 남부\n소싱·물류", fill=NAVY, size=12)
_txt(s, 6.6, 6.3, 2.0, 0.5, "국경 연계", size=9, color=SUB_GRAY, align=PP_ALIGN.CENTER)
darrow(s, 6.75, 7.15, 8.35, 7.15)
dbox(s, 8.4, 6.2, 5.4, 1.9, "B2B 도매\n(초기 현금흐름)", fill=NAVY, size=12)
outs = [("PB · 고마진", CRIMSON, 5.0), ("프랜차이즈 로열티", NAVY, 6.55), ("물류 수익", NAVY, 8.1)]
for label, fill, ty in outs:
    darrow(s, 13.9, 7.15, 15.4, ty + 0.65)
    dbox(s, 15.5, ty, 5.0, 1.3, label, fill=fill, size=11)
bl(s, [
    (0, "Revenue Stream 4종 : 상품공급 마진 · PB · 로열티 · 물류"),
    (1, "특징 : Low Investment · High Margin · Fast Expansion"),
], t=10.6, line_h=0.85)
pageno(s)

# ══════════ 11 · 실행① (표) ══════════
s = add_content(prs, "실행①", "단계별 진입 및 Gate 기준", tier="F",
                lead="단계별 검증 통과 시에만 확장 — Test → Scale → Platform",
                footnote="※ Gate 수치는 파일럿 설계 시 확정 (부록B)")
data = [
    ["단계", "내용", "전환 Gate (가정치)"],
    ["Step1  B2B", "상품 공급 저위험 진입·현금흐름", "월 공급액 · 활성 파트너 N개 · 목표 마진율 달성"],
    ["Step2  Franchise", "마스터 프랜차이즈 무자본 확장", "가맹 점포 수 · 점당 매출 · 로열티 안정화"],
    ["Step3  Platform", "유통 + 물류 생태계 완성", "— (장기 비전)"],
]
add_fin_table(s, MARGIN_L, 4.30, BODY_W, 4.2, data,
              col_w=[4.8, 8.0, BODY_W - 12.8], header_rows=1, col_align=["l", "l", "l"])
bl(s, [(0, "Gate 미달 시 확장 보류·모델 수정 — 손실 하방 차단")], t=9.0, line_h=0.8)
pageno(s)

# ══════════ 12 · 실행② (로드맵 표) ══════════
s = add_content(prs, "실행②", "실행 로드맵", tier="F",
                lead="개시 12개월 전 착수, KPI 기반 관리")
data = [
    ["시점", "활동"],
    ["M-12 (개시 12개월 전)", "프로젝트 착수 · 시장 정밀조사"],
    ["M-10", "물류 인프라 · 국경 통관 설계"],
    ["M-8", "파트너 발굴 · 실사"],
    ["M-3", "상품 박람회 · 소싱 확정"],
    ["M-0", "B2B 사업 개시"],
]
add_fin_table(s, MARGIN_L, 4.30, BODY_W, 5.1, data,
              col_w=[7.5, BODY_W - 7.5], header_rows=1, col_align=["c", "l"], hl_rows=[5])
bl(s, [(0, "KPI : 매출 · 취급 SKU · 파트너 수 · 공급 리드타임")], t=10.0, line_h=0.8)
pageno(s)

# ══════════ 13 · 리스크 (표) ══════════
s = add_content(prs, "리스크", "대응 체계", tier="F",
                lead="관리 가능한 리스크 구조")
data = [
    ["리스크", "대응"],
    ["파트너", "복수 파트너 · 사전 실사"],
    ["물류 · 통관", "국경 통관 사전 검증(M-10) · 현지 3PL 병행"],
    ["경쟁 (Makro 등 도매 확장)", "K-상품 · PB · 신선 특화 차별화"],
    ["환율 · 통화", "달러 결제 기본 · 정부 리엘화 정책 모니터링 (완전 달러화 아님)"],
    ["프랜차이즈 법제 불확실성", "현지 법률 검토 선행"],
]
add_fin_table(s, MARGIN_L, 4.30, BODY_W, 6.0, data,
              col_w=[7.2, BODY_W - 7.2], header_rows=1, col_align=["l", "l"])
pageno(s)

# ══════════ 14 · 제언 ══════════
s = add_content(prs, "제언", "방향 및 향후 과제", tier="P",
                lead="저자본 파일럿 우선 추진, Gate 검증 후 단계 확장")
bl(s, [
    (0, "(제언) B2B 파일럿 저위험 진입, Gate 통과 시 프랜차이즈·플랫폼 확장"),
    (0, "(기대효과) 저투자 조기 현금흐름 · 베트남 역량 재활용 · 통합 공급 틈새 선점"),
    (0, "(향후 검토과제)"),
    (1, "① 경쟁사 점포·포맷 실사   ② 국경 물류 리드타임·관세 타당성"),
    (1, "③ 프랜차이즈 법제   ④ 재무 추정 정교화"),
    (0, "(다음 단계) 하반기 내 현지 정밀조사 착수 → 연내 파일럿 실행계획 수립"),
], line_h=0.82)
pageno(s)

# ══════════ 부록A ══════════
s = add_content(prs, "부록A", "시장 데이터 — 캄보디아", tier="F", footnote=SRC_MKT)
data = [
    ["항목", "값", "출처"],
    ["인구", "약 1,740만", "Worldometer/UN"],
    ["중위연령", "26.4세", "UN"],
    ["생산연령 비중", "64%", "UN"],
    ["1인당 GDP", "약 2,760달러 (2025)", "IMF"],
    ["소매 성장률", "CAGR 약 7~9%", "6Wresearch"],
    ["E-commerce", "약 17.8억 달러 (2025) · 초기", "US Commerce(trade.gov)"],
    ["통화", "부분 달러화 (외화예금 80%+, 리엘 병용)", "Khmer Times/MoC"],
    ["규제", "외국인 유통 제한 최소 · 독점 수입권 · 100% 외자/JV/프랜차이지", "MoC"],
]
add_fin_table(s, MARGIN_L, 3.10, BODY_W, 8.6, data,
              col_w=[4.5, 12.5, BODY_W - 17.0], header_rows=1, col_align=["l", "l", "l"], font_size=10)
pageno(s)

# ══════════ 부록B ══════════
s = add_content(prs, "부록B", "재무 추정 (가정치·예시)", tier="F",
                footnote="※ 본 표는 가정치(예시) — 실제 수치는 재무 추정 정교화 후 확정")
_txt(s, MARGIN_L, 3.05, BODY_W, 0.6, "【 가정치(예시) — 파일럿 설계 시 확정 】",
     size=FONT_PT["bullet0"], bold=True, color=CRIMSON)
data = [
    ["Phase", "투자 (가정치·예시)", "매출 (가정치·예시)", "회수 (가정치·예시)"],
    ["Step1  B2B", "저투자", "상품공급 마진 현금흐름", "조기"],
    ["Step2  Franchise", "무자본 (로열티형)", "로열티 + 공급 마진", "중기"],
    ["Step3  Platform", "단계적 투자", "플랫폼·물류 수익", "장기"],
]
add_fin_table(s, MARGIN_L, 3.85, BODY_W, 4.2, data,
              col_w=[5.0, 6.7, 7.0, BODY_W - 18.7], header_rows=1, col_align=["l", "l", "l", "c"])
_txt(s, MARGIN_L, 8.4, BODY_W, 0.8,
     "Gate 가정치 : 월 공급액 · 활성 파트너 수 · 목표 마진율 (Step 전환 조건, 예시)",
     size=FONT_PT["bullet1"], color=INK)
pageno(s)

# ══════════ 부록C ══════════
s = add_content(prs, "부록C", "검증 필요 항목 (타당성 조사 체크리스트)", tier="F")
data = [
    ["검증 항목", "세부", "비고"],
    ["경쟁사 점포수·포맷", "AEON · Makro · Lucky · Chip Mong · Thai Huot 실사", "현지 조사"],
    ["국경 통관", "목바이/바벳 리드타임 · 관세", "타당성 1순위"],
    ["프랜차이즈 법제", "현지 법률 검토", "선행"],
    ["국내 시장 성숙 근거", "대형마트 성숙 근거 수치 보강", "근거 보강"],
]
add_fin_table(s, MARGIN_L, 3.10, BODY_W, 5.6, data,
              col_w=[6.0, BODY_W - 12.0, 6.0], header_rows=1, col_align=["l", "l", "c"])
pageno(s)

out = os.path.join(os.path.dirname(__file__), "..", "output", "캄보디아_확장기획.pptx")
prs.save(out)
print("saved", os.path.normpath(out), "| slides:", len(prs.slides._sldIdLst))
