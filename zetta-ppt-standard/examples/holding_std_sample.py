# -*- coding: utf-8 -*-
"""지주 양식 정본 헬퍼(holding_std) 시험 빌드. 내용은 자리 표시용 예시.
실행: python3 examples/holding_std_sample.py [출력 폴더]"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from holding_std import *   # noqa: E402,F401

out_dir = sys.argv[1] if len(sys.argv) > 1 else "."
prs = new_deck()

# ── 1. 기본형 (좌우 2단)
s = h_chrome(prs, "운영 구조 비교", sub="① 예시")
h_head(s, "L", "단계별 비교")
h_head(s, "R", "비용 상세 분석")

h_caption(s, LT, 3.98, "(A Vs. B)")
y = 4.62
h_table(s, LT, y, [2.4, 3.2, 3.2, 3.2], [0.87, 2.2, 2.2, 2.2, 2.2], [
    ["구 분", "A", "B", "차이"],
    ["소싱", "기준 선별", "원가 우선", [[("품질", True)]]],
    ["물류", "센터 경유", "점포 직납", [[("물류 비용", True)]]],
    ["진열", "소포장", "단순 적재", [[("운영 인력", True)]]],
    ["가격", "본부 통제", "현장 권한", [[("현장 대응", True)]]]])
h_thick(s, LT + 8.8, y, 3.2, 9.67)
h_graybox(s, LT, 14.60, LTW, BOT - 14.60, lines=[
    [("① A는 표준화, B는 현장 재량 중심", True)],
    [("→ ", True), ("단계마다 비용 차이 발생", True, NAVY)]])
h_redline(s, LT + 0.35 + tw("→ ", 11, True), 17.20, tw("단계마다 비용 차이 발생", 11, True) + 0.1)

h_bullet(s, RT, 4.07, [("B의 부대 비용, ", True), ("A 대비 약 ▲40%", True, NAVY)])
h_caption(s, RT, 4.75, "※ 원물 가격 동일 가정", size=9)
y = 5.28
h_table(s, RT, y, [3.0, 3.0, 3.0, 3.0], [0.9, 1.07, 1.07, 1.07, 1.07, 1.07], [
    ["구 분", "A", "B", "A 대비"],
    ["원물", "1,000원", "1,000원", "-"],
    ["부대 비용", [[("134원", True)]], [[("80원", True)]], "약 ▲40%"],
    ["포장", "15원", "6원", "약 ▲60%"],
    ["물류", "31원", "15원", "약 ▲50%"],
    ["최종 원가", "1,134원", "1,080원", "약 ▲5%"]], gray_rows=(5,))
h_redline_center(s, RT + 3.0 + 1.5, y + 2.90, "134원", 10)
h_redline_center(s, RT + 6.0 + 1.5, y + 2.90, "80원", 10)
h_tri(s, RT + RTW / 2, y + 6.45)
h_graybox(s, RT, 12.60, RTW, BOT - 12.60, lines=[
    [("① 포장, 물류 비용 최소화", True)],
    [("→ ", True), ("총원가 약 ▲5% 우위", True, NAVY)],
    [("② 절감한 원가를 판매가에 반영", True)]])
h_footnote(s, "1. 예시 수치")

# ── 2. 카드형 (전폭 요지 + 2×2 카드)
s = h_chrome(prs, "운영 특징")
h_bullet(s, 1.25, 2.85, "대상의 특징을 네 가지 관점으로 정리", w=24.5)
h_bullet(s, 1.25, 3.75, "본부 표준화보다 현장 중심으로 운영", w=24.5)
for (x, y0, label) in ((LT, 5.03, "주요\n업무"), (CARD_RT, 5.03, "진열"), (LT, 11.72, "마케팅"), (CARD_RT, 11.72, "구색")):
    cx, cy, cw_ = h_card(s, x, y0, label)
    h_text(s, cx, cy, cw_, 0.54, "① 핵심 항목", size=11, bold=True)
    h_text(s, cx + 0.53, cy + 0.73, cw_ - 0.6, 0.53, "- 세부 설명", size=9)
    h_text(s, cx, cy + 1.62, cw_, 0.54, "② 핵심 항목", size=11, bold=True)
    h_text(s, cx + 0.53, cy + 2.35, cw_ - 0.6, 0.53, [[("- ", False), ("강조할 세부", True, NAVY)]], size=9)

# ── 3. Summary 장
h_summary(prs, [
    ["첫째 본문 장의 결론 한 문장.", "근거를 잇는 둘째 줄."],
    [[("둘째 본문 장의 결론, ", True), ("핵심 구절은 남색", True, NAVY), (".", True)], "근거를 잇는 둘째 줄."],
    ["셋째 본문 장의 결론 한 문장.", [("부기는 ", True), ("(괄호 10pt 보통)", False, None, 10), (" 으로 표기.", True)]]])

path = os.path.join(out_dir, report_filename("지주 양식 정본 시험", "0.1"))
save_deck(prs, path)
print(path)
