# ZETTA PPT 빌더 — 재사용 가이드

> 이 표준으로 **어떤 보고서든** 만들기 위한 실무 지침. 원칙·근거는 `SPEC_v4.md`,
> 이 문서는 **함수 카탈로그 + 레이아웃 레시피 + 정밀 기하 기준 + 복사용 스켈레톤**.
> 새 보고서는 `reports/_template_onepager.py` 를 복사해 내용만 채우면 된다.

---

## 0. 목업 선제시(빌드 전 승인 게이트)

빌드(python-pptx)·렌더 QA(soffice→PNG 반복)는 비용이 크다. **실제 pptx 를 만들기 전에 저비용
텍스트 목업으로 구조·문구를 확정하고 승인받는다.** 슬라이드별로 `[제목·유형] / 리드(=결론) /
본문 요소(표 열·행, 타임라인 마디, 차트 계열·값, 도식 노드) / 강조 / 각주` 만 적어 제시하고,
승인 후에만 아래 빌드로 넘어간다. 미확정 수치는 `〔조사〕`/`__` 로 남긴다(임의 생성 금지).
대규모 덱은 대표 1~2장 상세 + 나머지 1줄 요약으로 압축한다. (플러그인 `SKILL.md` A단계와 동일.)

## 0. 30초 시작

```python
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "zetta-ppt-standard"))
from zetta_ppt_standard import *

prs = new_deck()                                  # A4 가로 빈 덱
s = add_content(prs, "동향", "제목", tier="P",     # 헤더 + 리드(선택) 크롬만 세팅
                lead="핵심 한 줄 결론")
add_bullets(s, [(0, "① 소제목"), (1, "하위 근거")], l=MARGIN_L, t=4.15, w=BODY_W)
prs.save("out.pptx")
```

렌더 검수: `soffice --headless --convert-to pdf out.pptx && pdftoppm -png -r 150 out.pdf qa`
(맑은 고딕 미설치 환경은 부호 주위가 벌어져 보이나 **소스 문자열 기준**으로 판정 — §5 QA).

---

## 1. 빌더 함수 카탈로그

좌표는 모두 **cm**, 원점 좌상단. 대부분 종료 y(다음 객체 시작점)를 반환한다.

### 덱·크롬 (원형)
| 함수 | 용도 |
|---|---|
| `new_deck()` | A4 가로 빈 프레젠테이션 |
| `add_cover(prs, title, org, date, tag)` | 표지(C) — 태그·대제목(밑줄)·조직/일자 |
| `add_divider(prs, num, title, items)` | 챕터 간지(D) — 회색 풀블리드·`N. 제목`·①②③ |
| `add_content(prs, category, title, tier, lead, unit, footnote)` | 본문 슬라이드(P/F) — 헤더+리드+단위+각주 크롬만, 본문은 자유 배치 |
| `add_closing(prs, company)` | 종료(Z) 마감 장표 |
| `add_header / add_lead / add_unit / add_footnote` | 크롬 요소 개별 배치(보통 add_content 가 호출) |

### 콘텐츠 블록
| 함수 | 용도 |
|---|---|
| `add_bullets(slide, items, l, t, w, line_h)` | 개조식 불릿. `items=[(level, text)]` — 0 `•`볼드 / 1 `-` / 2 `·`. **종료 y 반환** |
| `add_block_header(slide, l, t, w, text)` | 【 소제목 】 12pt 볼드 |
| `add_hline(slide, l, t, w)` | 가로 괘선(구분) |

### 표 (재무형·2×2 맵)
| 함수 | 용도 |
|---|---|
| `add_fin_table(slide, l, t, w, h, data, col_w, header_rows, merges, hl_rows, hl_cols, hl_cells, dim_cells, header_cols, bold_cols, col_align, font_size)` | 담백한 그리드 표. `merges=[(r0,r1,col)]` 세로 병합, `hl_rows` 강조 행(연블루+네이비), `dim_cells` 옅은 회색 셀. 병합 연속 셀은 `None`. |
| `add_matrix2x2(slide, l, t, w, h, col_labels, row_labels, cells, star, corner, dim)` | 2×2 포지셔닝 맵을 **단일 3×3 표 객체**로. `star` 강조 셀, `dim=(행,열)` 빈 사분면 → `해당 없음`(회색) |

### 타임라인
| 함수 | 용도 |
|---|---|
| `add_timeline(slide, l, t, w, milestones)` | 가로 타임라인(경과형 E) — 굵은 축+화살촉, 상단 날짜·하단 라벨 |
| `add_htimeline(slide, l, t, w, steps, when_h, content_h)` | 전폭 밴드용 가로 4노드 — 노드 위 시점(볼드)·아래 내용(2줄 허용) |
| `add_vtimeline(slide, l, t, w, h, steps, max_gap)` | 세로 타임라인 — 좁은 세로 컬럼(좌우 분할 밴드 등) 채움. `max_gap` 노드 간격 상한 |

### 2단 컬럼 프레임 (본장형)
| 함수 | 용도 |
|---|---|
| `add_col_header(slide, l, t, w, text)` | 컬럼 헤더 바(#EAEEF6, 검정 볼드). **종료 y = t + COL_HEADER_H** |
| `add_vsep(slide, x, t, h)` | 중앙 세로 구분선(#BFBFBF 0.75pt 실선) |
| `add_conclusion_box(slide, l, t, w, text, h)` | 네이비 박스(백색 볼드) — **요약 스트립 전용**(예: 캄보디아 관리원칙). ⚠ So-What 은 리드에 두고 이 박스로 반복 금지 |
| `add_specbox(slide, l, t, w, h, rows, header)` | 스펙 패널(프레임+헤더+라벨:값 균등) — 표 아닌 패널. 맵과 쌍둥이용 |
| `add_frame(slide, l, t, w, h)` | 얇은 테두리 박스(타임라인 등 감싸기) |
| `add_insight_box(slide, l, t, w, h, items, tab)` | 회색 테두리 시사점 박스 + 짙은 탭 |

> ⚠ **본장 금지**: 마케팅 칩·화살표 흐름(`add_flow_row`)·키메시지/단계 박스는 정기협의체 본장 원형이 아니다(SPEC 3-1). 담백한 표·컬럼 헤더·결론 박스로 구성.

---

## 2. 레이아웃 레시피

> ⚠ **6:4 좌우 2단 폐기(v4.1)**: 좌우 폭 배분·바닥선 정렬 부담과 한쪽 성김 탓에 제거. 경쟁사 동향·현황형의 **정본은 전폭 상하 구조(A)**.

### A. 전폭 상하 구조 (근거 표 중심) — 기본값·정본
전폭 밴드 적층(좌우 2단 아님): **결론(So-What)은 리드메시지에 선치** → ① 도입 배경(불릿) → ② **전폭 근거 표**(`add_fin_table`, 그룹 병합 + 신규 열/행 강조 — 장표 중심) → ③ **가로 타임라인**(`add_htimeline`, 본문 최종 요소). **하단 네이비 결론 박스는 비권장**(결론은 리드에 있음 — 반복 금지). 전 객체 동일 폭 `BODY_W` 끝선 정렬, 요소 높이·간격으로 하단까지 채움. → `reports/_template_onepager.py`

### B. 재무 전폭 표 / 근거 표 + 해설
상단 전폭 `add_fin_table`(손익·KPI·비교) → 하단 2~3단 분해(【 】 + 불릿).

### C. 상하 밴드 + 1개 밴드만 좌우 분할
전폭 밴드 적층에서 한 밴드만 좌우 분할이 필요할 때(그 밴드에 한해 §3-3 중앙선·바닥선 기하 적용). 로드맵은 `add_htimeline`.

프레임워크 선택은 **논리 흐름**이 정한다(배경→근거→실행=전폭 상하). 단순 상하 나열은 금지 — 각 밴드가 논증 단계여야 한다.

---

## 3. 정밀 기하 기준 (정기협의체 필수 — 이번 세션 확정)

정기협의체 자료의 핵심은 **엄격한 끝선 정렬 + 일정한 간격**이다. 위반 시 '엉성'해 보인다.

1. **끝선 정렬**: 전 객체(불릿·표·타임라인·결론)는 **동일 폭**으로 좌·우 끝선을 맞춘다.
   전폭 상하 구조는 모두 `BODY_W`. 서로 다른 폭 금지.
2. **일정 간격**: 블록 간 간격을 토큰 하나(`GB`)로 통일. 들쭉날쭉한 여백 금지.
3. **중앙 세로선 침범 금지 (좌우 분할 밴드에만)**: 좌우 분할이 있을 때 표 끝선과 세로선 사이
   **여백(≥0.5cm)** 을 좌·우 대칭으로 확보(`GAP` 0.66→1.10). 표가 세로선에 **닿으면 안 됨**.
   전폭 상하 구조는 세로선이 없어 해당 없음.
4. **표 밀도**: 행을 늘어뜨리지 말 것 — 조밀하게. 근거 표는 **전폭으로 크게** 살린다.
5. **하단 처리 — 2가지 모드 중 택1** (하단 여백은 재발 금지 1순위):
   - **채움 모드**: 밀도를 다소 넓히더라도 객체를 하단선(`CONTENT_BOTTOM`=17.55, 1.5cm 여백)까지 분산해 페이지를 채우고
     각주를 그 직하에 둔다. (콘텐츠가 페이지를 채워야 할 때)
   - **컴팩트 모드**: 표·객체를 조밀하게 두고 **각주를 콘텐츠 직하로 끌어올려** 콘텐츠-각주
     간극을 없앤다. 각주를 페이지 바닥(17.90)에 **방치해 콘텐츠와 벌어지게 두지 말 것.**
6. **좌우 바닥선 정렬 (좌우 분할 밴드에만)**: 좌우 분할 밴드는 두 컬럼 최하단 객체 하단선을 **같은 y**로. 전폭 상하 구조는 밴드가 순차 적층되므로 해당 없음.

측정 검증(빌드 후): (좌우 분할 시) 표 끝선–`VSEP_X` 대칭 거리, 각 밴드 폭 일치, 각주 y(콘텐츠 직하)를 실측해 확인.

---

## 4. 색·타이포 빠른 참조 (임의 값 금지)

| 상수 | 값 | 용도 |
|---|---|---|
| `NAVY` #000066 | 면 채움 전용(결론 박스) — **글자색 금지** |
| `CRIMSON` #C30C3E | 경고·리스크 강조 |
| `INK`/`BLACK` #000/#222 | 본문/제목 글자 |
| `TH_PRIMARY` #EAEEF6 | 표·컬럼 헤더 채움 |
| `HL_FILL`/`HL_TEXT` #DEEBF7/#002060 | 강조 행 채움/글자 — **파란 글자 유일 예외** |
| `GRID` #BFBFBF | 표 격자·세로선 |
| `SUB_GRAY` #595959 | 보조·의도된 공백(`해당 없음`) |

타이포(`FONT_PT`): 제목 20 / 리드 14 / •12 / -11 / ·10 / 표 10 / 각주 9. **파란 글자 전면 금지.**
간격 토큰: `BODY_TOP_LEAD=4.10` `GAP_SECTION=0.50` `GAP_HEAD=0.55` `LINE_H=0.48`
레이아웃(v4.2 여백·충전 액자): `MARGIN_L=MARGIN_R=MARGIN_B=1.50` · `BODY_W=24.52` · `CONTENT_BOTTOM=17.55`(콘텐츠는 여기까지 충전) · `COL_GAP=0.66`(2단은 1.10 권장).

---

## 5. QA 체크리스트 (빌드 후 필수)

- [ ] **한자 0자**: `[一-鿿]` 검출 = 0 (순한글).
- [ ] **표 개수**: 의도한 수와 일치(2×2 맵형은 1개). 스펙 패널·프레임은 표 객체 아님.
- [ ] **부호 앞 공백 0**: 쉼표 앞 공백 `\S [,，]` = 0(자동). 가운뎃점은 **열거 구분용(양쪽 공백 허용)** 과 복합어 내부(공백 금지)를 구분해야 하므로 육안 판정. 렌더의 벌어짐은 폰트 대체 아티팩트 — 소스 기준.
- [ ] **오버플로 없음**: 셀·박스 텍스트 이탈·줄바꿈 깨짐 없음.
- [ ] **끝선/간격/중앙선**: 표 끝선-세로선 대칭 여백, 좌우 바닥선 정렬(§3).
- [ ] **하단 처리**: 채움 또는 컴팩트+각주직하 — 콘텐츠-각주 벌어짐 없음.
- [ ] **파란 글자 없음**: 강조 셀(HL_TEXT) 외 파란 글자 0.
- [ ] 렌더(PDF→PNG) 육안 검수.
```
# 한자·부호·표수 자동 점검 스니펫
python3 - <<'PY'
import re
from pptx import Presentation
prs = Presentation("out.pptx"); t=[]; tbl=0
for sl in prs.slides:
    for sh in sl.shapes:
        if sh.has_table:
            tbl += 1
            for i, row in enumerate(sh.table.rows):
                for j in range(len(sh.table.columns)):
                    t.append(sh.table.cell(i, j).text)
        if sh.has_text_frame:
            t.append(sh.text_frame.text)
a = "\n".join(t)
# 쉼표 앞 공백만 자동 검출(가운뎃점은 열거용 허용 → 육안). 렌더 벌어짐은 폰트 아티팩트.
print("표:", tbl, "| 한자:", re.findall(r'[一-鿿]', a), "| 쉼표앞공백:", re.findall(r'\S [,，]', a))
PY
```
