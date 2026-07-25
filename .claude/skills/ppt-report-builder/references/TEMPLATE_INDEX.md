# 템플릿 용도 분류 색인 — 57종 (하우스 17 + 맥킨지 40)

> **목적**: "이 내용을 어느 템플릿에 담을 것인가"를 즉시 판정하기 위한 색인.
> 판정 근거는 ① 원본 카탈로그의 Use/Don't-use ② Zelazny 차트 선택법 ③ Minto Pyramid·SCQA
> ④ 우리 표준(`FRAMEWORKS.md` 논리 계열 A~G, `BODY_FRAMES.md` 실측 F1~F12).
> 전 템플릿 **A4 가로 27.52×19.05cm · 1.5cm 액자** 통일 — 판형은 고민 대상이 아니다.

---

## 0. 3단 판정 절차

1. **문서 유형** — 설득·제안형 / 현황·대시보드형 / 동향·비교형 / 실적·재무형 (`FRAMEWORKS.md` §2-1)
2. **논리 관계(핵심 동사)** — *비교한다 · 변한다 · 위치한다 · 나눈다 · 흐른다 · 기인한다 · 얽힌다 · 메운다*
3. **데이터 형태** — 정성 서술 / 항목 비교 / 시계열 / 구성비 / 좌표 / 계층

> **Zelazny 원칙**: 차트는 *데이터*가 아니라 **말하려는 비교(comparison)** 가 정한다.
> 5대 비교 유형 — **구성비(Component) · 항목(Item) · 시계열(Time series) · 분포(Distribution) · 상관(Correlation)**.
> 먼저 "내가 주장할 한 문장"을 쓰고, 그 문장의 비교 유형에 맞는 형태를 고른다.

---

## 1. 논리 관계 → 템플릿 (핵심 색인)

| 논리 관계 | 데이터 형태 | 1순위 | 대안 |
|---|---|---|---|
| **결론 선언** (Pyramid 최상단) | 정성 | `executive_summary_takeaways` | `executive_summary_paragraph`(문어체) · `dark_navy_summary`(단일 메시지) |
| **단일 수치 임팩트** | 수치 1개 | `stat_hero` | `f_bigstat`(하우스) |
| **항목 비교(Item)** | 범주 5~12 | `column_comparison` | `f_compare_table`(정성 속성) |
| **옵션 평가** | 옵션 2~4 × 기준 3~6 | `comparison_table`(하비볼) | `f_assessment` · `pros_cons`(옵션 1개) |
| **상태 전환(전/후)** | 2개 상태 | `two_column_compare` | `f_before_after` |
| **시계열 — 단순 성장** | 5~10 기간 | `column_simple_growth` | `line_chart`(다계열) |
| **시계열 — 변곡** | 성장 국면 2개 | `column_split_growth` | — |
| **시계열 — 실적 vs 전망** | 실적+예측 | `column_historic_forecast` | — |
| **구성비(Component)** | 전체 100% 분해 | `stacked_column_chart` | `funnel`(축소형) |
| **다계열 범주 비교** | 2~4 계열 | `grouped_column_chart` | `line_chart` |
| **위치(2축 연속)** | 5~15 개체 | `bubble_chart` | `bubble_chart_takeaways`(해설 필요) |
| **포트폴리오(BCG)** | 점유율×성장 | `growth_share` | — |
| **우선순위(3×3)** | 임팩트×소요기간 | `prioritization_matrix` | `f_matrix`(2×2 하우스) |
| **분해(MECE)** | 이슈 계층 | `issue_tree` | `f_issue_tree` · `f_pyramid`(전략 체계) |
| **흐름·순서** | 4~6 단계 | `process_flow_horizontal` | `f_process` |
| **인과·기여 분해** | 증감 요인 | **`f_waterfall`**(하우스 — 맥킨지 세트에 워터폴 없음) | `issue_tree` |
| **기여 목표 달성** | 기초→증분→목표 | **`f_waterfall`** | — |
| **병렬 축 3개** | 3개 테마 | `three_trends_icons`/`_numbered`/`_table` | `f_cards` |
| **병렬 축 5~7개** | 5~7 영역 | `five_key_areas`(1줄 설명) · `overview_areas`(불릿) | — |
| **현황 제원·개요** | 항목:값 | **`f_spec_overview`**(하우스 — 현황 보고 정본) | `assessment_table` |
| **KPI 다중 지표** | 4~8 지표 | `kpi_dashboard` | `f_kpi_tiles` · `assessment_table`(목표/실적 필요 시) |
| **목표 대비 상태** | 목표·실적·신호등 | `assessment_table` | `kpi_dashboard` |
| **조직·인력** | 보고체계 | `org_chart` | `team_chart`(기능×역할) · `project_team_circles`(팀 구성) |
| **일정 — 3/4 단계** | 단계형 | `phases_chevron_3` · `phases_table_4` · `waves_timeline_4` | `f_timeline` |
| **일정 — 다과제×주차** | 상세 계획 | `gantt_timeline` | `f_gantt` |
| **일정 — 단기 블록** | 3~4 블록 | `process_activities` | — |
| **정성 근거 2갈래** | 좌우 대비 | `f_two_col`(하우스) | — |
| **고객의 소리** | 인용 1개 | `quote_slide` | — |
| **구조(표지·간지·목차)** | — | `cover_slide` · `section_divider` · `agenda` | — |

---

## 2. 문서 유형별 표준 구성 (덱 골격)

### A. 설득·제안형 (전략 제안·투자 승인) — SCQA + Pyramid
```
cover_slide → agenda
→ executive_summary_takeaways          (A: 결론 선치)
→ column_* / line_chart                (S: 시장·현황 근거)
→ issue_tree 또는 growth_share         (C: 문제·구조 진단)
→ comparison_table                     (Q→A: 옵션 평가·권고)
→ phases_chevron_3 / gantt_timeline    (실행)
→ f_waterfall                          (목표 기여 분해 — 종착)
```

### B. 현황·대시보드형 (운영 보고·QBR)
```
f_spec_overview     (현황 제원 — 정본)
→ kpi_dashboard      (지표 한눈)
→ assessment_table   (목표 대비 신호등)
→ gantt_timeline     (진행 현황)
```

### C. 동향·비교형 (경쟁사 동향·벤치마킹)
```
executive_summary_takeaways
→ f_compare_table 또는 comparison_table   (핵심 비교)
→ column_comparison / growth_share        (정량 위치)
→ f_timeline                              (경과)
```

### D. 실적·재무형
```
kpi_dashboard → column_historic_forecast → stacked_column_chart
→ f_waterfall(증감 요인) → assessment_table
```

---

## 3. 혼동하기 쉬운 짝 (오선택 방지)

| 상황 | ✗ 잘못 | ✓ 올바름 | 이유 |
|---|---|---|---|
| x축이 시간 | `column_comparison` | `column_simple_growth` 계열 | 항목 비교 ≠ 시계열 |
| 축이 연속값 | `prioritization_matrix` | `bubble_chart` | 3×3 밴드는 범주형 전용 |
| 축이 점유율×성장 | `prioritization_matrix` | `growth_share` | BCG는 전용 템플릿 |
| 옵션 1개 | `comparison_table` | `pros_cons` | 비교 대상이 없음 |
| 지표가 목표·실적 필요 | `kpi_dashboard` | `assessment_table` | KPI 타일엔 목표열 없음 |
| 보고체계가 아닌 문제 분해 | `org_chart` | `issue_tree` | 계층의 성격이 다름 |
| 10주 이상 계획 | `phases_*` | `gantt_timeline` | 단계형은 3~4개까지 |
| 부분-전체 | `grouped_column_chart` | `stacked_column_chart` | 구성비는 누적 |
| **현황 보고인데 비교표를 hero** | `comparison_table` | **`f_spec_overview`** | 유형/hero 불일치(§2-1) |

---

## 4. 하우스(`frames.py`) vs 맥킨지(`mck.py`) 선택

**같은 판형(A4 가로)이라 혼용 가능하나, 톤은 문서 단위로 통일.**

| 기준 | 하우스 `frames` | 맥킨지 `mck` |
|---|---|---|
| 톤 | 네이비 `203864` + 크림슨, 담백 | 딥네이비 `0F2A4A` + 브라이트블루 |
| 크롬 | 가운데 제목+밑줄 · **■ 리드에 결론** | 좌측 제목+전폭 밑줄 · 섹션 마커 |
| 강점 | 국내 정기협의체 관례, 표·현황 밀도, **워터폴** | 차트 다양성, 하비볼, 조직도, 간트 |
| 권장 | 국내 임원 보고·정기협의체·현황 | 컨설팅 제안·전략 리뷰·영문 혼용 |

**하우스에만 있는 것**: `f_waterfall`(기여 분해) · `f_spec_overview`(현황 제원) · `f_two_col`.
**맥킨지에만 있는 것**: 하비볼 비교표 · BCG · 조직도 3종 · 간트 · 버블 · 라인/누적/그룹 차트 · 표지/간지/목차 · 인용.

---

## 5. 선택 후 반드시 확인 (`FRAMEWORKS.md` §4)

- [ ] 문서 유형과 hero가 일치하는가(현황형에 비교표 hero 금지)
- [ ] 같은 프레임 3연속이 아닌가
- [ ] 관계 있는 항목을 불릿로 늘어놓지 않았는가(도형 구조화)
- [ ] 1.5cm 액자 안이 하단까지 찼는가
- [ ] 결론이 리드에 선치되었는가

---

### 출처
- 원본 카탈로그: `vendor/CATALOG.md` (mckinsey-pptx, MIT © AX Labs)
- Gene Zelazny, *Say It With Charts* — 5대 비교 유형과 차트 형태 대응
- Barbara Minto, *Pyramid Principle* / SCQA — 결론 선치·MECE 분해
