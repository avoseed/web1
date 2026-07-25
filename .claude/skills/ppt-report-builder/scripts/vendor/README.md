# vendor/mckinsey_pptx — 맥킨지 스타일 템플릿 엔진 (도입 원본)

출처: https://github.com/seulee26/mckinsey-pptx (MIT License, © 2026 AX Labs — 이승필)
라이선스 전문: `LICENSE.mckinsey-pptx`

## 도입 방식
**원본 무수정 vendor.** 템플릿 로직·테마를 고치지 않고 그대로 포함한다.
한국어 사용은 상위 래퍼 `../mck.py` 가 담당(폰트·푸터만 교체, 팔레트·레이아웃은 원본 유지).

## 규격 (원본)
- 판형 16:9 (13.333 × 7.5 in) · 폰트 Arial(→ 한국어는 맑은 고딕으로 교체)
- 팔레트 dark_navy `#0F2A4A` · bright_blue `#2E9BD6` · status green/amber/red
- 여백 L/R 0.45in · 제목 0.45in · 제목 밑줄 1.15in · 본문 1.40in · 푸터 7.05in

## 우리 하우스 톤과의 관계
| | 판형 | 용도 |
|---|---|---|
| `frames.py` (하우스) | A4 가로 27.52×19.05cm | 국내 정기협의체·현황 보고 **정본** |
| `mck.py` (이 엔진) | 16:9 | 컨설팅 제안·전략 리뷰·영문 혼용 덱 |

판형이 달라 **한 덱에 섞지 않는다.** 문서 성격으로 택1.

## 갱신
원본 리포에서 `mckinsey_pptx/` 를 다시 복사하면 된다(로컬 수정 없음이 전제).

---

## A4 가로 대응 패치 (avoseed)
원본은 16:9(13.333×7.5in) 캔버스 기준이라 일부 템플릿이 **수평 좌표를 상수로** 갖는다.
우리 표준 판형(A4 가로 27.52×19.05cm = 10.8346×7.5in — **높이는 원본과 동일**)에서도
정합하도록 아래를 패치했다. **세로 좌표·팔레트·타이포는 원본 그대로.**

### 추가 (base.py)
```python
DESIGN_W_IN, DESIGN_M_IN = 13.333, 0.45
hx(theme, x)     # 설계 x  → 현재 판형 x
hw(theme, w)     # 설계 폭 → 현재 판형 폭
hbox(theme, box) # (l,t,w,h) 중 수평 성분만 변환
```
본문 폭 비율 `k = (현재 본문폭) / (13.333 - 2×0.45)` 로 비례 매핑한다.

### 적용 지점
| 파일 | 대상 |
|---|---|
| `column_chart.py` | `DEFAULT_CHART_BOX` · `DEFAULT_TAKEAWAY_BOX` · `TAKEAWAY_DIVIDER_X` (호출 시 `hbox/hx` 적용) |
| `extra_charts.py` | `DEFAULT_CHART_BOX` 파생 좌표 |
| `bubble_chart.py` | 인라인 `plot_box` 4곳 · 우측 패널 · 우선순위 매트릭스 범례(`leg_x`·텍스트 폭) |

그 외 템플릿은 이미 `theme.layout` 에서 폭을 계산하므로 **수정 불필요**.

### 검증
10종 대표 템플릿 A4 빌드 → 전 슬라이드 **가로 넘침 0** (좌 0.59in / 우 ≤10.83in),
LibreOffice 렌더 육안 확인 완료.

### 원본 갱신 시
`mckinsey_pptx/` 재복사 후 위 3개 파일에 동일 패치를 다시 적용할 것.
