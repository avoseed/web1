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
