# ppt-report-builder (플러그인)

python-pptx 로 임원 보고용 한국어 PPT(장표·원페이저·덱)를 생성하는 **스킬 번들** 플러그인.
Claude Cowork/Code 에서 설치하면 "임원 보고 장표 만들어줘" 류 요청 시 `ppt-report-builder`
스킬이 로드되어 검증된 빌더·디자인 토큰·정밀 기하 규칙·렌더 QA 워크플로를 적용한다.

## 담긴 것
```
skills/ppt-report-builder/
  SKILL.md                     # 진입점(워크플로·규칙·QA)
  scripts/zetta_ppt_standard.py  # ZETTA v4 표준 톤 빌더
  scripts/tone_v02.py            # v0.2 임원 톤(헤더·리딩·빨강강조·표·박스·플로우·차트·피라미드)
  references/SPEC_v4.md · BUILDER_GUIDE.md · template_onepager.py
```
두 디자인 톤(담백한 표준 / 네이비·빨강 임원), 재무형 표·가로/세로 타임라인·2×2 맵·피라미드·
돈흐름 플로우·막대/컬럼 차트를 지원한다.

## 내장 규칙(요약)
- **목업 선(先)승인**: 실제 pptx 빌드·렌더 전에 저비용 텍스트 목업으로 구조·내용 확정 → 승인 후 빌드(토큰 절약)
- 결론(So-What)은 **리드메시지에 선치**, 하단 결론 박스 반복 금지
- **파란 글자 금지**(강조 셀 예외) · 크림슨/빨강은 경고·핵심 수치 전용
- 정밀 기하: 끝선 정렬 · 일정 간격 · 중앙선 침범 금지 · 본문 70%+ 충전
- **데이터 정직성**: 미확정 값은 `〔조사〕`/`__` 공란 — 임의 수치 생성 금지, 시장 지표 Source 각주
- 문장부호 앞 공백 금지 · 기본 순한글(요청 시 한자 예외) · 오버플로 금지

## 설치 (Claude Cowork / Code)
이 리포(`avoseed/web1`)가 마켓플레이스입니다(`.claude-plugin/marketplace.json`).

```
/plugin marketplace add avoseed/web1
/plugin install ppt-report-builder@avoseed-ppt-marketplace
```

- Cowork 앱: 플러그인/마켓플레이스 설정에서 위 마켓플레이스를 추가 후 설치.
- 로컬 개발: `/plugin marketplace add /path/to/web1` 로 로컬 경로 추가 가능.

설치 후에는 별도 호출 없이 PPT 보고서 요청 시 스킬이 자동 트리거되며, `/ppt-report-builder`
로 명시 호출도 가능합니다.

## 소스 정본 / 갱신
편집 정본은 리포의 **`zetta-ppt-standard/`**(빌더 `zetta_ppt_standard.py`·`tone_v02.py`·`docs/`)
와 **`reports/_template_onepager.py`** 입니다. 이 플러그인의 `skills/…/scripts`·`references`
는 그 정본에서 **생성된 번들**(자체 완결·배포용)이며, `SKILL.md` 만 플러그인 고유 문서입니다.

정본을 수정하면 번들을 재생성하세요:
```bash
bash plugins/ppt-report-builder/sync.sh
```
(중복 방지를 위해 별도 `.claude/skills/` 복사본은 두지 않습니다 — 이 워크스페이스에서 쓰려면
플러그인을 설치합니다.)

## 요구 사항
- `python-pptx` (빌드) · LibreOffice + poppler(`pdftoppm`) (렌더 QA) · 맑은 고딕(없으면 렌더 시
  Noto 대체 — 판정은 PPTX 소스 문자열 기준).
