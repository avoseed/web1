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
  references/SPEC_v4.md · BUILDER_GUIDE.md · FRAMEWORKS.md · mockup_template.html · template_onepager.py
```
두 디자인 톤(담백한 표준 / 네이비·빨강 임원), 재무형 표·가로/세로 타임라인·2×2 맵·피라미드·
돈흐름 플로우·막대/컬럼 차트를 지원한다.

## 내장 규칙(요약)
- **목업 선(先)승인**: 실제 pptx 빌드·렌더 전에 저비용 목업으로 구조·내용 확정 → 승인 후 빌드(토큰 절약)
- **프레임 다양성**: 논리 계열(A~G)이 프레임을 정함 — ~25종 카탈로그(비교표·바차트·As-Is/To-Be·SWOT·평가매트릭스·추이·로드맵·순환·2×2·레이더·피라미드·이슈트리·퍼널·프로세스·밸류체인·워터폴·드라이버트리·허브앤스포크·브릿지 등, 없으면 조합·신설). 표+타임라인 스택·상단 4-KPI 스트립 반복 금지 — 단 톤·정렬 규정은 유지
- 결론(So-What)은 **리드메시지에 선치**, 하단 결론 박스 반복 금지
- **파란 글자 금지**(강조 셀 예외) · 크림슨/빨강은 경고·핵심 수치 전용
- 정밀 기하: 끝선 정렬 · 일정 간격 · 중앙선 침범 금지 · 본문 70%+ 충전
- **데이터 정직성**: 미확정 값은 `〔조사〕`/`__` 공란 — 임의 수치 생성 금지, 시장 지표 Source 각주
- 문장부호 앞 공백 금지 · 기본 순한글(요청 시 한자 예외) · 오버플로 금지

## 두 가지 사용 경로 (환경에 맞게)

| 실행 환경 | `/plugin` 사용 | 권장 방식 |
|---|---|---|
| **Claude Code 웹**(claude.ai/code) | ❌ 미지원(`/plugin isn't available`) | **프로젝트 스킬**(아래) — 설치 불필요 |
| **Cowork 데스크톱 앱 / CLI** | ✅ 지원 | 플러그인 설치 **또는** 프로젝트 스킬 |

> ⚠ **웹 세션에는 플러그인/마켓플레이스 시스템이 없다.** 웹에서 `/plugin ...` 은 아무 일도
> 하지 않는다(그래서 버전이 안 오르고 스킬도 안 뜬다). 웹에서는 **프로젝트 스킬**을 쓴다.

### A. 프로젝트 스킬 (이 리포에서 `/plugin` 없이 — 웹 포함 어디서나)
이 리포에는 **`.claude/skills/ppt-report-builder/`** 프로젝트 스킬이 포함되어 있다(정본에서
`sync.sh` 로 생성된 미러). **이 리포에서 작업하면 설치 없이 자동 로드**되며, PPT 보고서 요청 시
트리거되고 `/ppt-report-builder` 로 명시 호출도 된다. (다른 리포에서 쓰려면 이 폴더를 그 리포의
`.claude/skills/` 로 복사하면 된다.)

### B. 플러그인 (Cowork 앱 / CLI — 여러 프로젝트에 공통 배포)
이 리포(`avoseed/web1`)가 마켓플레이스입니다(`.claude-plugin/marketplace.json`).
```
/plugin marketplace add avoseed/web1
/plugin install ppt-report-builder@avoseed-ppt-marketplace
```
설치 후 PPT 보고서 요청 시 자동 트리거, `/ppt-report-builder` 명시 호출 가능.
(로컬 개발: `/plugin marketplace add /path/to/web1` 로 로컬 경로 추가 가능.)

## 업데이트 (새 버전이 안 잡힐 때)
`/plugin marketplace add` 는 **추가 시점의 master 를 로컬에 캐시**한다. 이후 커밋은 마켓플레이스
캐시를 **명시적으로 갱신**해야 반영된다. 최신 버전이 안 보이면 순서대로:

```
/plugin marketplace update avoseed-ppt-marketplace     # ① 캐시를 master 최신으로 재동기화
/plugin install ppt-report-builder@avoseed-ppt-marketplace   # ② 재설치(업데이트)
```
그래도 옛 버전이면 캐시를 완전히 비우고 다시 추가:
```
/plugin marketplace remove avoseed-ppt-marketplace
/plugin marketplace add avoseed/web1
/plugin install ppt-report-builder@avoseed-ppt-marketplace
```
- 갱신 후 **Claude Code 를 재시작(또는 세션 리로드)** 해야 스킬 내용이 다시 로드된다.
- 현재 버전 확인: `/plugin` 메뉴에서 `ppt-report-builder` 의 버전이 `plugin.json` 과 같은지 본다(현재 1.3.0).
- 참고: `/plugin` 계열 명령은 **Cowork 앱 / Claude Code CLI 전용**(웹 세션에선 실행 안 됨).

## 소스 정본 / 갱신
편집 정본은 리포의 **`zetta-ppt-standard/`**(빌더 `zetta_ppt_standard.py`·`tone_v02.py`·`docs/`)
와 **`reports/_template_onepager.py`** 하나뿐입니다. `SKILL.md` 만 번들 고유 문서입니다.

`sync.sh` 는 정본에서 **두 산출물**을 생성합니다(둘 다 손으로 고치지 말 것 — 정본만 고치고 재실행):
1. **플러그인 번들** `plugins/ppt-report-builder/skills/…` — Cowork/CLI 마켓플레이스 설치용.
2. **프로젝트 스킬** `.claude/skills/ppt-report-builder/` — 이 리포 세션(웹 포함)에서 `/plugin` 없이 자동 로드. (번들의 미러)

정본을 수정하면 재생성:
```bash
bash plugins/ppt-report-builder/sync.sh
```
정본은 여전히 `zetta-ppt-standard/` **하나** — 프로젝트 스킬·플러그인 번들은 모두 거기서 파생된
생성물이라 단일 소스 원칙은 유지됩니다.

## 요구 사항
- `python-pptx` (빌드) · LibreOffice + poppler(`pdftoppm`) (렌더 QA) · 맑은 고딕(없으면 렌더 시
  Noto 대체 — 판정은 PPTX 소스 문자열 기준).
