---
name: ppt-report-builder
description: Build polished Korean executive PPT reports and one-pagers (임원 보고 장표·원페이저·deck) with python-pptx. Use whenever the user asks to create or edit a .pptx report, 장표, 원페이저, deck, or slide from supplied content — including tables, timelines, pyramids, money-flow/flow diagrams, and bar/column charts. Two built-in design tones (ZETTA v4 standard, v0.2 executive). Produces the .pptx plus a LibreOffice PNG render for visual QA.
---

# PPT Report Builder

python-pptx 로 임원 보고용 PPT(장표·원페이저·덱)를 생성하는 스킬. 검증된 빌더 모듈·디자인 토큰·레이아웃 원형·정밀 기하 규칙을 번들로 제공한다. 새 보고서는 **아래 워크플로**를 따른다.

## 무엇이 들어있나
- `scripts/zetta_ppt_standard.py` — **ZETTA v4 표준 톤** 빌더. A4 가로, 맑은고딕 3중지정, 크롬(브래킷 헤더·리드·각주), 재무형 표(`add_fin_table`), 컬럼 헤더/결론 박스, 타임라인(가로 `add_htimeline`/세로 `add_vtimeline`), 2×2 맵, 불릿 등.
- `scripts/tone_v02.py` — **v0.2 임원 톤** 헬퍼. 네이비 번호박스 헤더(`hdr`)·리딩(`lead`)·하단 중앙 빨강 이탤릭 강조(`redconcl`)·네이비 헤더 표(`tbl`)·카드 박스(`box`)·플로우 노드(`flownode`)·화살표(`arrow`)·**차트**(`column_chart`/`bar_chart`)·피라미드(사다리꼴)·페이지번호.
- `references/SPEC_v4.md` — 표준 명세(토큰·프레임워크·정밀 기하 §3-4·문구 원칙).
- `references/BUILDER_GUIDE.md` — 함수 카탈로그·레시피·QA 체크리스트.
- `references/FRAMEWORKS.md` — **프레임워크 플레이북**. 논리 계열(A~G)별 프레임 카탈로그(~25종: 비교표·바차트·As-Is/To-Be·SWOT·평가매트릭스·추이·타임라인·멀티트랙 로드맵·3-지평선·순환·2×2·퍼셉추얼맵·레이더·피라미드·이슈트리·구성비·레이어드·퍼널·프로세스·밸류체인·머니플로우·워터폴·드라이버트리·허브앤스포크·RACI·브릿지) + 각 언제/언제 아님·빌더 매핑, 반패턴 체크. 닫힌 목록 아님 — 없으면 조합·신설.
- `references/mockup_template.html` — **A단계 목업(승인용) HTML 스캐폴드**. 주요 프레임 원형을 톤·정렬 규정대로 담은 카드(차트·표·매트릭스·피라미드·플로우·워터폴·트리·As-Is/To-Be·브릿지·SWOT·하비볼·간트·퍼널·순환·허브앤스포크). 복제해 실제 장표처럼 채워 보여준다(빌드·렌더 없음).
- `references/template_onepager.py` — 복사용 전폭 상하 구조 스켈레톤.

## 워크플로 (2단계 — 목업 승인 후 빌드)

빌드(python-pptx)와 렌더 QA(LibreOffice→PNG, 이상 시 재렌더 반복)는 **토큰 비용이 크다.**
그래서 **A단계에서 저비용 텍스트 목업으로 구조·내용을 먼저 확정**하고, **승인 뒤에만** B단계
(실제 pptx 빌드·렌더 QA)로 넘어간다. 목업 단계에서는 python-pptx·soffice 를 실행하지 않는다.

### A단계 — 목업 선제시(승인 게이트, 필수)
1. **논리 골격(스토리라인) 설계 — 슬라이드·프레임 이전에**(`references/FRAMEWORKS.md` §1). 덱 단위 논리 구조화:
   ① 덱의 **한 줄 결론**(Governing thought) 확정 → ② **스토리라인 골격 택1**(SCQA / Pyramid(Minto) / 문제-원인-해법-계획 / 현황-목표-격차-실행) → ③ **각 마디 = 1~2 슬라이드**로 분해(마디가 슬라이드의 존재 이유).
   **A단계의 첫 산출물은 이 골격**(마디 목록 + 마디별 논리 라벨) — 슬라이드 목업 이전에 먼저 제시·승인받는다(큰 덱일수록 필수).
2. **톤 선택 + 마디별 프레임 판정** — 톤: 담백한 표준이면 `zetta_ppt_standard`, 네이비/빨강 임원 덱이면 `tone_v02`.
   **프레임은 마디(슬라이드)마다 논리 관계로 정한다**(`FRAMEWORKS.md` §2~§3 선택표·카탈로그): 대비→비교표/바차트, 시간→추이/타임라인 스파인, 위치→2×2 매트릭스, 구성→피라미드/트리, 흐름→프로세스/플로우, 인과→드라이버/워터폴, 병렬→분할 카드, **전환(전/후)→As-Is/To-Be 대형 2박스+화살표**, 격차→브릿지.
   ⚠ **"리드+불릿+표+가로타임라인" 스택 반복 금지**(같은 프레임 3연속 재검토). **관계가 있으면 불릿 대신 도형으로 구조화**(구조 우선, SPEC §3-0). **도형 채움색은 선택** — 무채색 외곽선으로 충분하면 칠하지 않고, 색은 강조에만.
3. **시각 목업 제시(기본) — HTML** : `references/mockup_template.html` 을 복제해 슬라이드마다 `.slide`
   카드 1개로 **실제 장표처럼** 배치·문구·표·차트(형태)·강조를 채워 보여준다. A4 가로 비율·톤 색
   (v0.2 네이비/빨강·표준 담백)을 그대로 써서 승인 전 레이아웃 감을 준다. **python-pptx·soffice 는
   실행하지 않는다** — HTML 정적 미리보기라 토큰이 싸다(차트는 CSS 막대로 형태만).
   - 규칙은 목업에서부터: 결론=리드, 파란 글자 금지, 미확정 수치 `〔조사〕`/`__` 공란(임의 생성 금지), 부호 앞 공백 금지, 순한글.
   - **대규모 덱**은 대표 1~2장만 HTML 상세, 나머지는 아래 1줄 텍스트 요약으로 압축.
   - 초경량이 필요하면 HTML 대신 **텍스트 와이어프레임**만:
     ```
     [S1 · 톤=v0.2 · 유형=표지/본문/간지]  헤더/번호 · 리드(=결론) · 본문(표 열/행 | 타임라인 마디 | 차트 계열·값 | 도식 | 불릿) · 강조(빨강/HL) · 각주/Source
     ```
4. **승인/수정 대기** — 사용자가 OK 하거나 수정본을 확정할 때까지 **B단계로 넘어가지 않는다.** 수정 요청은 목업만 고쳐 다시 제시(빌드 반복 회피).

### B단계 — 빌드·렌더 QA(승인 후에만)
5. **빌드 스크립트 작성** — 스킬의 `scripts/` 를 sys.path 에 넣고 import:
   ```python
   import os, sys
   SK = os.path.join(os.path.dirname(__file__), ".claude", "skills", "ppt-report-builder", "scripts")
   sys.path.insert(0, SK)                 # 경로는 실제 스킬 위치에 맞게
   from zetta_ppt_standard import *       # 표준 톤
   from tone_v02 import *                 # (선택) v0.2 임원 톤
   prs = new_deck()
   # ... 슬라이드 배치 ...
   prs.save("out.pptx")
   ```
6. **빌드**: `python3 build_xxx.py`
7. **렌더 QA(필수)**: LibreOffice headless → PNG 로 **전 슬라이드 육안 검수**.
   ```bash
   soffice --headless --convert-to pdf out.pptx && pdftoppm -png -r 130 out.pdf qa
   ```
   (맑은고딕 미설치 환경은 부호 주위가 벌어져 보이나 **소스 문자열 기준**으로 판정.)
8. **이상 시 해당 요소만 수정 후 재렌더** 반복 → 전 항목 통과 시 완료.

> 목업(A)에서 확정한 구조·문구는 빌드(B)에서 바꾸지 않는다. 빌드 중 배치 문제로 목업과
> 달라지면 목업을 갱신해 재승인받는다 — 렌더 QA 는 잘림·오버플로 등 **구현 품질** 검수용이다.

## 반드시 지킬 규칙
- **목업 선(先)승인**: 실제 pptx 빌드·soffice 렌더 전에 **목업으로 구조·내용을 승인**받는다(A단계). 토큰 절약을 위해 승인 전 빌드 금지.
- **논리 골격 먼저**: 프레임 선택 이전에 덱의 **스토리라인**(SCQA/Pyramid/문제-원인-해법/현황-목표-격차-실행)을 세우고 마디를 슬라이드로 분해한다(`FRAMEWORKS.md` §1). 골격을 목업 첫 산출물로 승인받는다.
- **프레임 다양성**: 논리 관계가 프레임을 정한다(`FRAMEWORKS.md`). 한 스택(표+타임라인) 반복 금지 — 마디별로 매트릭스·피라미드·플로우·차트·카드·브릿지를 논리에 맞게 쓴다. **단, 프레임이 달라도 톤·정밀 기하(끝선 정렬·일정 간격·중앙선·70% 충전)는 그대로 지킨다.**
- **구조 우선·색 최소(SPEC §3-0)**: 슬라이드 1차 목표는 *형태만 봐도* 논리가 드러나는 것. 관계 있는 항목은 불릿 대신 **도형으로 구조화**. **도형 채움색은 선택** — 무채색 외곽선으로 충분하면 칠하지 않고, 색(네이비 면·크림슨)은 강조에만 아껴 쓴다.
- **색·글자**: 파란 글자 전면 금지(표준 톤). 강조 셀(연블루) 안의 네이비 글자만 예외. 크림슨/빨강 = 경고·핵심 수치 강조 전용.
- **결론 = 리드메시지에 선치**(Pyramid). **하단 네이비 결론 박스로 반복 금지**(SPEC 3-1/3-2).
- **정밀 기하(§3-4)**: 한 컬럼 전 객체 동일 폭 끝선 정렬 · 일정 간격 · (좌우 분할 시) 표가 중앙 세로선 침범 금지 · 하단 여백 방치 금지(본문 70%+ 충전, 각주는 콘텐츠 직하).
- **문장부호**: 쉼표·복합어 내 가운뎃점 앞 공백 금지(열거용 ` · ` 는 허용).
- **한자**: 기본 순한글(빌드 후 `[一-鿿]` 검출 0). 사용자가 특정 한자(예: `全`·`化`)를 **명시 요청**할 때만 예외.
- **데이터 정직성 ↔ 밀도**: 미확정 **수치**는 임의 생성 금지(`〔조사〕`/`__`) — 단 **공란이 화면을 지배 금지.** 시장 메커니즘·경쟁 구조·인과·전략·운영 등 **질적 내용은 충실히 채운다**(수치 없이도 서술 가능). 정량이 대부분 미상이면 숫자 자리표를 늘어놓지 말고 **정성 논리로 재구성**. `〔조사〕`는 특정 수치에만 최소로(대략 본문 1/4↑ 공란이면 재설계). 시장 지표엔 Source 각주. 개조식 명사 종결.
- **오버플로 금지**: 표·도형·차트 텍스트 잘림/이탈 없음. 도식은 지도 이미지 대신 자체 도형.

## QA 자동 점검 스니펫
```python
import re
from pptx import Presentation
prs = Presentation("out.pptx"); t=[]; tbl=0
for sl in prs.slides:
    for sh in sl.shapes:
        if sh.has_table:
            tbl+=1
            for i,row in enumerate(sh.table.rows):
                for j in range(len(sh.table.columns)): t.append(sh.table.cell(i,j).text)
        if sh.has_text_frame: t.append(sh.text_frame.text)
a="\n".join(t)
print("표:",tbl,"| 한자:",re.findall(r'[一-鿿]',a),"| 쉼표앞공백:",re.findall(r'\S [,，]',a))
```

## 세부는 references 참조
프레임워크 유형·레시피·함수 시그니처·색/타이포 토큰은 `references/BUILDER_GUIDE.md`, 원칙·정밀 기하·문구 규칙은 `references/SPEC_v4.md` 를 필요할 때 읽는다. 전폭 상하 구조 새 보고서는 `references/template_onepager.py` 를 복사해 «채우기» 부분만 수정.
