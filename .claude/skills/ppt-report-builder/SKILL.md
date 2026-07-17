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
- `references/template_onepager.py` — 복사용 전폭 상하 구조 스켈레톤.

## 워크플로 (매번 이 순서)
1. **톤·프레임워크 선택** — 담백한 표준 보고서면 `zetta_ppt_standard`, 네이비/빨강 임원 덱이면 `tone_v02`. 논리 흐름이 프레임워크를 정한다(단순 나열 금지).
2. **빌드 스크립트 작성** — 스킬의 `scripts/` 를 sys.path 에 넣고 import:
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
3. **빌드**: `python3 build_xxx.py`
4. **렌더 QA(필수)**: LibreOffice headless → PNG 로 **전 슬라이드 육안 검수**.
   ```bash
   soffice --headless --convert-to pdf out.pptx && pdftoppm -png -r 130 out.pdf qa
   ```
   (맑은고딕 미설치 환경은 부호 주위가 벌어져 보이나 **소스 문자열 기준**으로 판정.)
5. **이상 시 해당 요소만 수정 후 재렌더** 반복 → 전 항목 통과 시 완료.

## 반드시 지킬 규칙
- **색·글자**: 파란 글자 전면 금지(표준 톤). 강조 셀(연블루) 안의 네이비 글자만 예외. 크림슨/빨강 = 경고·핵심 수치 강조 전용.
- **결론 = 리드메시지에 선치**(Pyramid). **하단 네이비 결론 박스로 반복 금지**(SPEC 3-1/3-2).
- **정밀 기하(§3-4)**: 한 컬럼 전 객체 동일 폭 끝선 정렬 · 일정 간격 · (좌우 분할 시) 표가 중앙 세로선 침범 금지 · 하단 여백 방치 금지(본문 70%+ 충전, 각주는 콘텐츠 직하).
- **문장부호**: 쉼표·복합어 내 가운뎃점 앞 공백 금지(열거용 ` · ` 는 허용).
- **한자**: 기본 순한글(빌드 후 `[一-鿿]` 검출 0). 사용자가 특정 한자(예: `全`·`化`)를 **명시 요청**할 때만 예외.
- **데이터 정직성**: 미확정 값은 **절대 임의 생성 금지** — `〔조사〕`/`__` 공란 유지. 시장 지표에 Source 각주. 개조식 명사 종결.
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
