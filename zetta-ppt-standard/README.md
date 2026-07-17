# ZETTA / 온라인사업단 표준 PPT 빌더 (v4)

롯데마트 온라인사업단·ZETTA 보고서용 **표준 PPT 생성 도구**.
내부토의 공통양식(v3)과 Ocado 정기협의체 산출물을 역설계하여 확정한
디자인 토큰·슬라이드 원형을 `python-pptx` 코드로 encode.

## 특징

- **A4 가로 27.52×19.05cm** 고정, **맑은 고딕 3중지정**(한글·영문·숫자)
- 팔레트: 네이비 `#000066` · 크림슨 `#C30C3E` · 표헤더 `#EAEEF6`/`#D8E0EC`
- 전 페이지 공통 크롬: 좌상단 브래킷 헤더 · ■ 리드메시지 · 우상단 단위 · 좌하단 각주
- 원형 6종: 표지(C) · 챕터 간지(D) · P계열 본문 · F계열 본문 · 재무형 표(T) · 종료(Z)

## 설치

```bash
pip install -r requirements.txt
```

## 사용

```python
from zetta_ppt_standard import *

prs = new_deck()
add_cover(prs, "11차 Ocado Project 정기협의체",
          org="온라인사업단", date="2026. 07. 06", tag="내부 토의용")
add_divider(prs, 1, "26년 상반기 경영실적",
            ["경영실적 종합현황", "손익계산서(PL)", "투자현황"])

s = add_content(prs, "경영실적①", "경영실적 종합현황", tier="P",
                lead="핵심 한 줄 메시지", unit="(단위 : 억원, %)",
                footnote="※ ISF+CFC 통합 기준")
data = [["구분","'25년","'26년","증감","증감률"],
        ["총매출","1,677","1,815","+138","+8.2"]]
add_fin_table(s, MARGIN_L, BODY_TOP, BODY_W, 8.0, data,
              col_w=[7.16,4.5,4.5,4.5,4.5], header_rows=1,
              col_align=["c","r","r","r","r"])

add_closing(prs)
prs.save("out.pptx")
```

전체 데모(표지→간지→P표→F→종료):

```bash
python examples/build_demo.py   # demo.pptx 생성
```

## 다른 보고서에 재사용

- **`docs/BUILDER_GUIDE.md`** — 재사용 실무 가이드: 함수 카탈로그 · 레이아웃 레시피 · **정밀 기하 기준**(끝선 정렬·일정 간격·중앙선 침범 금지·하단 채움) · QA 체크리스트.
- **`reports/_template_onepager.py`** — 복사용 스타터(전폭 상하 구조). `cp` 후 «채우기» 부분만 수정하면 새 보고서 완성.

```bash
cp reports/_template_onepager.py reports/build_<주제>.py   # 복사
#   → «채우기» 표시 부분(헤드라인·불릿·표·타임라인·결론·각주)만 수정
python3 reports/build_<주제>.py
```

## 구조

```
zetta_ppt_standard.py     # 빌더 모듈 (토큰·크롬·원형·재무표·타임라인·컬럼 프레임)
docs/SPEC_v4.md           # 표준 명세서 (토큰·원형·프레임워크·정밀 기하 기준 3-4·문구 원칙)
docs/BUILDER_GUIDE.md     # 재사용 가이드 (함수 카탈로그·레시피·기하 기준·QA)
reports/_template_onepager.py  # 복사용 스타터 템플릿 (전폭 상하 구조)
examples/build_demo.py    # 전체 데모
requirements.txt
```

## 시각검수

```bash
# LibreOffice 필요
soffice --headless --convert-to pdf out.pptx
pdftoppm -jpeg -r 130 out.pdf slide
```

## 라이선스 / 주의

사내 표준 서식 기반 — **비공개 리포 권장**. 실적 수치·내부 자료는 커밋 금지.
