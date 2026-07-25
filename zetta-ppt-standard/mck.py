# -*- coding: utf-8 -*-
"""
mck.py — 맥킨지 템플릿(44종) 한국어 프리셋 래퍼
================================================================
`vendor/mckinsey_pptx` (MIT, AX Labs — 이승필) 엔진을 **한국어 임원 보고용**으로
바로 쓰도록 감싼 얇은 레이어. 원본 템플릿 로직은 수정하지 않는다.

우리 하우스 톤(`frames.py`)과의 관계
  - `frames.py`  : A4 가로 · 정기협의체 하우스 스타일(실측 F1~F12). **국내 임원 보고 정본.**
  - `mck.py`     : 16:9 · 맥킨지 컨설팅 스타일(44종). 컨설팅 제안·전략 리뷰·영문 혼용 덱에.
  둘은 판형이 달라 **한 덱에 섞지 않는다.** 문서 성격에 따라 택1.

사용
    from mck import deck
    b = deck(marker="경쟁사 동향")
    b.add("comparison_table", title="…", options=[...], criteria=[...], recommended_index=1)
    b.save("out.pptx")

템플릿 카탈로그: `vendor/CATALOG.md` · 키 목록은 `KEYS` 참조.
"""
import os
import sys
from dataclasses import replace

_VENDOR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
if _VENDOR not in sys.path:
    sys.path.insert(0, _VENDOR)

from mckinsey_pptx import PresentationBuilder, DEFAULT_THEME  # noqa: E402

KO_FONT = "맑은 고딕"


def ko_theme(family=KO_FONT, copyright_text="ⓒ 롯데마트 · 롯데슈퍼 온라인사업단"):
    """한국어 폰트·푸터로 교체한 테마 (팔레트·레이아웃은 맥킨지 원본 유지)."""
    return replace(
        DEFAULT_THEME,
        typography=replace(DEFAULT_THEME.typography, family=family),
        copyright_text=copyright_text,
    )


def deck(marker=None, theme=None, **kw):
    """한국어 프리셋이 적용된 PresentationBuilder 반환."""
    return PresentationBuilder(theme=theme or ko_theme(),
                               default_section_marker=marker, **kw)


# ── 템플릿 키 (registry) ─────────────────────────────────────
KEYS = {
    "요약·결론": [
        "executive_summary_takeaways",   # 결론 + 근거 불릿 (권장 서두)
        "executive_summary_paragraph",   # 문단형 요약
        "dark_navy_summary",             # 네이비 전면 메시지
        "stat_hero",                     # 대형 숫자 강조
        "quote_slide",                   # 인용
    ],
    "구조": ["cover_slide", "section_divider", "agenda"],
    "비교·평가": [
        "comparison_table",              # 옵션 × 기준 하비볼 (recommended_index 강조)
        "pros_cons",
        "two_column_compare",            # before_after 별칭
        "assessment_table",              # status_overview 별칭 — 신호등 평가
    ],
    "차트": [
        "column_comparison", "column_simple_growth", "column_split_growth",
        "column_historic_forecast", "stacked_column_chart", "grouped_column_chart",
        "line_chart", "bubble_chart", "bubble_chart_takeaways",
    ],
    "매트릭스": ["growth_share", "prioritization_matrix"],
    "조직·분해": ["org_chart", "team_chart", "project_team_circles", "issue_tree"],
    "타임라인·프로세스": [
        "phases_chevron_3", "phases_table_4", "waves_timeline_4",
        "gantt_timeline", "process_activities", "process_flow_horizontal",
        "overview_areas",
    ],
    "다요소": [
        "three_trends_icons", "three_trends_table", "three_trends_numbered",
        "five_key_areas", "funnel", "kpi_dashboard",
    ],
}

# 인자 스키마 — 원본 독스트링 실측(오용 시 플레이스홀더가 남으므로 키를 정확히 지킬 것)
SCHEMA = {
    "executive_summary_takeaways":
        'sections=[{"takeaway":str, "bullets":[str]}], final_conclusion=str',
    "executive_summary_paragraph": 'paragraphs=[str], subtitle=str',
    "dark_navy_summary":  'body=str("[라벨]: 본문" 형태면 라벨 볼드), eyebrow=str',
    "stat_hero":          'stat=str, stat_label=str, context=str, source_text=str',
    "quote_slide":        'quote=str, author=str, author_title=str',
    "cover_slide":        'title, subtitle, client, date, confidentiality',
    "section_divider":    'section_number=str, section_title=str, subtitle=str',
    "agenda":             'items=[str], active_index=int',
    "comparison_table":
        'options=[str], criteria=[{"name":str, "scores":[0~4], "notes":[str]}], recommended_index=int',
    "pros_cons":          'pros=[str], cons=[str], pros_label=str, cons_label=str',
    "two_column_compare": 'left_label=str, right_label=str, left_items=[str], right_items=[str]',
    "assessment_table":
        'categories=[{"name":str, "rows":[{"kpi","target","actual","status_label","status":"green|amber|red"}]}]',
    "prioritization_matrix":
        'items=[{"name", "x_band":0|1|2, "y_band":0|1|2, "status":"green|amber|red"}]',
    "growth_share":       'bus=[{"name","x","y","size","quadrant"?}], x_max, y_max',
    "gantt_timeline":
        'weeks=[…], workstreams=[{"name","start_week":int,"end_week":int,"color":"blue_light|blue_dark"}], milestones=[{"week":int,"label":str}]',
    "phases_chevron_3":
        'phases=[{"label","timeframe","deliverables":[str],"people":[str]}]',
    "phases_table_4":     'phases=[{"name","description","activities":[str],"outcomes":[str]}]',
    "waves_timeline_4":   'waves=[{"name","headline","timeframe","activities":[str],"deliverables":[str]}]',
    "process_activities": 'steps=[{"name","subtitle","activities":[str],"interaction":str,"deliverable":str}]',
    "process_flow_horizontal": 'steps=[{"name","description"}]',
    "overview_areas":     'areas=[{"name","bullets":[str]}], call_out=str',
    "funnel":             'stages=[{"name","value","description"}]',
    "kpi_dashboard":
        'kpis=[{"label","value","delta","delta_dir":"up|down|flat","context"}], columns=int',
    "issue_tree":
        'root=str, main_drivers=[{"label","secondaries":[{"label","underlying":[str]}]}]',
    "org_chart":          'ceo=str, branches=[{"head","reports":[str]}]',
    "team_chart":         'functions=[{"name","description","roles":[{"name","kind":"filled|outline"}]}]',
    "project_team_circles": 'leader={"name","description","icon"}, members=[동일 형태]',
    "three_trends_icons": 'trends=[{"label","bullets":[str],"icon"}]',
    "three_trends_table": 'trends=[{"name","description":[str],"examples":[str]}]',
    "three_trends_numbered": 'trends=[{"label","bullets":[str]}]',
    "five_key_areas":     'areas=[{"name","description"}]',
    "column_comparison":  'categories=[str], values=[float], focus_index=int, takeaways=[str]',
    "column_simple_growth": 'categories, values, growth_pct=str, data_label, data_unit',
    "column_split_growth":  'categories, values, split_index=int, growth_pct_first, growth_pct_second',
    "column_historic_forecast": 'categories, values, forecast_from_index=int, historic_growth, forecast_growth',
    "stacked_column_chart": 'categories=[str], series=[{"name","values":[…]}]',
    "grouped_column_chart": 'categories=[str], series=[{"name","values":[…]}]',
    "line_chart":         'categories=[…], series=[{"name","values":[…],"color"?}]',
    "bubble_chart":       'bubbles=[{"name","x","y","size"}], x_max, y_max, x_label, y_label',
}


def catalog():
    """카탈로그 출력 — 프레임 선택 시 참조."""
    for grp, ks in KEYS.items():
        print("[%s]" % grp)
        for k in ks:
            print("   %-30s %s" % (k, SCHEMA.get(k, "")))
