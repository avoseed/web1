#!/usr/bin/env bash
# 플러그인 스킬 번들을 편집 정본(zetta-ppt-standard/ + reports/_template_onepager.py)에서 재생성.
# 정본을 고치면 이 스크립트를 돌려 플러그인 번들을 갱신한다. (SKILL.md 는 플러그인 고유 — 동기화 대상 아님)
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
root="$(cd "$here/../.." && pwd)"
src="$root/zetta-ppt-standard"
dst="$here/skills/ppt-report-builder"

cp "$src/zetta_ppt_standard.py"       "$dst/scripts/zetta_ppt_standard.py"
cp "$src/tone_v02.py"                 "$dst/scripts/tone_v02.py"
cp "$src/docs/SPEC_v4.md"             "$dst/references/SPEC_v4.md"
cp "$src/docs/BUILDER_GUIDE.md"       "$dst/references/BUILDER_GUIDE.md"
cp "$src/docs/mockup_template.html"   "$dst/references/mockup_template.html"
cp "$root/reports/_template_onepager.py" "$dst/references/template_onepager.py"

echo "synced plugin bundle from 정본 (zetta-ppt-standard/ + template)"
