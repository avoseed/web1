#!/usr/bin/env bash
# 스킬 번들을 편집 정본(zetta-ppt-standard/ + reports/_template_onepager.py)에서 재생성.
# 정본을 고치면 이 스크립트를 돌려 두 산출물을 갱신한다:
#   (1) 플러그인 번들   plugins/ppt-report-builder/skills/...   ← Cowork/CLI 마켓플레이스 설치용
#   (2) 프로젝트 스킬   .claude/skills/ppt-report-builder/       ← 이 리포 세션(웹 포함)에서 /plugin 없이 자동 로드
# 정본은 여전히 zetta-ppt-standard/ 하나. SKILL.md 만 번들 고유(→ 프로젝트 스킬로 미러).
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
root="$(cd "$here/../.." && pwd)"
src="$root/zetta-ppt-standard"
dst="$here/skills/ppt-report-builder"

cp "$src/zetta_ppt_standard.py"       "$dst/scripts/zetta_ppt_standard.py"
cp "$src/tone_v02.py"                 "$dst/scripts/tone_v02.py"
cp "$src/docs/SPEC_v4.md"             "$dst/references/SPEC_v4.md"
cp "$src/docs/BUILDER_GUIDE.md"       "$dst/references/BUILDER_GUIDE.md"
cp "$src/docs/FRAMEWORKS.md"          "$dst/references/FRAMEWORKS.md"
cp "$src/docs/mockup_template.html"   "$dst/references/mockup_template.html"
cp "$root/reports/_template_onepager.py" "$dst/references/template_onepager.py"

echo "synced plugin bundle from 정본 (zetta-ppt-standard/ + template)"

# (2) 프로젝트 스킬 = 플러그인 번들의 미러(SKILL.md + scripts + references). 손으로 고치지 말 것 — 정본만 고치고 이 스크립트 실행.
proj="$root/.claude/skills/ppt-report-builder"
rm -rf "$proj"
mkdir -p "$proj"
cp -R "$dst/SKILL.md" "$dst/scripts" "$dst/references" "$proj/"
echo "synced project skill  .claude/skills/ppt-report-builder  (플러그인 번들 미러 — /plugin 없이 로드)"
