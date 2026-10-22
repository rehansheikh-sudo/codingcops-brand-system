#!/usr/bin/env bash
# Build dist/codingcops-brand.zip for Claude.ai
# (Organization settings → Skills, or Settings → Capabilities → Skills → Upload skill).
# The zip contains the skill folder itself, as Claude.ai requires.
set -euo pipefail
REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS="$REPO/.claude/skills"
NAME="codingcops-brand"
OUT="$REPO/dist/$NAME.zip"

# Sanity checks: SKILL.md frontmatter name must match the folder name.
head -5 "$SKILLS/$NAME/SKILL.md" | grep -q "^name: $NAME\$" \
  || { echo "SKILL.md 'name:' must be '$NAME'"; exit 1; }
python3 "$SKILLS/$NAME/scripts/check_brand.py" --self-test >/dev/null
python3 -c "import json,sys; json.load(open(sys.argv[1]))" "$SKILLS/$NAME/tokens/brand-tokens.json"

mkdir -p "$REPO/dist"
rm -f "$OUT"
( cd "$SKILLS" && zip -qr "$OUT" "$NAME" -x "*/.DS_Store" "*/__pycache__/*" "*.pyc" "*/.gitkeep" )
echo "Built $OUT ($(du -h "$OUT" | cut -f1))"
