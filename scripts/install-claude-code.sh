#!/usr/bin/env bash
# Make the CodingCops brand skill available in every Claude Code session
# (macOS / Linux), and install the brand fonts so decks and documents display
# correctly on this computer.
#
#   ./scripts/install-claude-code.sh            # skill + fonts
#   ./scripts/install-claude-code.sh --no-fonts # skill only
#
# The skill is linked, not copied: `git pull` in this repo updates it.
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$REPO/.claude/skills/codingcops-brand"
DEST_DIR="$HOME/.claude/skills"
DEST="$DEST_DIR/codingcops-brand"

[ -f "$SRC/SKILL.md" ] || { echo "Can't find $SRC/SKILL.md. Run this from the cloned repo."; exit 1; }
mkdir -p "$DEST_DIR"

if [ -L "$DEST" ]; then
  rm "$DEST"
elif [ -e "$DEST" ]; then
  BK="$DEST.backup-$(date +%Y%m%d%H%M%S)"
  mv "$DEST" "$BK"
  echo "Moved an existing copy to $BK"
fi
ln -s "$SRC" "$DEST"
echo "✓ Skill linked: $DEST -> $SRC"

if [ "${1:-}" != "--no-fonts" ]; then
  case "$(uname -s)" in
    Darwin) FONT_DIR="$HOME/Library/Fonts" ;;
    *)      FONT_DIR="$HOME/.local/share/fonts" ;;
  esac
  mkdir -p "$FONT_DIR"
  cp -n "$SRC"/assets/fonts/*.otf "$SRC"/assets/fonts/*.ttf "$FONT_DIR"/ 2>/dev/null || true
  command -v fc-cache >/dev/null && fc-cache -f >/dev/null 2>&1 || true
  echo "✓ Fonts installed to $FONT_DIR (restart PowerPoint/Word to see them)"
fi

echo
echo "Done. In any folder, run 'claude' and ask, for example:"
echo "  /codingcops-brand Create a 10-slide pitch deck for Acme. Save it here as acme-pitch.pptx."
echo "Update later with: git -C \"$REPO\" pull --ff-only"
