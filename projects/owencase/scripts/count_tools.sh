#!/usr/bin/env bash
# Count MCP tools registered in the server.
# Usage: bash scripts/count_tools.sh

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CATALOG="$ROOT/src/tools/catalog.py"

echo "=== Total ==="
total=$(grep -c "^[[:space:]]*ToolSpec('ppt_com\|^[[:space:]]*ToolSpec('tools" "$CATALOG")
printf "%4d  tools\n" "$total"

echo ""
echo "=== README states ==="
readme="$ROOT/README.md"
stated=$(sed -nE 's/.*\|[[:space:]]*\*\*([0-9]+)\*\*[[:space:]]*\|.*/\1/p' "$readme" | head -1)
stated=${stated:-?}
printf "%4s  tools (README.md)\n" "$stated"

if [ "$total" -ne "$stated" ] 2>/dev/null; then
  echo ""
  echo "WARNING: Mismatch! Update README.md (and README_ja.md) to reflect the actual count."
fi
