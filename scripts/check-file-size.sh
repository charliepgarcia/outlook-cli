#!/usr/bin/env bash
# Hard gate: no source file may exceed MAX_LOC lines.
# At the ceiling, extract a whole concern into a new file.
# Never raise the limit, never just shave lines.
set -euo pipefail

MAX_LOC="${MAX_LOC:-350}"
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

fail=0
while IFS= read -r file; do
  lines=$(wc -l < "$file" | tr -d ' ')
  if [ "$lines" -gt "$MAX_LOC" ]; then
    printf 'FAIL  %5s LOC  %s\n' "$lines" "$file"
    fail=1
  fi
done < <(find src -type f -name '*.ts' 2>/dev/null | sort)

if [ "$fail" -ne 0 ]; then
  echo ""
  echo "File-size gate FAILED (ceiling ${MAX_LOC} LOC)."
  echo "Fix by extracting a concern into a new file - not by raising MAX_LOC."
  exit 1
fi

count=$(find src -type f -name '*.ts' 2>/dev/null | wc -l | tr -d ' ')
echo "file-size gate OK - ${count} source files, all <= ${MAX_LOC} LOC"
