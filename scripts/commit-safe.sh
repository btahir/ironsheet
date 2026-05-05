#!/usr/bin/env bash
set -euo pipefail

message="${1:-}"

if [[ -z "$message" ]]; then
  echo 'usage: npm run commit:safe -- "commit message"' >&2
  exit 2
fi

git rev-parse --is-inside-work-tree >/dev/null

git add -A

# The planning spec is useful locally, but should not be part of normal product commits.
git reset -- IRONSHEET_SPEC.md >/dev/null 2>&1 || true

if git diff --cached --quiet; then
  echo "commit:safe: no staged changes to commit"
  exit 0
fi

npm run verify
git commit -m "$message"

