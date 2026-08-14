#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

npm run build

git add -A
if git diff --cached --quiet; then
  echo "没有需要提交的变更"
  exit 0
fi

msg="${1:-chore: rebuild dist}"
git commit -m "$msg"

echo "已提交: $msg"
git status -sb
