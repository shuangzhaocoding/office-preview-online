#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

npm run build

git add -A
if git diff --cached --quiet; then
  echo "没有需要提交的变更"
else
  if [ -n "${1:-}" ]; then
    msg="$1"
  else
    printf "请输入 commit 信息: "
    IFS= read -r msg
  fi
  if [ -z "${msg}" ]; then
    echo "commit 信息不能为空，已取消提交"
    exit 1
  fi
  git commit -m "$msg"
  echo "已提交: $msg"
fi

branch="$(git rev-parse --abbrev-ref HEAD)"

echo "推送到 origin (${branch})"
git push origin "$branch"

if git remote get-url github >/dev/null 2>&1; then
  echo "推送到 github (${branch})"
  git push github "$branch"
fi

git status -sb
