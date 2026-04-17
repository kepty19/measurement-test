#!/usr/bin/env bash
# Xcode Command Line Tools が入った Mac で実行してください。
set -euo pipefail
cd "$(dirname "$0")"

MSG="${1:-chore: sync measurement test app}"

git add -A
if git diff --staged --quiet; then
  echo "コミットする変更がありません。"
  exit 0
fi

git commit -m "$MSG"
git push origin main
echo "GitHub へ push しました。"
