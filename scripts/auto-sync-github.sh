#!/bin/zsh
set -euo pipefail

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
LOCK_DIR="$REPO_DIR/.git/auto-sync.lock"

cd "$REPO_DIR"

if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  exit 0
fi

trap 'rmdir "$LOCK_DIR" 2>/dev/null || true' EXIT

if [[ -n "$(git status --porcelain)" ]]; then
  git add -A

  if ! git diff --cached --quiet; then
    timestamp="$(TZ=Asia/Shanghai date '+%Y-%m-%d %H:%M:%S')"
    git commit -m "chore: auto-sync ${timestamp}"
  fi
fi

if [[ "$(git rev-list --count origin/main..HEAD)" -eq 0 ]]; then
  exit 0
fi

for dns_server in 9.9.9.9 1.1.1.1; do
  github_ip="$(dig +short github.com "@${dns_server}" | head -n 1)"
  [[ -z "$github_ip" ]] && continue

  git config --local http.curloptResolve "github.com:443:${github_ip}"
  if GIT_HTTP_LOW_SPEED_LIMIT=1 GIT_HTTP_LOW_SPEED_TIME=20 git push origin main; then
    exit 0
  fi
done

exit 1
