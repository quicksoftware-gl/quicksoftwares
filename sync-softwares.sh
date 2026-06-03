#!/bin/bash
# Regenerate softwares-manifest.js from Softwares.xlsx and, if it changed,
# commit + push so the live site picks up the edits.
#
# Wired into cron (every 10 minutes) by the install line in this repo's README/
# cron block. Run manually with: ./sync-softwares.sh
set -euo pipefail

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$REPO_DIR"

# macOS GUI processes have a fuller PATH than cron. Make sure python3 + git
# resolve when this is launched from launchd/cron.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

LOG_TAG="[$(date '+%Y-%m-%d %H:%M:%S')] sync-softwares"

./update-softwares.py

# Stage the manifest, then ask git if there's actually anything to commit.
# Staging first lets us catch untracked-new-file AND modified cases uniformly.
git add softwares-manifest.js
if git diff --cached --quiet -- softwares-manifest.js; then
  echo "$LOG_TAG no manifest changes"
  exit 0
fi

git -c user.name="quicksoftwares-bot" \
    -c user.email="quicksoftwares.global@gmail.com" \
    commit -m "chore: refresh softwares-manifest.js from Softwares.xlsx"

# Only push if a remote is configured. Suppress non-zero exit if push fails so
# the cron keeps running — the next tick will retry.
if git remote get-url origin >/dev/null 2>&1; then
  if git push origin HEAD; then
    echo "$LOG_TAG pushed"
  else
    echo "$LOG_TAG push failed (will retry next run)" >&2
  fi
else
  echo "$LOG_TAG no origin remote; committed locally only"
fi
