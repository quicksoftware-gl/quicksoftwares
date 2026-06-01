#!/bin/bash
# Watch catalog/ and re-run update-catalog.sh on every change.
# Keep this running in a terminal — Ctrl+C to stop.
#
# Requires fswatch:  brew install fswatch

set -euo pipefail
cd "$(dirname "$0")"

if ! command -v fswatch >/dev/null 2>&1; then
  echo "Error: fswatch not found. Install with:  brew install fswatch" >&2
  exit 1
fi

# Run once at startup so the manifest is current.
./update-catalog.sh

echo "→ Watching catalog/ for changes…  (Ctrl+C to stop)"
# --latency 1: debounce 1 second; --event-flags: ignore noisy events
fswatch -o catalog/ | while read -r _; do
  ./update-catalog.sh
done
