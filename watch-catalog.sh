#!/bin/bash
# Watch all catalog sources and regenerate their manifests on every change.
# Keep this running in a terminal — Ctrl+C to stop.
#
#   catalog/        -> catalog-manifest.js   (update-catalog.sh)
#   Windows_Pic/    -> pics-manifest.js       (update-pics.py)
#   Macbook_Pic/    -> pics-manifest.js       (update-pics.py)
#   Softwares.xlsx  -> softwares-manifest.js  (update-softwares.py)
#
# Requires fswatch:  brew install fswatch

set -euo pipefail
cd "$(dirname "$0")"

if ! command -v fswatch >/dev/null 2>&1; then
  echo "Error: fswatch not found. Install with:  brew install fswatch" >&2
  exit 1
fi

regenerate_all() {
  ./update-catalog.sh
  python3 update-pics.py
  python3 update-softwares.py
}

# Paths to watch (only those that exist).
WATCH_PATHS=()
for p in catalog Windows_Pic Macbook_Pic Softwares.xlsx; do
  [ -e "$p" ] && WATCH_PATHS+=("$p")
done

# Run once at startup so every manifest is current.
regenerate_all

echo "→ Watching ${WATCH_PATHS[*]} for changes…  (Ctrl+C to stop)"
# --latency 1: debounce 1 second so a burst of file events triggers one rebuild.
fswatch -o --latency 1 "${WATCH_PATHS[@]}" | while read -r _; do
  regenerate_all
done
