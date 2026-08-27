#!/bin/bash
# One-command refresh: fetch -> blend -> publish, so the sheet's
# "Draft Tools > Refresh data from GitHub" picks up the new numbers.
#
#   ./refresh.sh          # incremental: only missing/stale players
#   ./refresh.sh --full   # refetch everything from the API
set -e
cd "$(dirname "$0")"

python3 fetch_fp.py "${1:---cache}"
python3 blend.py

if git diff --quiet -- out/blended.csv out/idp.csv; then
  echo "No projection changes to publish."
else
  git add out/blended.csv out/idp.csv
  git commit -q -m "Refresh projections $(date +%Y-%m-%d\ %H:%M)"
  git push -q
  echo "Published. In the sheet: Draft Tools > Refresh data from GitHub."
fi
