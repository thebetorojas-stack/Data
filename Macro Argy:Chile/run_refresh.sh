#!/usr/bin/env bash
# One-click refresh — pulls Bloomberg + Haver deltas, regenerates both workbooks.
set -euo pipefail
cd "$(dirname "$0")"

echo
echo "=== EM Macro & Credit refresh ==="
echo

python -m scripts.refresh_all "$@"

echo
echo "Workbooks ready in ./outputs/"
