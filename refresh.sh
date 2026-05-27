#!/usr/bin/env bash
# Full refresh: drive Outlook export (Windows) to convergence, then ingest (WSL).
#   ./refresh.sh                 normal run
#   ./refresh.sh --max-passes 5  cap export loop
#   ./refresh.sh --no-ingest     skip step 2

set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
PS1_WIN="$(wslpath -w "$ROOT/outlook_export.ps1")"
LOG_DIR="$ROOT/logs"
mkdir -p "$LOG_DIR"
STAMP="$(date +%Y%m%d-%H%M%S)"

MAX_PASSES=10
DO_INGEST=1
while [[ $# -gt 0 ]]; do
    case "$1" in
        --max-passes) MAX_PASSES="$2"; shift 2 ;;
        --no-ingest)  DO_INGEST=0; shift ;;
        *) echo "unknown arg: $1" >&2; exit 2 ;;
    esac
done

echo "=== Step 1: Outlook export (loop until 0 new, max $MAX_PASSES passes) ==="
total_new=0
for pass in $(seq 1 "$MAX_PASSES"); do
    log="$LOG_DIR/export-$STAMP-pass$pass.log"
    echo "--- pass $pass --- (log: $log)"
    powershell.exe -ExecutionPolicy Bypass -File "$PS1_WIN" 2>&1 | tee "$log"
    # last "Done. New items this run: N" line
    new=$(grep -oE 'New items this run: [0-9]+' "$log" | tail -1 | grep -oE '[0-9]+$' || echo 0)
    echo "    pass $pass: $new new items"
    total_new=$(( total_new + new ))
    if [[ "$new" -eq 0 ]]; then
        echo "--- converged after $pass pass(es) ---"
        break
    fi
done
echo "=== Step 1 done. Total new items written: $total_new ==="

if [[ "$DO_INGEST" -eq 0 ]]; then
    echo "--no-ingest set; stopping."
    exit 0
fi

echo
echo "=== Step 2: LanceDB ingest ==="
ingest_log="$LOG_DIR/ingest-$STAMP.log"
# shellcheck disable=SC1091
source "$ROOT/.venv/bin/activate"
cd "$ROOT"
time python3 -u ingest.py --batch 16 2>&1 | tee "$ingest_log"
echo "=== Step 2 done. Log: $ingest_log ==="
