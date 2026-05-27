#!/usr/bin/env bash
#
# Cron wrapper for refresh.sh.
#
# Runs the Outlook -> mbox -> LanceDB pipeline. Designed for an aggressive
# schedule (every 5 minutes) so it uses flock to skip overlapping runs and
# bounds the log file to keep disk usage flat.
#
# Crontab entry:
#   */5 * * * * /home/wuttke/outlook-rag/cron-refresh.sh
#
# Notes:
#  - cron has a minimal PATH; we prepend the Windows dirs so refresh.sh can
#    invoke `powershell.exe` unqualified, and /usr/bin so flock/tee are found.
#  - LOCK is held non-blocking: a slow run (e.g. first ingest after many new
#    mails) will cause the next tick to skip rather than queue up.
#  - Logs are appended to logs/cron.log (gitignored) and truncated to the
#    last 5 MB after each run.

set -u

ROOT="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="$ROOT/logs"
mkdir -p "$LOG_DIR"
LOG="$LOG_DIR/cron.log"
LOCK="$LOG_DIR/cron-refresh.lock"

# cron PATH is minimal; refresh.sh calls powershell.exe by name.
export PATH="/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0:/mnt/c/WINDOWS/system32:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"

# Non-blocking lock: skip this tick if a previous run is still going.
exec 9>"$LOCK"
if ! flock -n 9; then
    echo "=== $(date -Is) skipped (lock held) ===" >> "$LOG"
    exit 0
fi

{
    echo
    echo "=== $(date -Is) cron-refresh start ==="
    "$ROOT/refresh.sh"
    rc=$?
    echo "=== $(date -Is) cron-refresh end rc=$rc ==="
} >> "$LOG" 2>&1

# Bound the log: if > 10 MB, keep only the last 5 MB.
if [ -f "$LOG" ] && [ "$(stat -c%s "$LOG")" -gt 10485760 ]; then
    tail -c 5242880 "$LOG" > "$LOG.tmp" && mv "$LOG.tmp" "$LOG"
fi
