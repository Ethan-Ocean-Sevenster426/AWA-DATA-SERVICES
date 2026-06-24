#!/usr/bin/env bash
# Run every AWA data-service report and email a plain-English status summary.
# The actual orchestration + emailing lives in run_reports.py (so the report
# list, last-success tracking and the email are all in one place). This wrapper
# stays as the systemd entry point. Exit code is non-zero if any report failed.
cd "$(dirname "$0")"
PY="${PYTHON:-python3}"
exec "$PY" run_reports.py
