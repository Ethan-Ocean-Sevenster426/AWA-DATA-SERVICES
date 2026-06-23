#!/usr/bin/env bash
# Run every AWA data-service report in sequence. Each report is independent and
# self-contained (pull from CargoWise -> build Excel -> upload to SharePoint), so
# a failure in one does not stop the others; the overall exit code is non-zero if
# any report failed (so systemd marks the run failed and it shows in the logs).
cd "$(dirname "$0")"
PY="${PYTHON:-python3}"

SCRIPTS=(
  transit_report.py
  cycle_count_report.py
  rtu_report.py
  open_rtu_report.py
  pkg_report.py
  rcn_report.py
  condor_cleanup_report.py
  services_pending_report.py
  rcn_pending_services_report.py
  mikes_bonded_check_report.py
  unknown_received_report.py
)

rc=0
for s in "${SCRIPTS[@]}"; do
  echo "==================== $(date -u '+%Y-%m-%d %H:%M:%S')Z  $s ===================="
  if ! "$PY" "$s"; then
    echo "!!!! $s FAILED (continuing) !!!!"
    rc=1
  fi
done
echo "==================== run_all done (exit $rc) ===================="
exit $rc
