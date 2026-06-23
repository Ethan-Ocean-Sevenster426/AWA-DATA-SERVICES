#!/usr/bin/env bash
# One-time / re-runnable installer for ALL AWA Data Services reports on a Linux host
# (systemd). Run as root from the repo root:  sudo bash deploy/install.sh
# Safe to re-run after a `git pull` to push out new/updated report scripts.
set -euo pipefail

APP_DIR=/opt/awa-data-services
ENV_DIR=/etc/awa-data-services
ENV_FILE="$ENV_DIR/transit_report.env"     # shared env (CargoWise + Azure/SharePoint creds)
SERVICE_USER=awa

echo ">> Creating service user '$SERVICE_USER' (if missing)"
id -u "$SERVICE_USER" >/dev/null 2>&1 || useradd --system --create-home --shell /usr/sbin/nologin "$SERVICE_USER"

echo ">> Installing app to $APP_DIR"
mkdir -p "$APP_DIR/output"
cp *.py run_all.sh requirements.txt "$APP_DIR/"
cp -r data "$APP_DIR/" 2>/dev/null || true
chmod +x "$APP_DIR/run_all.sh"

echo ">> Python venv + dependencies"
[ -d "$APP_DIR/.venv" ] || python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/.venv/bin/pip" install --quiet -r "$APP_DIR/requirements.txt"
chown -R "$SERVICE_USER:$SERVICE_USER" "$APP_DIR"

echo ">> Environment file at $ENV_FILE"
mkdir -p "$ENV_DIR"
if [ ! -f "$ENV_FILE" ]; then
  cp .env.example "$ENV_FILE"
  chmod 600 "$ENV_FILE"
  echo "   -> NEW env file created. EDIT $ENV_FILE and fill in real credentials, then re-run."
else
  echo "   -> existing env file kept (creds preserved)."
fi

echo ">> Installing combined systemd units (all reports, hourly)"
cp deploy/awa-reports.service /etc/systemd/system/
cp deploy/awa-reports.timer   /etc/systemd/system/
systemctl daemon-reload

# Retire the old transit-only daily timer if present (the combined runner includes transit now)
if systemctl list-unit-files | grep -q '^awa-transit-report.timer'; then
  echo ">> Disabling old awa-transit-report.timer (transit now runs via the combined timer)"
  systemctl disable --now awa-transit-report.timer 2>/dev/null || true
fi

systemctl enable --now awa-reports.timer

echo ">> Done. Useful commands:"
echo "   sudo systemctl start awa-reports.service      # run ALL reports once now"
echo "   journalctl -u awa-reports.service -f          # watch logs"
echo "   systemctl list-timers awa-reports.timer       # next scheduled run"
