#!/usr/bin/env bash
# One-time installer for the Transit External Report service on a Linux host (systemd).
# Run as root:  sudo bash deploy/install.sh
set -euo pipefail

APP_DIR=/opt/awa-data-services
ENV_DIR=/etc/awa-data-services
SERVICE_USER=awa

echo ">> Creating service user '$SERVICE_USER' (if missing)"
id -u "$SERVICE_USER" >/dev/null 2>&1 || useradd --system --create-home --shell /usr/sbin/nologin "$SERVICE_USER"

echo ">> Installing app to $APP_DIR"
mkdir -p "$APP_DIR/output"
cp transit_report.py requirements.txt "$APP_DIR/"
python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install --upgrade pip
"$APP_DIR/.venv/bin/pip" install -r "$APP_DIR/requirements.txt"
chown -R "$SERVICE_USER:$SERVICE_USER" "$APP_DIR"

echo ">> Setting up environment file at $ENV_DIR/transit_report.env"
mkdir -p "$ENV_DIR"
if [ ! -f "$ENV_DIR/transit_report.env" ]; then
  cp .env.example "$ENV_DIR/transit_report.env"
  chmod 600 "$ENV_DIR/transit_report.env"
  echo "   -> EDIT $ENV_DIR/transit_report.env and fill in real credentials"
fi

echo ">> Installing systemd units"
cp deploy/awa-transit-report.service /etc/systemd/system/
cp deploy/awa-transit-report.timer   /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now awa-transit-report.timer

echo ">> Done. Useful commands:"
echo "   sudo systemctl start awa-transit-report.service   # run once now"
echo "   journalctl -u awa-transit-report.service -f       # watch logs"
echo "   systemctl list-timers awa-transit-report.timer    # next run"
