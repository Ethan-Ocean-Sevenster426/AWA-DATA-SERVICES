# AWA Data Services — Transit External Customer Report

Automated pull of the **External Customer Report MOC** from CargoWise TWD (WiseGrid)
for the Condor branches (**DOR**, **CON**), formatted to match the master template and
uploaded to SharePoint. Built to run unattended on Linux.

## What it does

1. Authenticates to CargoWise (`/Glow/auth/v2`) and resolves the branch + department contexts by code.
2. For each branch, queries the `WhsItemReceiveConsignments` OData feed and applies the
   **External Customer Report MOC** filters:
   - RCN reference has **no word starting with `s00`**
   - a package **unloaded in the last `REPORT_MONTHS` months**
   - **no departed (`DEP`) packages**
   - booking party has **no word starting with `ISCM`**
3. Resolves consignor / consignee / booking-party names + full addresses, and the
   package counts (BKD / In Warehouse / DEP / Number of Packages).
4. Writes a single **Excel Table** (`External Report` sheet) with a `Branch` column,
   real `dd-mmm-yy hh:mm` dates and `CODE - Description` service levels.
5. Uploads the workbook to SharePoint via Microsoft Graph (app-only client credentials).

> **Note on `Overs`:** the column is included (to match the template) but set to `0` —
> it is a server-side booking-vs-received overage figure that isn't exposed in the data
> feed. All other columns are reproduced exactly.

## Configuration

All config is via environment variables — see [`.env.example`](.env.example). Nothing
secret is stored in the repo.

| Variable | Purpose |
|---|---|
| `CW_USERNAME`, `CW_PASSWORD` | CargoWise login |
| `CW_BRANCH_CODES` | branches to pull (default `DOR,CON`) |
| `REPORT_MONTHS` | rolling "unloaded in last N months" window (default `12`) |
| `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` | Graph app-only auth |
| `SHAREPOINT_HOSTNAME`, `SHAREPOINT_SITE_PATH`, `SHAREPOINT_FOLDER` | upload target |
| `UPLOAD` | `true` to upload, `false` to only write the file locally |

## Run locally

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env        # then edit .env with real values
python transit_report.py
```

## Deploy as a Linux service (systemd timer)

```bash
sudo bash deploy/install.sh
# edit /etc/awa-data-services/transit_report.env with real credentials
sudo systemctl start awa-transit-report.service      # run once now
journalctl -u awa-transit-report.service -f          # logs
systemctl list-timers awa-transit-report.timer       # next scheduled run
```

Schedule lives in [`deploy/awa-transit-report.timer`](deploy/awa-transit-report.timer)
(default daily 06:00).

## Run with Docker

```bash
docker build -t awa-transit-report .
docker run --rm --env-file .env -v "$PWD/output:/app/output" awa-transit-report
```

## Run with GitHub Actions (no server needed)

[`.github/workflows/transit-report.yml`](.github/workflows/transit-report.yml) runs the
job daily on a hosted Linux runner. Add the credentials under
**Settings → Secrets and variables → Actions** (secrets for passwords/keys, variables for
the non-secret config).

## Security

- Secrets are **never** committed — `.env` is git-ignored; use systemd `EnvironmentFile`,
  Docker env, or GitHub Actions secrets in production.
- The Graph app uses **app-only** auth with `Sites.ReadWrite.All`; scope it to only what's needed.
- Recommend keeping this repository **private**.
