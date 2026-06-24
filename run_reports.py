#!/usr/bin/env python3
"""
AWA Data Services - orchestrator + plain-English status email
=============================================================
Runs every report in sequence (independent - one failing does not stop the rest),
remembers the last successful pull time per report, and emails a human-readable
status summary at the end. Exit code is non-zero if any report failed.

Run this instead of run_all.sh (run_all.sh now just calls it). Env for the email:
NOTIFY_SENDER, NOTIFY_TO (see notify.py). If those aren't set, it still runs and
just skips the email.
"""
import os, sys, re, json, subprocess, datetime

HERE = os.path.dirname(os.path.abspath(__file__))

def _load_dotenv(path=None):
    path = path or os.path.join(HERE, ".env")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())
_load_dotenv()

PY = os.environ.get("PYTHON", sys.executable)
STATE_FILE = os.path.join(os.environ.get("OUTPUT_DIR", os.path.join(HERE, "output")), "_status.json")
TZ = "UTC"

# script -> friendly name (order = run order)
REPORTS = [
    ("transit_report.py",              "External Customer Report (DOR+CON)"),
    ("cycle_count_report.py",          "Cycle Count (CON)"),
    ("rtu_report.py",                  "KPIU / RTU (DOR+CON)"),
    ("open_rtu_report.py",             "Open RTU's (DOR+CON)"),
    ("pkg_report.py",                  "PKG - No Consignment (DOR+CON)"),
    ("rcn_report.py",                  "RCN - No Booking Party (DOR+CON)"),
    ("condor_cleanup_report.py",       "Condor Booking Party Cleanup (CON)"),
    ("services_pending_report.py",     "Services Pending - Freight On Hand (DOR+CON)"),
    ("rcn_pending_services_report.py", "Services Pending (DOR+CON)"),
    ("mikes_bonded_check_report.py",   "Mike's Bonded Check (DOR)"),
    ("unknown_received_report.py",     "Unknown Received Report (CON)"),
    ("dtu_report.py",                  "DTU (CCC)"),
    ("open_dtu_report.py",             "Open DTU's (CCC)"),
]

def now():
    return datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)

def load_state():
    try:
        return json.load(open(STATE_FILE, encoding="utf-8"))
    except Exception:
        return {}

def save_state(s):
    os.makedirs(os.path.dirname(STATE_FILE) or ".", exist_ok=True)
    json.dump(s, open(STATE_FILE, "w", encoding="utf-8"), indent=1)

def parse_rows(out):
    m = re.findall(r"\((\d[\d,]*)\s+rows\)", out)
    return int(m[-1].replace(",", "")) if m else None

def plain_error(out):
    """Translate the tail of the log into one human sentence."""
    low = out.lower()
    if "login failed" in low or "claim/staff" in low:
        return "Could not log in to CargoWise (check the username/password)."
    if "required environment variable" in low:
        m = re.search(r"required environment variable (\w+)", out)
        return f"A setting is missing on the server: {m.group(1) if m else 'see logs'}."
    if "branch code" in low and "not found" in low:
        return "A CargoWise branch in the config wasn't found for this login."
    if "department" in low and "not found" in low:
        return "The CargoWise department in the config wasn't found."
    if "401" in out:
        return "CargoWise rejected the session (login expired and re-login failed)."
    if "graph" in low and ("403" in out or "401" in out):
        return "SharePoint upload was refused (permissions/credentials)."
    if "timed out" in low or "timeout" in low:
        return "A request timed out (CargoWise or SharePoint was slow/unreachable)."
    # fallback: last non-empty line
    lines = [l.strip() for l in out.strip().splitlines() if l.strip()]
    return ("Unexpected error: " + lines[-1]) if lines else "Unknown error (no output)."

def fmt_when(iso):
    if not iso:
        return "never"
    try:
        dt = datetime.datetime.fromisoformat(iso)
        return dt.strftime("%d %b %Y %H:%M UTC")
    except Exception:
        return iso

def run_one(script):
    p = subprocess.run([PY, os.path.join(HERE, script)], capture_output=True, text=True, cwd=HERE)
    out = (p.stdout or "") + ("\n" + p.stderr if p.stderr else "")
    return p.returncode == 0, out

def main():
    state = load_state()
    started = now()
    results = []
    for script, name in REPORTS:
        print(f"==== {now().isoformat()} {script} ====", flush=True)
        ok, out = run_one(script)
        print(out, flush=True)
        rows = parse_rows(out)
        st = state.get(script, {})
        if ok:
            st["last_success"] = now().isoformat()
            st["rows"] = rows if rows is not None else st.get("rows")
            err = None
        else:
            err = plain_error(out)
        state[script] = st
        results.append({"script": script, "name": name, "ok": ok, "rows": rows,
                        "last_success": st.get("last_success"),
                        "last_rows": st.get("rows"), "error": err})
    save_state(state)

    ok_n = sum(1 for r in results if r["ok"])
    fail_n = len(results) - ok_n
    finished = now()
    overall = "ALL OK" if fail_n == 0 else f"{fail_n} FAILED"
    subj = f"AWA Reports - {overall} - {finished.strftime('%d %b %Y %H:%M UTC')}"

    rowsf = lambda r: ("-" if r is None else f"{r:,}")
    lines = []
    lines.append(f"<p>Here is the status of your CargoWise &rarr; SharePoint reports.</p>")
    lines.append(f"<p><b>Run finished:</b> {finished.strftime('%d %b %Y %H:%M UTC')}<br>"
                 f"<b>Overall:</b> {ok_n} of {len(results)} reports updated successfully.</p>")
    lines.append("<table cellpadding='6' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px'>")
    lines.append("<tr style='background:#f0f0f0'><th align='left'>Report</th><th align='left'>Status</th>"
                 "<th align='right'>Rows</th><th align='left'>Last successful pull</th></tr>")
    for r in results:
        if r["ok"]:
            status = "<span style='color:#137333'>&#10004; Updated</span>"
            rows = rowsf(r["rows"])
            when = fmt_when(r["last_success"])
            note = ""
        else:
            status = "<span style='color:#c5221f'>&#10008; Failed</span>"
            rows = rowsf(r["last_rows"]) + " (last good)"
            when = fmt_when(r["last_success"])
            note = f"<br><span style='color:#c5221f'>{r['error']}</span>"
        bg = "#ffffff" if r["ok"] else "#fdecea"
        lines.append(f"<tr style='background:{bg}'><td>{r['name']}{note}</td><td>{status}</td>"
                     f"<td align='right'>{rows}</td><td>{when}</td></tr>")
    lines.append("</table>")
    if fail_n:
        lines.append("<p>The failed reports above still show their <i>last good</i> data on SharePoint; "
                     "they'll be retried on the next hourly run.</p>")
    html = "".join(lines)

    print(subj)
    try:
        import notify
        notify.send_email(subj, html)
    except Exception as e:
        print(f"Could not send status email: {e}", file=sys.stderr)

    sys.exit(1 if fail_n else 0)

if __name__ == "__main__":
    main()
