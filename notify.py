#!/usr/bin/env python3
"""
AWA Data Services - email notifier (Microsoft Graph, app-only)
==============================================================
send_email(subject, html_body) sends mail through the same Azure app used for
SharePoint uploads. Requires the app to also have the **Mail.Send** application
permission (admin-consented) and a real sender mailbox in the tenant.

Env:
  AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET   (already set)
  NOTIFY_SENDER   mailbox the email is sent *from* (e.g. reports@yourtenant.com)
  NOTIFY_TO       comma-separated recipients
"""
import os, sys, json, urllib.request, urllib.parse

def _env(k, d=None):
    return os.environ.get(k, d)

def _token():
    tenant = _env("AZURE_TENANT_ID"); cid = _env("AZURE_CLIENT_ID"); sec = _env("AZURE_CLIENT_SECRET")
    body = urllib.parse.urlencode({"client_id": cid, "client_secret": sec,
                                   "scope": "https://graph.microsoft.com/.default",
                                   "grant_type": "client_credentials"}).encode()
    r = urllib.request.urlopen(urllib.request.Request(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"}))
    return json.load(r)["access_token"]

def send_email(subject, html_body):
    sender = _env("NOTIFY_SENDER")
    to = [x.strip() for x in (_env("NOTIFY_TO", "") or "").split(",") if x.strip()]
    if not sender or not to:
        print("NOTIFY_SENDER / NOTIFY_TO not set; skipping email", file=sys.stderr)
        return False
    tok = _token()
    msg = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        },
        "saveToSentItems": False,
    }
    url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(sender)}/sendMail"
    req = urllib.request.Request(url, data=json.dumps(msg).encode(), method="POST",
                                 headers={"Authorization": "Bearer " + tok,
                                          "Content-Type": "application/json"})
    urllib.request.urlopen(req)  # 202 Accepted, no body
    print(f"Status email sent to {', '.join(to)}")
    return True
