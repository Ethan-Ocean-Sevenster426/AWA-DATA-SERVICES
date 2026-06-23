#!/usr/bin/env python3
"""
Shared SharePoint uploader (Microsoft Graph, app-only client-credentials).

upload(local_path, folder) puts the file into:
    {SHAREPOINT_ROOT}/{folder}/{filename}
on the configured SharePoint site (default root: "Clients/ISCM").

Before uploading, any existing file of the same name in that folder is MOVED into
an "Archive" subfolder with a timestamp suffix (e.g. "KPIU_20260623-051507.xlsx"),
so each run preserves the previous version. Archived copies beyond
SHAREPOINT_ARCHIVE_KEEP (default 168) are pruned, newest kept.

Reads AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET / SHAREPOINT_HOSTNAME
/ SHAREPOINT_SITE_PATH (and optional SHAREPOINT_ROOT, SHAREPOINT_ARCHIVE_KEEP) from
the environment. No secrets are stored in code.
"""
import os, sys, json, logging, datetime
import urllib.request, urllib.parse, urllib.error

GRAPH = "https://graph.microsoft.com/v1.0"
log = logging.getLogger("sp_upload")

def _env(k, d=None):
    return os.environ.get(k, d)

def _http(url, data=None, headers=None, method=None, timeout=180):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    return urllib.request.urlopen(req, timeout=timeout)

def _enc(path):
    """URL-encode a drive-relative path segment-by-segment (keep the slashes)."""
    return "/".join(urllib.parse.quote(seg) for seg in path.split("/"))

def _connect():
    tenant = _env("AZURE_TENANT_ID"); cid = _env("AZURE_CLIENT_ID"); sec = _env("AZURE_CLIENT_SECRET")
    host = _env("SHAREPOINT_HOSTNAME"); site = _env("SHAREPOINT_SITE_PATH")
    for k, v in {"AZURE_TENANT_ID": tenant, "AZURE_CLIENT_ID": cid, "AZURE_CLIENT_SECRET": sec,
                 "SHAREPOINT_HOSTNAME": host, "SHAREPOINT_SITE_PATH": site}.items():
        if not v:
            sys.exit(f"FATAL: UPLOAD=true but {k} is not set")
    body = urllib.parse.urlencode({"client_id": cid, "client_secret": sec,
                                   "scope": "https://graph.microsoft.com/.default",
                                   "grant_type": "client_credentials"}).encode()
    tok = json.load(_http(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
                          data=body, method="POST",
                          headers={"Content-Type": "application/x-www-form-urlencoded"}))["access_token"]
    H = {"Authorization": "Bearer " + tok}
    site_obj = json.load(_http(f"{GRAPH}/sites/{host}:{site}", headers=H))
    drive = json.load(_http(f"{GRAPH}/sites/{site_obj['id']}/drive", headers=H))
    return H, drive["id"]

def _archive_existing(H, did, rel, fname):
    """Move rel/fname into rel/Archive (timestamped) if it exists; prune old archives."""
    stem, ext = os.path.splitext(fname)
    try:
        item = json.load(_http(f"{GRAPH}/drives/{did}/root:/{_enc(rel + '/' + fname)}", headers=H))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return                      # nothing to archive
        raise
    # ensure the Archive folder exists
    try:
        af = json.load(_http(f"{GRAPH}/drives/{did}/root:/{_enc(rel + '/Archive')}", headers=H))
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
        af = json.load(_http(f"{GRAPH}/drives/{did}/root:/{_enc(rel)}:/children",
                             data=json.dumps({"name": "Archive", "folder": {},
                                              "@microsoft.graph.conflictBehavior": "fail"}).encode(),
                             method="POST", headers={**H, "Content-Type": "application/json"}))
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    _http(f"{GRAPH}/drives/{did}/items/{item['id']}",
          data=json.dumps({"parentReference": {"id": af["id"]}, "name": f"{stem}_{ts}{ext}"}).encode(),
          method="PATCH", headers={**H, "Content-Type": "application/json"})
    log.info("  archived previous version -> Archive/%s_%s%s", stem, ts, ext)
    # prune old archives (keep newest SHAREPOINT_ARCHIVE_KEEP)
    try:
        keep = int(_env("SHAREPOINT_ARCHIVE_KEEP", "168"))
        kids = json.load(_http(f"{GRAPH}/drives/{did}/items/{af['id']}/children?$select=id,name&$top=999",
                               headers=H)).get("value", [])
        mine = sorted([c for c in kids if c["name"].startswith(stem + "_") and c["name"].endswith(ext)],
                      key=lambda c: c["name"], reverse=True)
        for c in mine[keep:]:
            _http(f"{GRAPH}/drives/{did}/items/{c['id']}", method="DELETE", headers=H)
    except Exception as e:
        log.warning("  archive prune skipped (%s)", e)

def upload(local_path, folder):
    root = (_env("SHAREPOINT_ROOT", "Clients/ISCM") or "").strip("/")
    H, did = _connect()
    rel = "/".join(p for p in (root, (folder or "").strip("/")) if p)
    fname = os.path.basename(local_path)
    _archive_existing(H, did, rel, fname)
    url = f"{GRAPH}/drives/{did}/root:/{_enc(rel + '/' + fname)}:/content"
    with open(local_path, "rb") as f:
        content = f.read()
    res = json.load(_http(url, data=content, method="PUT",
                          headers={**H, "Content-Type": "application/octet-stream"}))
    log.info("Uploaded: %s (%s bytes)", res.get("webUrl"), res.get("size"))
