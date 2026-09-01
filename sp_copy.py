"""Copy a single file from one SharePoint location to another via Microsoft Graph.

Auth: Entra (Azure AD) app registration, client credentials
(tenant id + client id + secret). Requires the application permission
Sites.ReadWrite.All (or Sites.Selected granted on both sites) with admin consent.

Setup:
    pip install msal requests

    set SP_TENANT_ID=<tenant id>
    set SP_CLIENT_ID=<client id>
    set SP_CLIENT_SECRET=<secret value>

Usage:
    python sp_copy.py
    python sp_copy.py "<source path>" "<dest folder path>"

Paths are relative to the document library root, e.g. "Reports/report.xlsx".
"""

import os
import sys
import time
from urllib.parse import quote, urlparse

import msal
import requests

# ---------------------------------------------------------------- config ----
SRC_SITE = "https://contoso.sharepoint.com/sites/SiteA"
SRC_LIBRARY = "Documents"           # document library display name
SRC_PATH = "Reports/report.xlsx"    # path inside that library

DST_SITE = "https://contoso.sharepoint.com/sites/SiteB"
DST_LIBRARY = "Documents"
DST_FOLDER = "Archive"              # "" for library root

NEW_NAME = None                     # None keeps the source filename

TENANT_ID = os.getenv("SP_TENANT_ID")
CLIENT_ID = os.getenv("SP_CLIENT_ID")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")

GRAPH = "https://graph.microsoft.com/v1.0"
# ----------------------------------------------------------------------------


def get_token():
    missing = [
        n
        for n, v in (
            ("SP_TENANT_ID", TENANT_ID),
            ("SP_CLIENT_ID", CLIENT_ID),
            ("SP_CLIENT_SECRET", CLIENT_SECRET),
        )
        if not v
    ]
    if missing:
        raise SystemExit("Missing env vars: " + ", ".join(missing))

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise SystemExit(
            f"Auth failed: {result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]


def api(session, method, url, **kwargs):
    r = session.request(method, url, **kwargs)
    if not r.ok:
        raise SystemExit(f"{method} {url} -> {r.status_code}: {r.text[:500]}")
    return r


def resolve_drive(session, site_url, library):
    """Return the driveId of `library` on `site_url`."""
    parts = urlparse(site_url)
    site_ref = f"{parts.netloc}:{parts.path.rstrip('/')}"
    site = api(session, "GET", f"{GRAPH}/sites/{site_ref}").json()

    drives = api(session, "GET", f"{GRAPH}/sites/{site['id']}/drives").json()["value"]
    for d in drives:
        if d["name"].lower() == library.lower():
            return d["id"]
    names = ", ".join(d["name"] for d in drives)
    raise SystemExit(f"Library '{library}' not found on {site_url}. Available: {names}")


def get_item(session, drive_id, path):
    enc = quote(path.strip("/"))
    return api(session, "GET", f"{GRAPH}/drives/{drive_id}/root:/{enc}").json()


def get_folder(session, drive_id, path):
    path = path.strip("/")
    if not path:
        return api(session, "GET", f"{GRAPH}/drives/{drive_id}/root").json()
    return get_item(session, drive_id, path)


def wait_for_copy(session, monitor_url, timeout=300):
    """Graph copy is async: poll the monitor URL until it settles."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        # Monitor URL is pre-authenticated; sending our header breaks it.
        r = requests.get(monitor_url)
        if not r.ok:
            raise SystemExit(f"Copy monitor -> {r.status_code}: {r.text[:300]}")
        status = r.json()
        state = status.get("status")
        if state == "completed":
            return status
        if state == "failed":
            raise SystemExit(f"Copy failed: {status.get('error')}")
        time.sleep(2)
    raise SystemExit(f"Copy did not finish within {timeout}s")


def copy_file(src_path, dst_folder):
    token = get_token()
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})

    src_drive = resolve_drive(session, SRC_SITE, SRC_LIBRARY)
    dst_drive = resolve_drive(session, DST_SITE, DST_LIBRARY)

    item = get_item(session, src_drive, src_path)
    folder = get_folder(session, dst_drive, dst_folder)

    body = {"parentReference": {"driveId": dst_drive, "id": folder["id"]}}
    if NEW_NAME:
        body["name"] = NEW_NAME

    r = api(
        session,
        "POST",
        f"{GRAPH}/drives/{src_drive}/items/{item['id']}/copy",
        json=body,
    )

    monitor = r.headers.get("Location")
    if monitor:
        wait_for_copy(session, monitor)

    name = NEW_NAME or item["name"]
    print(f"Copied {SRC_LIBRARY}/{src_path} -> {DST_LIBRARY}/{dst_folder}/{name}")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else SRC_PATH
    dst = sys.argv[2] if len(sys.argv) > 2 else DST_FOLDER
    copy_file(src, dst)
