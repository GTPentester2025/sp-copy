"""Copy a single file from one SharePoint location to another.

App-only auth (client id + secret). Same-site copies use SharePoint's
server-side CopyTo (no bytes cross the network). Cross-site copies stream the
file down and back up.

Setup:
    pip install Office365-REST-Python-Client
    set SP_CLIENT_ID=<client id>
    set SP_CLIENT_SECRET=<client secret>

Usage:
    python sp_copy.py
    python sp_copy.py "<source server-relative path>" "<dest server-relative folder>"
"""

import io
import os
import sys

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# ---------------------------------------------------------------- config ----
SRC_SITE = "https://contoso.sharepoint.com/sites/SiteA"
SRC_FILE = "/sites/SiteA/Shared Documents/report.xlsx"

DST_SITE = "https://contoso.sharepoint.com/sites/SiteB"
DST_FOLDER = "/sites/SiteB/Shared Documents/Archive"

OVERWRITE = True

CLIENT_ID = os.getenv("SP_CLIENT_ID")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET")
# ----------------------------------------------------------------------------


def ctx_for(site_url):
    if not (CLIENT_ID and CLIENT_SECRET):
        raise SystemExit("Set SP_CLIENT_ID and SP_CLIENT_SECRET.")
    return ClientContext(site_url).with_credentials(
        ClientCredential(CLIENT_ID, CLIENT_SECRET)
    )


def copy_file(src_site, src_file, dst_site, dst_folder, overwrite=True):
    name = os.path.basename(src_file)
    dst_path = f"{dst_folder.rstrip('/')}/{name}"
    src = ctx_for(src_site)

    if src_site.rstrip("/").lower() == dst_site.rstrip("/").lower():
        # Same site: let SharePoint do the copy server-side.
        src.web.get_file_by_server_relative_url(src_file).copyto(
            dst_folder, overwrite
        ).execute_query()
    else:
        buf = io.BytesIO()
        src.web.get_file_by_server_relative_url(src_file).download(
            buf
        ).execute_query()

        dst = ctx_for(dst_site)
        dst.web.get_folder_by_server_relative_url(dst_folder).upload_file(
            name, buf.getvalue()
        ).execute_query()

    print(f"Copied {src_file} -> {dst_path}")
    return dst_path


if __name__ == "__main__":
    src_file = sys.argv[1] if len(sys.argv) > 1 else SRC_FILE
    dst_folder = sys.argv[2] if len(sys.argv) > 2 else DST_FOLDER
    copy_file(SRC_SITE, src_file, DST_SITE, dst_folder, OVERWRITE)
