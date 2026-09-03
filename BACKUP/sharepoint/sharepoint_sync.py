"""
sharepoint_sync.py
Downloads the tracking Excel file from SharePoint/OneDrive using an Azure AD
app registration (client credentials flow), so the dashboard can run
unattended on any machine -- no OneDrive desktop sync client required.

Two configuration modes are supported -- use whichever matches your case:

MODE 1 -- Shared link (use this when someone else shared the file with you,
and you don't know/control the owner's site URL or folder structure -- this
is the common case for a file living in another person's OneDrive):
    SHAREPOINT_TENANT_ID        Azure AD tenant ID
    SHAREPOINT_CLIENT_ID        App registration (application) client ID
    SHAREPOINT_CLIENT_SECRET    App registration client secret VALUE (not the Secret ID)
    SHAREPOINT_SHARE_LINK       The full sharing link URL (e.g. the
                                 https://bosch-my.sharepoint.com/:x:/... link)

    Requires Microsoft Graph Application permission Files.Read.All (or
    Sites.Read.All) with admin consent granted -- NOT the SharePoint resource
    permission used by Mode 2.

MODE 2 -- Direct site + file path (use this for a file in a Team Site
document library that you control the path for):
    SHAREPOINT_TENANT_ID
    SHAREPOINT_CLIENT_ID
    SHAREPOINT_CLIENT_SECRET
    SHAREPOINT_SITE_URL         e.g. https://boschtenant.sharepoint.com/sites/YourSite
    SHAREPOINT_FILE_URL         Server-relative path to the file, e.g.
                                 /sites/YourSite/Shared Documents/AI Champions Initiative Report.xlsx

If neither mode's variables are fully set, is_configured() returns False and
app.py falls back to reading EXCEL_SOURCE_PATH as a plain local file
(unchanged behavior).
"""

import base64
import os
import tempfile
from pathlib import Path

import msal
import requests
from office365.runtime.auth.token_response import TokenResponse
from office365.sharepoint.client_context import ClientContext


class SharePointConfigError(Exception):
    """Raised when required SharePoint environment variables are missing."""


class SharePointDownloadError(Exception):
    """Raised when auth or download fails, with a user-friendly message."""


COMMON_VARS = ["SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET"]
SHARE_LINK_VARS = COMMON_VARS + ["SHAREPOINT_SHARE_LINK"]
SITE_PATH_VARS = COMMON_VARS + ["SHAREPOINT_SITE_URL", "SHAREPOINT_FILE_URL"]


def _mode() -> str | None:
    """Returns 'share_link', 'site_path', or None if neither mode is fully configured."""
    if all(os.environ.get(v) for v in SHARE_LINK_VARS):
        return "share_link"
    if all(os.environ.get(v) for v in SITE_PATH_VARS):
        return "site_path"
    return None


def is_configured() -> bool:
    """True if either Mode 1 or Mode 2's env vars are fully set."""
    return _mode() is not None


def _get_proxies() -> dict:
    """Reads HTTPS_PROXY / HTTP_PROXY from the environment (e.g. a local CNTLM
    proxy at http://127.0.0.1:3128) and returns a requests-style proxies
    dict. Returns {} if none are set, in which case no proxy is used."""
    proxies = {}
    https_proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    http_proxy = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy")
    if https_proxy:
        proxies["https"] = https_proxy
    if http_proxy:
        proxies["http"] = http_proxy
    return proxies


def _acquire_graph_token(tenant_id: str, client_id: str, client_secret: str, proxies: dict) -> str:
    msal_app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        proxies=proxies or None,
    )
    token = msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        error_desc = token.get("error_description", "Unknown auth error")
        raise SharePointDownloadError(f"Azure AD token request failed: {error_desc}")
    return token["access_token"]


def _encode_share_url(share_url: str) -> str:
    """Encodes a sharing URL into the 'shareId' token Microsoft Graph's
    /shares/{shareId} endpoint expects. Per Microsoft's documented scheme:
    base64-encode the URL, convert to unpadded URL-safe base64, prefix 'u!'."""
    b64 = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    b64_url_safe = b64.replace("/", "_").replace("+", "-").rstrip("=")
    return f"u!{b64_url_safe}"


def _download_via_share_link(local_path: Path):
    tenant_id = os.environ["SHAREPOINT_TENANT_ID"]
    client_id = os.environ["SHAREPOINT_CLIENT_ID"]
    client_secret = os.environ["SHAREPOINT_CLIENT_SECRET"]
    share_link = os.environ["SHAREPOINT_SHARE_LINK"]
    proxies = _get_proxies()

    access_token = _acquire_graph_token(tenant_id, client_id, client_secret, proxies)
    share_id = _encode_share_url(share_link)

    url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers, proxies=proxies or None, stream=True, timeout=60)

    if resp.status_code != 200:
        raise SharePointDownloadError(
            f"Microsoft Graph returned HTTP {resp.status_code} when resolving the share link.\n"
            f"Response: {resp.text[:500]}"
        )

    fd, tmp_path = tempfile.mkstemp(dir=local_path.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        os.replace(tmp_path, local_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def _sharepoint_resource_root(site_url: str) -> str:
    """Azure AD's SharePoint app-only resource principal is always the root
    domain (e.g. https://bosch-my.sharepoint.com), never a specific site or
    personal-drive path -- using the full site_url as the token scope causes
    AADSTS500011 'resource principal not found' even when everything else
    (permissions, consent, secret) is correct."""
    from urllib.parse import urlparse
    parsed = urlparse(site_url)
    return f"{parsed.scheme}://{parsed.netloc}"


def _download_via_site_path(local_path: Path):
    tenant_id = os.environ["SHAREPOINT_TENANT_ID"]
    client_id = os.environ["SHAREPOINT_CLIENT_ID"]
    client_secret = os.environ["SHAREPOINT_CLIENT_SECRET"]
    site_url = os.environ["SHAREPOINT_SITE_URL"]
    file_url = os.environ["SHAREPOINT_FILE_URL"]
    proxies = _get_proxies()
    resource_root = _sharepoint_resource_root(site_url)

    def _acquire_token():
        msal_app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            proxies=proxies or None,
        )
        token = msal_app.acquire_token_for_client(scopes=[f"{resource_root}/.default"])
        if "access_token" not in token:
            error_desc = token.get("error_description", "Unknown auth error")
            raise SharePointDownloadError(f"Azure AD token request failed: {error_desc}")
        return TokenResponse.from_json(token)

    ctx = ClientContext(site_url).with_access_token(_acquire_token)

    fd, tmp_path = tempfile.mkstemp(dir=local_path.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            ctx.web.get_file_by_server_relative_url(file_url).download(f).execute_query()
        os.replace(tmp_path, local_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def download_excel_from_sharepoint(local_path: str) -> str:
    """Downloads the configured file to local_path, using whichever mode
    (share_link or site_path) has its env vars fully set.

    Downloads to a temp file first and only replaces local_path on success,
    so a failed/partial download never corrupts the last known-good copy
    that the dashboard is currently reading.

    Returns local_path on success. Raises SharePointConfigError or
    SharePointDownloadError with a message suitable for showing in the UI.
    """
    mode = _mode()
    if mode is None:
        raise SharePointConfigError(
            "SharePoint isn't fully configured. Set either the SHAREPOINT_SHARE_LINK "
            "variables (for a file someone else shared with you) or the "
            "SHAREPOINT_SITE_URL / SHAREPOINT_FILE_URL variables (for a file in a "
            "Team Site you control the path for)."
        )

    local_path = Path(local_path)
    local_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        if mode == "share_link":
            _download_via_share_link(local_path)
        else:
            _download_via_site_path(local_path)
        return str(local_path)

    except SharePointDownloadError:
        raise
    except Exception as e:
        proxies = _get_proxies()
        proxy_note = (
            f"Proxy detected and used: {proxies}"
            if proxies else
            "No HTTPS_PROXY/HTTP_PROXY found in this process's environment -- "
            "if you're behind CNTLM, this is almost certainly the actual problem."
        )
        raise SharePointDownloadError(
            f"Couldn't download the file (mode: {mode}).\n\n{proxy_note}\n\nDetails: {e}"
        )