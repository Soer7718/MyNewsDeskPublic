# src/sharepoint_uploader.py
import os
import requests
from pathlib import Path

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _get_token():
    tenant_id = os.getenv("GRAPH_TENANT_ID")
    client_id = os.getenv("GRAPH_CLIENT_ID")
    client_secret = os.getenv("GRAPH_CLIENT_SECRET")

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }

    resp = requests.post(url, data=data)
    if not resp.ok:
        raise RuntimeError(f"Token fejl ({resp.status_code}): {resp.text}")
    return resp.json()["access_token"]


def _graph_get(url, headers):
    resp = requests.get(url, headers=headers)
    if not resp.ok:
        raise RuntimeError(f"Graph API fejl ({resp.status_code}) på {url}:\n{resp.text}")
    return resp.json()


def _get_drive_id(headers):
    hostname = os.getenv("SP_SITE_HOSTNAME")
    site_path = os.getenv("SP_SITE_PATH")
    library = os.getenv("SP_DOC_LIBRARY")

    site = _graph_get(f"{GRAPH_BASE}/sites/{hostname}:/{site_path}", headers)
    site_id = site["id"]

    drives = _graph_get(f"{GRAPH_BASE}/sites/{site_id}/drives", headers)["value"]

    for d in drives:
        if d.get("name", "").lower() == library.lower():
            return d["id"]

    available = [d.get("name") for d in drives]
    raise RuntimeError(f"Bibliotek '{library}' ikke fundet. Tilgængelige: {available}")


def upload_to_sharepoint(local_file: str, folder: str = None):
    """
    Uploader en lokal fil til SharePoint.

    Args:
        local_file: sti til den lokale fil
        folder:     SharePoint-mappe (overstyrer SP_FOLDER_PATH fra .env hvis angivet)
    """
    token = _get_token()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _get_drive_id(headers)

    if folder is None:
        folder = os.getenv("SP_FOLDER_PATH", "")

    file_name = Path(local_file).name
    item_path = f"{folder}/{file_name}" if folder else file_name
    upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{item_path}:/content"

    with open(local_file, "rb") as f:
        headers["If-Match"] = "*"
        resp = requests.put(upload_url, headers=headers, data=f.read())

    if not resp.ok:
        raise RuntimeError(f"Upload fejl ({resp.status_code}): {resp.text}")

    return resp.json().get("webUrl")


def download_from_sharepoint(file_name: str, local_path: str, folder: str = None):
    """
    Downloader en fil fra SharePoint til en lokal sti.

    Args:
        file_name:  filnavnet på SharePoint (fx 'Impact report - standard report sample.pptx')
        local_path: lokal destination (fx 'output/impact_report.pptx')
        folder:     SharePoint-mappe (overstyrer SP_FOLDER_PATH fra .env hvis angivet)
    """
    token = _get_token()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _get_drive_id(headers)

    if folder is None:
        folder = os.getenv("SP_FOLDER_PATH", "")

    item_path = f"{folder}/{file_name}" if folder else file_name
    download_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{item_path}:/content"

    resp = requests.get(download_url, headers=headers)
    if not resp.ok:
        raise RuntimeError(f"Download fejl ({resp.status_code}): {resp.text}")

    os.makedirs(Path(local_path).parent, exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(resp.content)

    print(f"[sharepoint] Downloadet: {file_name} → {local_path}")
    return local_path
