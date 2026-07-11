"""Google Drive uploader for the daily report.

Reads service account credentials from either:
  * GOOGLE_APPLICATION_CREDENTIALS_JSON  (env var containing full JSON string —
    used in GitHub Actions), or
  * GOOGLE_APPLICATION_CREDENTIALS  (env var containing a filesystem path —
    used for local dev).

Uploads a file into a Drive folder identified by GDRIVE_FOLDER_ID. The Drive
folder MUST be shared with the service account email as Editor.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _load_credentials() -> service_account.Credentials:
    """Load credentials from env — JSON string or file path."""
    creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if creds_json:
        info = json.loads(creds_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if creds_path:
        return service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    raise RuntimeError(
        "No Google credentials found. Set GOOGLE_APPLICATION_CREDENTIALS_JSON (JSON "
        "content) or GOOGLE_APPLICATION_CREDENTIALS (path to json file)."
    )


def _drive_service():
    return build("drive", "v3", credentials=_load_credentials(), cache_discovery=False)


def file_exists_in_folder(name: str, folder_id: str) -> dict | None:
    """Return metadata for a same-named file in the folder, or None."""
    service = _drive_service()
    # escape single quotes inside the filename to keep the query well-formed
    safe = name.replace("'", "\\'")
    query = f"name = '{safe}' and '{folder_id}' in parents and trashed = false"
    resp = service.files().list(
        q=query,
        fields="files(id, name, webViewLink)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_file(file_path: Path, folder_id: str, mime_type: str = XLSX_MIME) -> dict:
    """Upload a file into the folder. Returns {id, name, webViewLink}."""
    service = _drive_service()
    metadata = {"name": file_path.name, "parents": [folder_id]}
    media = MediaFileUpload(str(file_path), mimetype=mime_type, resumable=True)
    return service.files().create(
        body=metadata,
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()


if __name__ == "__main__":
    import sys
    folder = os.environ["GDRIVE_FOLDER_ID"]
    result = upload_file(Path(sys.argv[1]), folder)
    print(json.dumps(result, indent=2))
