# drive.py

import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# 1. CONFIGURATION
SERVICE_ACCOUNT_FILE = "buoyant-voyage-461916-r8-1e94cd664e30.json"  # ← put your key here
SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_NAME = "Tax Invoice"

# 2. AUTHENTICATE
_creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
_drive_service = build("drive", "v3", credentials=_creds)

def get_folder_id(folder_name: str = FOLDER_NAME) -> str | None:
    """Return the first matching folder ID, or None."""
    query = (
        f"name = '{folder_name}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    res = _drive_service.files().list(q=query, fields="files(id)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

def upload_bytes_to_drive(
    file_bytes: bytes,
    filename: str,
    folder_id: str,
    mime_type: str = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
) -> str:
    """
    Upload an in-memory byte sequence as a file to Google Drive.
    Returns the new file's Drive ID.
    """
    fh = io.BytesIO(file_bytes)
    media = MediaIoBaseUpload(fh, mimetype=mime_type, resumable=True)
    metadata = {"name": filename, "parents": [folder_id]}
    file = _drive_service.files().create(
        body=metadata, media_body=media, fields="id"
    ).execute()
    return file["id"]
