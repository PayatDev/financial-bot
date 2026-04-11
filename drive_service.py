import os
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")


def get_drive_service():
    sa_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(sa_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def upload_file(local_path: str, filename: str) -> str:
    """Upload file to Drive folder — return file id"""
    service = get_drive_service()
    meta = {"name": filename, "parents": [FOLDER_ID]}
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    f = service.files().create(body=meta, media_body=media, fields="id").execute()
    file_id = f.get("id")
    print(f"✅ Uploaded → {filename} (id: {file_id})")
    return file_id


def rename_file(file_id: str, new_name: str):
    """Rename file in Drive"""
    service = get_drive_service()
    service.files().update(fileId=file_id, body={"name": new_name}).execute()
    print(f"✅ Renamed → {new_name}")
