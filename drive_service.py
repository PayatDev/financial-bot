import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

DRIVE_ROOT_FOLDER_ID = os.environ.get("DRIVE_ROOT_FOLDER_ID")
FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")

def create_folder(folder_name: str) -> str:
    service = get_drive_service()
    folder = service.files().create(
        body={
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [DRIVE_ROOT_FOLDER_ID]
        },
        fields="id"
    ).execute()
    folder_id = folder.get("id")
    print(f"✅ Created folder → {folder_name} (id: {folder_id})")
    return folder_id

def upload_file_to_folder(local_path: str, filename: str, folder_id: str) -> str:
    service = get_drive_service()
    ext = os.path.splitext(filename)[1].lower()
    mimetype = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pdf": "application/pdf",
    }.get(ext, "application/octet-stream")
    media = MediaFileUpload(local_path, mimetype=mimetype, resumable=True)
    f = service.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media,
        fields="id"
    ).execute()
    file_id = f.get("id")
    print(f"✅ Uploaded → {filename} (id: {file_id})")
    return file_id

def get_drive_service():
    creds = Credentials(
        token=None,
        refresh_token=os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN"),
        client_id=os.environ.get("GOOGLE_OAUTH_CLIENT_ID"),
        client_secret=os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET"),
        token_uri="https://oauth2.googleapis.com/token"
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds)


def upload_file(local_path: str, filename: str) -> str:
    service = get_drive_service()
    meta = {"name": filename, "parents": [FOLDER_ID]}
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )
    f = service.files().create(
        body=meta,
        media_body=media,
        fields="id"
    ).execute()
    file_id = f.get("id")
    print(f"✅ Uploaded → {filename} (id: {file_id})")
    return file_id


def rename_file(file_id: str, new_name: str):
    service = get_drive_service()
    service.files().update(fileId=file_id, body={"name": new_name}).execute()
    print(f"✅ Renamed → {new_name}")
