from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "credentials1.json"
FOLDER_ID = "1QsD3gOOaAAyhsjFnQG2Pp362XgJRu1tJ"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
_creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )

# public

drive_service = build("drive", "v3", credentials=_creds)
sheets_service = build("sheets", "v4", credentials=_creds)

def get_sheets(all=False):
    results = []
    initial_result = (
        drive_service.files()
        .list(q=f"'{FOLDER_ID}' in parents", fields="nextPageToken, files(id, name)")
        .execute()
    )
    results.append(initial_result["files"])
    token = initial_result.get("nextPageToken")

    if all:
        while token:
            result = (
                drive_service.files()
                .list(
                    q=f"'{FOLDER_ID}' in parents",
                    fields="nextPageToken, files(id, name)",
                    pageToken=token,
                )
                .execute()
            )
            results.extend(result["files"])
            token = result.get("nextPageToken")
    return results[0]