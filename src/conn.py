from googleapiclient.discovery import build
from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = 'credentials1.json'
FOLDER_ID = '1QsD3gOOaAAyhsjFnQG2Pp362XgJRu1tJ'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Authenticate using the service account key file
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)

# Build the Google Drive API client
drive_service = build('drive', 'v3', credentials=credentials)

# returns all files within specified folder
results = drive_service.files().list( q=f"'{FOLDER_ID}' in parents", fields="files(id, name)" ).execute()

for file in results['files']:
    print(file['name'] + "  ||  " + file['id'])

