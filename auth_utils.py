import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
import gspread

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
]

def get_services():
    """Initializes and returns Google Drive and GSpread clients."""
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    SERVICE_ACCOUNT_FILENAME = os.getenv('SERVICE_ACCOUNT_FILENAME', 'ct-kpi-automation-d56fab25ab61.json')
    SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, SERVICE_ACCOUNT_FILENAME)
    
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        # Fallback to default credentials if running in GCP environment without a key file
        try:
            import google.auth
            creds, _ = google.auth.default(scopes=SCOPES)
            # Log identity so you can verify it has Drive access
            print(f"Running in Cloud: Using service identity {getattr(creds, 'service_account_email', 'Default Service Account')}")
        except Exception:
            raise FileNotFoundError(f"Service account file not found at {SERVICE_ACCOUNT_FILE}")
    else:
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        print(f"Running Locally: Using service account {creds.service_account_email}")

    drive_service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    gc = gspread.authorize(creds)
    return drive_service, gc, creds