import io
import time
import pandas as pd
import gspread
from googleapiclient.http import MediaIoBaseDownload
from gspread_dataframe import set_with_dataframe

def find_folder_id(drive_service, folder_name):
    """Finds a folder ID by name."""
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = drive_service.files().list(
        q=query,
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora='allDrives'
    ).execute()
    items = results.get('files', [])
    if not items:
        return None
    return items[0]['id']

def resolve_folder_id(drive_service, folder_name, folder_label="Folder"):
    """
    Resolves a folder name to an ID. 
    Raises ValueError if a name is provided but no folder is found.
    """
    if not folder_name:
        return None
        
    folder_id = find_folder_id(drive_service, folder_name)
    if not folder_id:
        raise ValueError(f"No {folder_label.lower()} folder found with name: {folder_name}")
    return folder_id

def find_file_id(drive_service, filename, folder_id=None, mime_type=None):
    """
    Searches for a file by name and optional parent/MIME type.
    Returns the file ID or None if not found.
    """
    query = f"name = '{filename}' and trashed = false"
    if folder_id:
        query += f" and '{folder_id}' in parents"
    if mime_type:
        query += f" and mimeType = '{mime_type}'"

    results = drive_service.files().list(
        q=query,
        fields="files(id, parents, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora='allDrives'
    ).execute()

    items = results.get('files', [])
    if not items:
        return None
    
    if len(items) > 1:
        print(f"Warning: Found {len(items)} files with name '{filename}'. Using the first one.")
    
    return items[0]

def download_drive_file(drive_service, filename, folder_id=None):
    """
    Searches for a file and downloads it into a BytesIO stream.
    Handles both binary files and Google Sheets (via CSV export).
    Returns the stream and the parent folder ID.
    """
    target = find_file_id(drive_service, filename, folder_id)
    if not target:
        raise FileNotFoundError(f"No file found with name: {filename}")

    file_id = target['id']
    mime_type = target.get('mimeType', '')
    parents = target.get('parents', [])
    
    if not parents:
        raise ValueError(f"Could not determine parent folder for {filename}")
    parent_id = parents[0]

    # Handle Google Sheets vs regular files
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        request = drive_service.files().get_media(fileId=file_id)

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.seek(0)
    return fh, parent_id, file_id

def load_raw(fh, header_row=3, start_col=1):
    """
    Standardized Excel ingestion logic.
    Loads xlsx from defined header row, strips empty rows/cols and 'Unnamed' columns.
    """
    df = pd.read_excel(fh, header=header_row, engine="openpyxl")
    df = df.iloc[:, start_col:]
    df = df.dropna(how="all").reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    return df

def write_tab(ss, tab_name, dataframe, include_index=False, include_column_header=True):
    """
    Standardized helper to write a pandas DataFrame to a Google Sheet tab.
    Clears the tab if it exists, creates it if it doesn't.
    """
    try:
        ws = ss.worksheet(tab_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        # Default sizing: rows based on data + buffer, columns based on data + buffer
        rows = max(100, len(dataframe) + 100)
        cols = max(20, len(dataframe.columns) + 5)
        ws = ss.add_worksheet(title=tab_name, rows=rows, cols=cols)
    
    set_with_dataframe(ws, dataframe, include_index=include_index, include_column_header=include_column_header)
    print(f"   ✅ '{tab_name}' written ({len(dataframe)} rows)")
    time.sleep(1) # Respect API rate limits