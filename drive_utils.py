import io
import time
import pandas as pd
import gspread
from googleapiclient.http import MediaIoBaseDownload
from gspread_dataframe import set_with_dataframe

def find_folder_id(drive_service, folder_name):
    """Finds a folder ID by name across My Drive and Shared Drives."""
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
    Searches for a file by name and optional parent/MIME type across Shared Drives.
    Supports fallback matching without extensions.
    """
    query = f"trashed = false"
    if folder_id:
        query += f" and '{folder_id}' in parents"
    
    exact_query = f"{query} and name = '{filename}'"
    if mime_type:
        exact_query += f" and mimeType = '{mime_type}'"

    results = drive_service.files().list(
        q=exact_query,
        fields="files(id, parents, mimeType, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora='allDrives'
    ).execute()

    items = results.get('files', [])
    if items:
        return items[0]

    # Fallback search without extension
    base_name = filename.rsplit('.', 1)[0] if '.' in filename else filename
    fallback_query = f"{query} and (name = '{base_name}' or name contains '{base_name}')"

    results = drive_service.files().list(
        q=fallback_query,
        fields="files(id, parents, mimeType, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora='allDrives'
    ).execute()

    items = results.get('files', [])
    if not items:
        return None
    if len(items) > 1:
        print(f"Warning: Found {len(items)} files matching '{filename}'. Using the first one.")
    return items[0]

def download_drive_file(drive_service, filename, folder_id=None):
    """
    Searches for a file and downloads it into a BytesIO stream.
    Handles both binary files and Google Sheets (via XLSX export).
    Includes supportsAllDrives=True for Shared Drive compatibility.
    """
    target = find_file_id(drive_service, filename, folder_id)
    if not target:
        raise FileNotFoundError(f"No file found with name: {filename}")

    file_id   = target['id']
    mime_type = target.get('mimeType', '')
    parents   = target.get('parents', [])

    parent_id = parents[0] if parents else folder_id

    # Add supportsAllDrives=True to request objects to allow Shared Drive downloads
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        request = drive_service.files().get_media(
            fileId=file_id,
            supportsAllDrives=True
        )

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    max_retries = 5

    while not done:
        for attempt in range(max_retries):
            try:
                status, done = downloader.next_chunk()
                if status:
                    print(f"   Download Progress: {int(status.progress() * 100)}%", end="\r")
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    print(f"\n❌ Failed to download chunk for {filename} after {max_retries} attempts.")
                    raise e
                wait_time = 2 ** attempt
                print(f"\n⚠️ Download timeout/error: {e}. Retrying in {wait_time}s...")
                time.sleep(wait_time)

    fh.seek(0)
    return fh, parent_id, file_id

def load_raw(fh, header_row=3, start_col=1):
    """
    Standardized Excel ingestion logic.
    Loads xlsx from defined header row (default row 4), strips empty rows/cols and 'Unnamed' columns.
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
        rows = max(100, len(dataframe) + 100)
        cols = max(20, len(dataframe.columns) + 5)
        ws = ss.add_worksheet(title=tab_name, rows=rows, cols=cols)

    df_write = dataframe.copy()
    for col in df_write.columns:
        if pd.api.types.is_datetime64_any_dtype(df_write[col]):
            df_write[col] = df_write[col].dt.strftime("%Y-%m-%d").fillna("")
        elif pd.api.types.is_integer_dtype(df_write[col]):
            df_write[col] = df_write[col].fillna(0).astype(int)
        elif pd.api.types.is_float_dtype(df_write[col]):
            df_write[col] = df_write[col].fillna(0)
        else:
            df_write[col] = df_write[col].fillna("")

    import math
    def to_safe(val):
        if val is None or val is pd.NaT: return ""
        if isinstance(val, float) and math.isnan(val): return ""
        if hasattr(val, 'isoformat'): return val.isoformat()
        if hasattr(val, 'item'): return val.item()
        return val

    if include_column_header:
        data = [df_write.columns.tolist()] + [[to_safe(v) for v in row] for row in df_write.values.tolist()]
    else:
        data = [[to_safe(v) for v in row] for row in df_write.values.tolist()]
    ws.update(data, value_input_option='RAW')
    print(f"   ✅ '{tab_name}' written ({len(dataframe)} rows)")
    time.sleep(1)