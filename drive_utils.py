import io
import math
import re
import time
import numpy as np
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


def resolve_folder_path(drive_service, path_parts):
    """
    Traverses a folder path (e.g., ['CityTeam Impact Data', 'CityTeam KPIs', 'FY27', 'Program Goals'])
    in Google Drive / Shared Drives and returns the target folder ID.
    """
    parent_id = None
    for part in path_parts:
        safe_name = part.replace("'", "\\'")
        q = f"name = '{safe_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if parent_id:
            q += f" and '{parent_id}' in parents"
        res = drive_service.files().list(
            q=q,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora='allDrives'
        ).execute()
        files = res.get("files", [])
        if not files:
            return resolve_folder_id(drive_service, path_parts[-1], path_parts[-1])
        parent_id = files[0]["id"]
    return parent_id


def find_file_id(drive_service, filename, folder_id=None, mime_type=None):
    """
    Searches for a file by name and optional parent/MIME type across Shared Drives.
    Supports fallback matching without extensions.
    """
    query = "trashed = false"
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


def parse_numeric_cell(val, default=np.nan):
    """Safely parses numeric values from sheet cells."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return default
    s = str(val).strip().replace(",", "").replace("$", "").replace("%", "")
    if s == "" or s.lower() in {"nan", "none", "tbd", "n/a"}:
        return default
    try:
        num = float(s)
        return int(num) if num.is_integer() else num
    except ValueError:
        return default


def get_grid_metric(rows, row_label_patterns, col_header_patterns, default=np.nan, max_header_rows=25):
    """
    Finds a numeric value in a 2D sheet grid (`ws.get_all_values()`) by matching
    a row title and a column header (case-insensitive regex).
    Immune to inserted rows or shifted columns.
    """
    if isinstance(row_label_patterns, str):
        row_label_patterns = [row_label_patterns]
    if isinstance(col_header_patterns, str):
        col_header_patterns = [col_header_patterns]

    target_col = None
    header_row_idx = None
    for r_idx, row in enumerate(rows[:max_header_rows]):
        for c_idx, cell in enumerate(row):
            cell_clean = str(cell).strip()
            if cell_clean and any(re.search(pat, cell_clean, re.IGNORECASE) for pat in col_header_patterns):
                target_col = c_idx
                header_row_idx = r_idx
                break
        if target_col is not None:
            break

    if target_col is None:
        return default

    for r_idx in range(header_row_idx + 1, len(rows)):
        row = rows[r_idx]
        label_cells = [str(row[c]).strip() for c in range(min(target_col, len(row)))]
        if any(
            re.search(pat, cell, re.IGNORECASE)
            for cell in label_cells
            for pat in row_label_patterns
        ):
            raw_val = row[target_col] if target_col < len(row) else ""
            return parse_numeric_cell(raw_val, default=default)

    return default


def get_cumulative_quarter_metric(rows, row_label_patterns, active_quarter, fy_prefix="FY27", default=0, max_header_rows=25):
    """
    Sums a metric across quarters Q1..active_quarter (e.g., Q1 -> [Q1], Q2 -> [Q1, Q2], etc.)
    by matching row titles and 'FY27 Qx Goal' column headers dynamically.
    """
    quarter_order = ["Q1", "Q2", "Q3", "Q4"]
    q_clean = str(active_quarter).strip().upper()
    if q_clean not in quarter_order:
        q_clean = "Q1"

    quarters_to_sum = quarter_order[:quarter_order.index(q_clean) + 1]
    total = 0.0
    found_any = False

    for q in quarters_to_sum:
        col_pat = rf"^{re.escape(fy_prefix)}\s*{q}\s*Goal$"
        val = get_grid_metric(
            rows,
            row_label_patterns=row_label_patterns,
            col_header_patterns=col_pat,
            default=np.nan,
            max_header_rows=max_header_rows,
        )
        if pd.notna(val):
            total += float(val)
            found_any = True

    if not found_any:
        return default

    return int(total) if float(total).is_integer() else round(total, 2)


def load_program_goal_metrics(
    drive_service,
    gc,
    program_goal_files,
    active_quarter,
    metric_specs,
    fy_prefix="FY27",
    folder_path=None,
):
    """
    Generic loader that reads city FY27 KPI Goal spreadsheets once per (file_name, tab_name)
    and extracts requested metrics per program.

    `metric_specs` maps an output key -> dict with:
      - "patterns": regex string or list of regex strings for row title matching
      - "mode": "point" (active quarter only) or "cumulative" (Q1..active_quarter sum)
      - "default": fallback value if missing
    """
    if folder_path is None:
        folder_path = ["CityTeam Impact Data", "CityTeam KPIs", "FY27", "Program Goals"]

    goals_folder_id = resolve_folder_path(drive_service, folder_path)
    col_header_pattern = rf"^{re.escape(fy_prefix)}\s*{re.escape(active_quarter)}\s*Goal$"

    tab_cache = {}
    opened_workbooks = {}
    results = {key: {} for key in metric_specs}

    for program, (file_name, tab_name) in program_goal_files.items():
        cache_key = (file_name, tab_name)
        if cache_key not in tab_cache:
            if file_name not in opened_workbooks:
                file_info = find_file_id(drive_service, file_name, goals_folder_id)
                if not file_info:
                    raise FileNotFoundError(f"Could not find goal file '{file_name}' in Program Goals folder.")
                opened_workbooks[file_name] = gc.open_by_key(file_info["id"])

            ss = opened_workbooks[file_name]
            try:
                ws = ss.worksheet(tab_name)
            except gspread.WorksheetNotFound:
                ws = ss.get_worksheet(0)

            rows = ws.get_all_values()
            parsed_tab = {}

            for key, spec in metric_specs.items():
                mode = spec.get("mode", "point")
                default_val = spec.get("default", 0)
                patterns = spec["patterns"]

                if mode == "cumulative":
                    parsed_tab[key] = get_cumulative_quarter_metric(
                        rows,
                        row_label_patterns=patterns,
                        active_quarter=active_quarter,
                        fy_prefix=fy_prefix,
                        default=default_val,
                    )
                else:
                    parsed_tab[key] = get_grid_metric(
                        rows,
                        row_label_patterns=patterns,
                        col_header_patterns=col_header_pattern,
                        default=default_val,
                    )

            tab_cache[cache_key] = parsed_tab

        for key in metric_specs:
            results[key][program] = tab_cache[cache_key][key]

    return results


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

    def to_safe(val):
        if val is None or val is pd.NaT:
            return ""
        if isinstance(val, float) and math.isnan(val):
            return ""
        if hasattr(val, 'isoformat'):
            return val.isoformat()
        if hasattr(val, 'item'):
            return val.item()
        return val

    if include_column_header:
        data = [df_write.columns.tolist()] + [[to_safe(v) for v in row] for row in df_write.values.tolist()]
    else:
        data = [[to_safe(v) for v in row] for row in df_write.values.tolist()]
    ws.update(data, value_input_option='RAW')
    print(f"   ✅ '{tab_name}' written ({len(dataframe)} rows)")
    time.sleep(1)