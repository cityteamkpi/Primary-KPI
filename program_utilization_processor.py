# =========================================================================
# This script processes program utilization.
#
# 1. Reads "Clients in All Programs.xlsx.xlsx"
# 2. Creates "Clients in All Programs - Processed" Google Sheet 
#       - with Renew, Men's TP and Women's TP tabs
# =========================================================================


import warnings
import pandas as pd
import gspread
from zoneinfo import ZoneInfo
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime
from auth_utils import get_services
from drive_utils import download_drive_file, resolve_folder_id, find_file_id, load_raw, write_tab


# Suppress openpyxl warnings regarding styles and validation metadata
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_utilization_processing(

    input_report_file       = 'Clients in All Programs.xlsx',
    input_utilization_file  = 'Clients in All Programs - Processed',
    input_folder_name       = None,
    output_folder_name      = None
):

    # =========================================================================
    # Configuration CONSTANTS
    # =========================================================================
    RAW_DATA_HEADER_ROW = 1
    RAW_DATA_START_COL = 1
    COL_PROGRAM_HEADER = "Program Enrolling_2091"


    # =========================================================================
    # 1. Initialize Google Drive API Service using auth_utils 
    # (Google Drive access authorization) and find specified folders
    # =========================================================================
    drive_service, gc, _ = get_services()

    # Find (resolve) folder IDs
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download to memory the unprocessed input xlsx file from Google Drive
    # =========================================================================
    print(f"Downloading {input_report_file}...")
    report_stream, _, _ = download_drive_file(drive_service, input_report_file, input_folder_id)

    df_raw = load_raw(report_stream, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # print("DEBUG: df_raw loaded:\n", df_raw)

    # =========================================================================
    # 3.Find read/modify/write "Processed" file. It is a Google Sheet
    # =========================================================================
    util_file = find_file_id(drive_service, input_utilization_file, output_folder_id, "application/vnd.google-apps.spreadsheet")

    if not util_file:
        raise FileNotFoundError(f"Google Sheet not found: {input_utilization_file}")

    utilization_file_id = util_file['id']


    #==========================================================================
    # 4. Calculate client enrollment counts from the report input file
    # Identify program column: prioritize 'Program Enrolling', then index 3 (Column D)
    #==========================================================================
    if COL_PROGRAM_HEADER in df_raw.columns:
        program_col = COL_PROGRAM_HEADER
    elif any(col.startswith('Program Enrolling') for col in df_raw.columns):
        program_col = [col for col in df_raw.columns if col.startswith('Program Enrolling')][0]
    elif 'Program Enrolling' in df_raw.columns:
        program_col = 'Program Enrolling'
    elif len(df_raw.columns) >= 4:
        program_col = df_raw.columns[3]
    else:
        program_col = df_raw.columns[0]

    spreadsheet = gc.open_by_key(utilization_file_id)
    tabs_to_update = ["Renew", "Men Turning Point", "Women Turning Point"]

    # Identify relevant programs from the tabs to filter the debug output
    tracked_programs = set()
    for tab_name in tabs_to_update:
        try:
            header_row = spreadsheet.worksheet(tab_name).row_values(1)
            # Programs start at column B (index 1), skipping the Date column at index 0
            tracked_programs.update([str(h).strip() for h in header_row[1:] if h and not str(h).startswith('Unnamed')])
        except gspread.exceptions.WorksheetNotFound:
            continue

    program_counts = df_raw[program_col].astype(str).str.strip().value_counts()
    # Filter to only show programs that are currently listed in the tracking tabs
    filtered_counts = program_counts[program_counts.index.isin(tracked_programs)]
    # print(f"DEBUG: Program enrollment counts (filtered for tracking tabs):\n{filtered_counts}")


    #==========================================================================
    # 5. Process each tab in the utilization tracking sheet
    #==========================================================================
    today_date = datetime.now(ZoneInfo("America/Los_Angeles")).strftime('%-m/%-d/%y')

    for tab_name in tabs_to_update:
        try:
            sheet = spreadsheet.worksheet(tab_name)
            print(f"Updating tab: {tab_name}")
        except gspread.exceptions.WorksheetNotFound:
            print(f"Warning: Tab '{tab_name}' not found. Skipping.")
            continue

        df_sheet = get_as_dataframe(sheet)
        
        # Cleanup and set index
        df_sheet = df_sheet.dropna(how='all').dropna(axis=1, how='all')
        if not df_sheet.empty:
            df_sheet = df_sheet.set_index(df_sheet.columns[0])
            # Ignore unnamed columns (likely garbage or empty artifacts)
            df_sheet = df_sheet.loc[:, ~df_sheet.columns.str.match(r'^Unnamed')]
            df_sheet.index.name = 'Date'
        
        df_sheet = df_sheet.copy()

        # Standardize index to match string-based date format
        valid_dates = pd.to_datetime(df_sheet.index, errors='coerce', format='mixed')
        mask = valid_dates.notna()
        df_sheet = df_sheet[mask]
        if not df_sheet.empty:
            df_sheet.index = valid_dates[mask].strftime('%-m/%-d/%y')

        # Update today's row for every program column in this tab
        for program in df_sheet.columns:
            clean_prog = str(program).strip()
            count = int(program_counts.get(clean_prog, 0))
            df_sheet.loc[today_date, program] = count

        # Write back to Google Sheet using standardized helper
        write_tab(spreadsheet, tab_name, df_sheet.reset_index())

    print(f"Successfully updated Google Sheet: {input_utilization_file}")
    print(f"Link: https://docs.google.com/spreadsheets/d/{utilization_file_id}")

if __name__ == "__main__":
    run_utilization_processing()
