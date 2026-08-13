# ============================================================
# This script processes the Turning Point Report
#
# 1. Reads "Turning Point Report.xlsx"
# 2. Creates "Turning Point Report - Processed" Google Sheet 
#       - with Raw Data, Graduates, Completers, Occupancy, and Housed tabs
# ============================================================

import warnings
import time
import pandas as pd
import numpy as np

import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_turning_point_processing(
    input_file="Turning Point Report.xlsx",
    output_file="Turning Point Report - Processed",
    input_folder_name=None,
    output_folder_name=None
):
    # =========================================================================
    # Spreadsheet-Specific Constants
    # =========================================================================
    RAW_DATA_HEADER_ROW = 3
    RAW_DATA_START_COL  = 1

    COL_RECORD_ID    = "Record Id_102"
    COL_PROGRAM      = "Program Enrolling_2091"
    COL_START_DATE   = "Start Date_2090"
    COL_EXIT_DATE    = "Exit Date_2100"
    COL_EXIT_REASON  = "Primary Reason for Exit_2102"
    COL_HOUSED       = "Successfully Housed (Is Housing Healthy?)_4368"
    COL_INTERN_PROGRAM = "Intern Program_6619"

    PROGRAMS_TO_INCLUDE = [
        "Heritage Home", "Oakland Women Turning Point", "San Jose Men Turning Point",
        "Chester Men Turning Point", "Chester Women Turning Point", "GV Turning Point",
        "Oakland Men Turning Point", "Oakland Youth Collective", "Portland Community of Hope",
        "Portland Youth Collective", "San Jose Youth Collective", "Chester Turning Point",
    ]

    COMPLETER_REASONS = ["Completer", "Completion of Program"]

    # =========================================================================
    # 1. Initialize Services, find folders and Sync Constants
    # =========================================================================
    drive_service, gc, _ = get_services()

    # Initialize dynamic constants from Constants Google Sheet
    constants.sync_constants()

    # Find (resolve) folder IDs
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and load to memory Raw Data from xlsx file
    # =========================================================================
    print(f"Downloading {input_file}...")
    fh, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    df_raw = load_raw(fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # Initial Cleanup
    df_base = df_raw.copy()
    df_base[COL_PROGRAM] = df_base[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
    df_base[COL_START_DATE] = pd.to_datetime(df_base[COL_START_DATE], errors="coerce")
    df_base[COL_EXIT_DATE]  = pd.to_datetime(df_base[COL_EXIT_DATE],  errors="coerce")

    # Filter to target programs (handling interns remapping first)
    prog_lower_map = {p.lower(): p for p in PROGRAMS_TO_INCLUDE}
    def match_program(val):
        if pd.isna(val): return val
        return prog_lower_map.get(str(val).strip().lower(), str(val).strip())

    if COL_INTERN_PROGRAM in df_base.columns:
        is_intern = df_base[COL_PROGRAM] == "Program Graduate Intern"
        df_base.loc[is_intern, COL_PROGRAM] = df_base.loc[is_intern, COL_INTERN_PROGRAM].apply(match_program)

    df_base = df_base[df_base[COL_PROGRAM].isin(PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

    # =========================================================================
    # 3. Data Processing Functions
    # =========================================================================
    def process_exited_category(df_base, reasons_list=None, exact_match=None):
        df = df_base.copy()
        if exact_match:
            df["_priority"] = (df[COL_EXIT_REASON] == exact_match).astype(int)
        elif reasons_list:
            df["_priority"] = (df[COL_EXIT_REASON].isin(reasons_list)).astype(int)
        else:
            df["_priority"] = 0

        df = (df.sort_values(["_priority", COL_EXIT_DATE], ascending=[False, False])
              .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
              .drop(columns=["_priority"]))

        mask = (df[COL_EXIT_DATE].notna())
        if exact_match: mask &= (df[COL_EXIT_REASON] == exact_match)
        elif reasons_list: mask &= (df[COL_EXIT_REASON].isin(reasons_list))
        
        df = df[mask].copy()
        df["City"] = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"] = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= w_start) & (df[COL_EXIT_DATE] <= w_end)).astype(int)
        return df.reset_index(drop=True)

    def was_active_on(start_date, exit_date, ref_date):
        ref = pd.Timestamp(ref_date)
        if pd.isnull(start_date) or start_date > ref: return 0
        if pd.isnull(exit_date) or exit_date > ref: return 1
        return 0

    def process_occupancy(df_base):
        df = df_base.copy()
        df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
        df = (df.sort_values(["_no_exit", COL_START_DATE], ascending=[False, False])
              .drop_duplicates(subset=[COL_RECORD_ID], keep="first")
              .drop(columns=["_no_exit"]))

        df["City"] = df[COL_PROGRAM].apply(constants.assign_city)
        df[constants.OCC_PRIOR_LABEL] = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_PRIOR_DATE), axis=1)
        df[constants.OCC_CURRENT_LABEL] = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_CURRENT_DATE), axis=1)

        df["Capacity"] = df[COL_PROGRAM].map(constants.TP_CAPACITY)
        df["Goal"] = df[COL_PROGRAM].map(constants.TP_GOAL)
        df["Prior Period Actuals"] = df[COL_PROGRAM].map(constants.TP_PRIOR_FY)
        return df.reset_index(drop=True)

    def process_housed(df_base):
        # Specifically for Housed tab, we keep interns distinct if they exist
        df = df_base.copy()
        df["_is_housed"] = df[COL_HOUSED].astype(str).str.contains("Yes", case=False, na=False).astype(int) if COL_HOUSED in df.columns else 0
        df["_is_exited_housed"] = (df[COL_EXIT_DATE].notna() & (df["_is_housed"] == 1)).astype(int)

        df = (df.sort_values(["_is_exited_housed", "_is_housed", COL_EXIT_DATE, COL_START_DATE], ascending=[False, False, False, False])
              .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM], keep="first")
              .drop(columns=["_is_housed", "_is_exited_housed"]))

        df["City"] = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"] = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        
        df["Successfully Housed?"] = df[COL_HOUSED].astype(str).str.contains("Yes", case=False, na=False).astype(int) if COL_HOUSED in df.columns else 0

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            # Count actuals only if Successfully Housed? = 1
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= w_start) & (df[COL_EXIT_DATE] <= w_end) & (df["Successfully Housed?"] == 1)).astype(int)
        
        df["Capacity"] = df[COL_PROGRAM].map(constants.TP_CAPACITY)
        return df.reset_index(drop=True)

    # =========================================================================
    # 4. Execute Processing
    # =========================================================================
    print("Processing Graduates...")
    df_grad = process_exited_category(df_base, exact_match="Graduation")

    print("Processing Completers...")
    df_comp = process_exited_category(df_base, reasons_list=COMPLETER_REASONS)

    print("Processing Occupancy...")
    df_occ = process_occupancy(df_base)

    print("Processing Housed...")
    df_housed = process_housed(df_base)

    # =========================================================================
    # 4. Handle Output Spreadsheet
    # =========================================================================
    processed_file = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")
    if processed_file:
        spreadsheet = gc.open_by_key(processed_file['id'])
        print(f"📄 Opened existing: '{output_file}'")
    else:
        file_metadata = {
            'name': output_file,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [output_folder_id]
        }
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        spreadsheet = gc.open_by_key(new_sheet['id'])
        print(f"📄 Created new: '{output_file}'")



    # =========================================================================
    # 5. Export Tabs
    # =========================================================================
    write_tab(spreadsheet, "Raw Data", df_raw)
    write_tab(spreadsheet, "Turning Point Graduates", df_grad)
    write_tab(spreadsheet, "Turning Point Completers", df_comp)
    write_tab(spreadsheet, "Turning Point Occupancy", df_occ)
    write_tab(spreadsheet, "Successfully Housed", df_housed)

    # Delete lingering first tab (Sheet1)
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except: pass

    print(f"\n🎉 Done!: {spreadsheet.url}")


if __name__ == "__main__":
    run_turning_point_processing(
         input_folder_name="Apricot Report Incoming",
         output_folder_name="KPI Processed Data"
    )
