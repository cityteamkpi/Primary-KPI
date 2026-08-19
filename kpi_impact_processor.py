#!/usr/bin/env python3
"""
This script processes Mentorship, Family Reunifications & Baptisms
"""

import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import find_file_id, resolve_folder_id, download_drive_file, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_kpi_impact_processing(
    baptisms_input_file       = 'CityTeam Baptisms.xlsx',
    mentorship_input_file     = 'Mentorship.xlsx',
    reunifications_input_file = 'Family Reunifications.xlsx',
    output_file               = 'Mentorship Reunifications Baptisms - Processed',
    input_folder_name         = None,
    output_folder_name        = None
):
    # =========================================================================
    # Spreadsheet-Specific Constants
    # =========================================================================
    RAW_DATA_HEADER_ROW = 3
    RAW_DATA_START_COL  = 1

    COL_RECORD_ID  = "Record Id_102"
    COL_PROGRAM    = "Program Enrolling"
    COL_INTERN     = "Intern Program_6619"
    COL_MENTORSHIP = "Acquired Spiritual Mentor_5136"

    DATE_COLS = {
        "Mentorship"           : "Start Date for Mentoring_4365",
        "Family Reunifications": "Date of Reunification_5140",
        "CityTeam Baptisms"    : "Baptism Date_4378",
    }

    # =========================================================================
    # 1. Initialize Services
    # =========================================================================
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download Raw Data
    # =========================================================================
    print("Downloading source reports from Google Drive...")
    m_fh, _, _ = download_drive_file(drive_service, mentorship_input_file, input_folder_id)
    r_fh, _, _ = download_drive_file(drive_service, reunifications_input_file, input_folder_id)
    b_fh, _, _ = download_drive_file(drive_service, baptisms_input_file, input_folder_id)

    df_raw_mentorship = load_raw(m_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)
    df_raw_reunif     = load_raw(r_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)
    df_raw_baptisms   = load_raw(b_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # =========================================================================
    # 3. Processing Function
    # =========================================================================
    def process_spiritual(label, df_raw):
        date_col = DATE_COLS[label]
        df = df_raw.copy()

        print(f"\n--- {label} ---")
        print(f"   Raw rows: {len(df)}")

        if date_col not in df.columns:
            print(f"   Warning: '{date_col}' not found. Skipping.")
            return pd.DataFrame()

        # Parse date
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        print(f"   Non-null dates after parse: {df[date_col].notna().sum()}")

        # Filter to wide FY range — 4 years back to current FY end
        _wide_start = pd.Timestamp(f"{constants.fy_num - 4}-09-01")
        _wide_end   = pd.Timestamp(f"{constants.fy_num}-08-31")
        df = df[(df[date_col] >= _wide_start) & (df[date_col] <= _wide_end)].copy().reset_index(drop=True)
        print(f"   After FY range filter: {len(df)} rows")

        if df.empty:
            print(f"   Warning: No data remaining after date filter.")
            return pd.DataFrame()

        # Mentorship: filter to Acquired Spiritual Mentor = Yes
        if label == "Mentorship" and COL_MENTORSHIP in df.columns:
            before = len(df)
            df = df[df[COL_MENTORSHIP].astype(str).str.strip() == "Yes"].reset_index(drop=True)
            print(f"   After Mentor=Yes filter: {len(df)} rows (removed {before - len(df)})")

        # Intern priority flag
        df["_intern_priority"] = 0
        if COL_INTERN in df.columns:
            is_intern = df[COL_PROGRAM] == "Program Graduate Intern"
            df.loc[is_intern & df[COL_INTERN].notna(), "_intern_priority"] = 1
            df.loc[is_intern, COL_PROGRAM] = df.loc[is_intern, COL_INTERN]
            before = len(df)
            df = df[df[COL_PROGRAM] != "Program Graduate Intern"].reset_index(drop=True)
            print(f"   Remapped/excluded PGI: {before - len(df)} rows removed")

        # Remap Forward programs → Renew programs
        df[COL_PROGRAM] = df[COL_PROGRAM].apply(
            lambda p: constants.FORWARD_MAP.get(str(p).strip(), p) if not pd.isna(p) else p
        )

        # Assign City
        df["City"]    = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]    = df[date_col].apply(constants.get_fiscal_year)
        df["Quarter"] = df[date_col].apply(constants.get_fiscal_quarter)
        df["Year Q"]  = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        # Calculate Actuals BEFORE dedup so prior records are preserved
        actuals_cols = list(constants.ACTUALS_WINDOWS.keys())
        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[date_col].notna() & (df[date_col] >= w_start) & (df[date_col] <= w_end)).astype(int)

        # Dedup: aggregate actuals first, then dedup
        if COL_RECORD_ID in df.columns:
            before = len(df)
            id_actuals = df.groupby(COL_RECORD_ID)[actuals_cols].max()
            ascending_date = (label == "CityTeam Baptisms")
            direction = "earliest" if ascending_date else "latest"
            df = (
                df.sort_values(["_intern_priority", date_col], ascending=[False, ascending_date])
                  .drop_duplicates(subset=[COL_RECORD_ID], keep="first")
                  .drop(columns=["_intern_priority"])
            )
            df = df.drop(columns=actuals_cols).merge(id_actuals, on=COL_RECORD_ID, how="left")
            df = df.sort_values(date_col).reset_index(drop=True)
            print(f"   Dedup ({direction} date): {before} -> {len(df)} rows ({before - len(df)} removed)")
        else:
            df = df.drop(columns=["_intern_priority"])

        # Ensure Actuals are int
        for k in actuals_cols:
            if k in df.columns:
                df[k] = df[k].fillna(0).astype(int)

        # Explicit column selection
        BASE_COLS = [COL_RECORD_ID, COL_PROGRAM, "City", "Year", "Quarter", "Year Q"]
        df = df[[c for c in BASE_COLS + actuals_cols if c in df.columns]]

        return df

    # =========================================================================
    # 4. Execute Processing
    # =========================================================================
    df_mentorship = process_spiritual("Mentorship", df_raw_mentorship)
    df_reunif     = process_spiritual("Family Reunifications", df_raw_reunif)
    df_baptisms   = process_spiritual("CityTeam Baptisms", df_raw_baptisms)

    # =========================================================================
    # 5. Handle Output Spreadsheet
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
    # 6. Export Tabs
    # =========================================================================
    write_tab(spreadsheet, "Mentorship Raw Data",      df_raw_mentorship)
    write_tab(spreadsheet, "Mentorship Processed",     df_mentorship)
    write_tab(spreadsheet, "Reunifications Raw Data",  df_raw_reunif)
    write_tab(spreadsheet, "Reunifications Processed", df_reunif)
    write_tab(spreadsheet, "Baptisms Raw Data",        df_raw_baptisms)
    write_tab(spreadsheet, "Baptisms Processed",       df_baptisms)

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Done! {spreadsheet.url}")


if __name__ == "__main__":
    run_kpi_impact_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )