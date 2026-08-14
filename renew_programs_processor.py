#!/usr/bin/env python3
"""
This script processes the Renew Programs Report
"""

import warnings
import os
import time
import pandas as pd
import numpy as np

import gspread
from gspread_dataframe import set_with_dataframe
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

MONTHLY_HOURS = 173.33

def run_renew_processing(
    input_file="Renew Programs Report.xlsx",
    output_file="Renew Programs Report - Processed",
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
    COL_SOBER        = "Sober?_6585"
    COL_CHECKIN_DATE = "Checkin Date_6583"
    COL_HOURLY_WAGE  = "Hourly Wage_6586"
    COL_HOUSING_EXP  = "Housing Expense_6587"
    COL_HOW_CHECKED  = "How Checked_6584"
    COL_HOUSING_TYPE = "Housing_6588"
    COL_INTERN       = "Intern Program_6619"

    GRAD_REASONS = ["Completer", "Graduation", "Completion of Program"]

    # Programs for Renew analysis
    RENEW_PROGRAMS = ["Chester Renew", "Oakland Renew", "Portland Renew", "San Jose Men Renew", "GV Renew"]

    # =========================================================================
    # 1. Initialize Google Drive API Service using auth_utils 
    # =========================================================================
    drive_service, gc, _ = get_services()

    # Initialize dynamic constants from Google Sheet
    constants.sync_constants()

    # Find (resolve) folder IDs
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and Load Data
    # =========================================================================
    print(f"Downloading {input_file}...")
    fh, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    
    df_raw = load_raw(fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # Clean up dates and filter
    df_base = df_raw.copy()
    df_base[COL_START_DATE] = pd.to_datetime(df_base[COL_START_DATE], errors="coerce")
    df_base[COL_EXIT_DATE]  = pd.to_datetime(df_base[COL_EXIT_DATE],  errors="coerce")

    # Remap and filter to target programs
    df_base[COL_PROGRAM] = df_base[COL_PROGRAM].apply(lambda p: constants.FORWARD_MAP.get(str(p).strip(), p) if not pd.isna(p) else p)
    df_base = df_base[df_base[COL_PROGRAM].isin(RENEW_PROGRAMS)].reset_index(drop=True)

    # =========================================================================
    # 3. Data Processing Logic
    # =========================================================================
    def process_graduates(df_base):
        df = df_base.copy()
        df["_is_grad"] = (df[COL_EXIT_REASON].isin(GRAD_REASONS)).astype(int)
        df = (df.sort_values(["_is_grad", COL_EXIT_DATE], ascending=[False, False])
              .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
              .drop(columns=["_is_grad"]))

        df = df[(df[COL_EXIT_DATE].notna()) & (df[COL_EXIT_REASON].isin(GRAD_REASONS))].copy()
        df["City"] = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"] = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        df["Current FY Goal"] = df[COL_PROGRAM].map(constants.PROGRAM_GOALS)
        df["Current FY Projection"] = df[COL_PROGRAM].map(constants.PROGRAM_PROJECTIONS)
        df["Capacity"] = df[COL_PROGRAM].map(constants.PROGRAM_CAPACITY)
        df["Theoretical Maximum"] = df[COL_PROGRAM].map(constants.PROGRAM_THEORETICAL_MAX)

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

        df["Capacity"] = df[COL_PROGRAM].map(constants.OCCUPANCY_CAPACITY)
        df["Goal"] = df[COL_PROGRAM].map(constants.OCCUPANCY_GOAL)
        df["Prior FY Occupancy"] = df[COL_PROGRAM].map(constants.OCCUPANCY_PRIOR_FY)
        return df.reset_index(drop=True)

    def process_retention(df_base):
        df = df_base.copy()
        df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
        df = (df.sort_values(["_no_exit", COL_EXIT_DATE], ascending=[False, False])
              .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
              .drop(columns=["_no_exit"]))

        cutoff = pd.Timestamp(constants.RETENTION_START_CUTOFF)
        gap = (df[COL_EXIT_DATE] - df[COL_START_DATE]).dt.days
        df["City"] = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"] = df[COL_START_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_START_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"] = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        df["Entered Since Prior FY"] = ((df[COL_START_DATE] >= cutoff) & (df[COL_EXIT_DATE].isna() | (gap > 30))).astype(int)
        df["Exit After 30 Days"] = ((df[COL_START_DATE] >= cutoff) & df[COL_EXIT_DATE].notna() & (gap > 30) & (~df[COL_EXIT_REASON].isin(GRAD_REASONS))).astype(int)
        df["Graduated"] = ((df[COL_START_DATE] >= cutoff) & df[COL_EXIT_REASON].isin(GRAD_REASONS)).astype(int)
        df["Still in Program"] = ((df[COL_START_DATE] >= cutoff) & df[COL_EXIT_DATE].isna()).astype(int)
        return df.reset_index(drop=True)

    # =========================================================================
    # 4. Execute Processing
    # =========================================================================
    print("Processing Graduates...")
    df_grad = process_graduates(df_base)

    print("Processing Occupancy...")
    df_occ = process_occupancy(df_base)

    print("Processing Retention...")
    df_ret = process_retention(df_base)

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
    write_tab(spreadsheet, "Raw Data", df_raw)
    write_tab(spreadsheet, "Renew Graduates", df_grad)
    write_tab(spreadsheet, "Occupancy", df_occ)
    write_tab(spreadsheet, "Retention", df_ret)

  # Delete lingering first tab (Sheet1)
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Done!: {spreadsheet.url}")


if __name__ == "__main__":
    run_renew_processing(
         input_folder_name="Apricot Report Incoming",
         output_folder_name="KPI Processed Data"
    )
