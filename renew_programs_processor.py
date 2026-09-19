#!/usr/bin/env python3
"""
Renew Programs Processing Pipeline
Processes Graduates, Occupancy, and Retention tabs for Q4 FY26.
Retention cohort covers clients starting between 5/1/2026 (1 month prior) and 8/31/2026.
"""

import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_renew_processing(
    input_file="Renew Programs Report.xlsx",
    output_file="Renew Programs Report - Processed",
    input_folder_name=None,
    output_folder_name=None
):
    # =========================================================================
    # Header & Column Constants
    # =========================================================================
    RAW_DATA_HEADER_ROW = 3
    RAW_DATA_START_COL  = 1

    COL_RECORD_ID    = "Record Id_102"
    COL_PROGRAM      = "Program Enrolling_2091"
    COL_START_DATE   = "Start Date_2090"
    COL_EXIT_DATE    = "Exit Date_2100"
    COL_EXIT_REASON  = "Primary Reason for Exit_2102"
    COL_CHILD        = "Name of Child_6313"

    GRAD_REASONS   = constants.GRAD_REASONS
    RENEW_PROGRAMS = list(constants.RENEW_PROGRAMS)

    # Static Hardcoded Values Mapped by Program Enrolling
    FY27_CAPACITY_MAP = {
        "Chester Renew": 26,
        "GV Renew": 38,
        "Oakland Renew": 48,
        "Portland Renew": 40,
        "San Jose Men Renew": 65,
    }

    FY27_OCC_GOAL_MAP = {
        "Chester Renew": 23,
        "GV Renew": 34,
        "Oakland Renew": 43,
        "Portland Renew": 36,
        "San Jose Men Renew": 58,
    }

    # =========================================================================
    # 1. Initialize Services & Sync Constants
    # =========================================================================
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    # Explicitly enforce pd.Timestamp typing for quarter boundaries
    q_start_ts   = pd.Timestamp(constants.CURRENT_Q_START)      # 2026-06-01
    q_end_ts     = pd.Timestamp(constants.CURRENT_Q_END)        # 2026-08-31
    cohort_start = q_start_ts - pd.DateOffset(months=1)         # 2026-05-01 (1 month prior)
    far_future   = pd.Timestamp("2099-12-31")

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and Load Raw Data
    # =========================================================================
    print(f"Downloading {input_file}...")
    fh, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    df_raw = load_raw(fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    df_base = df_raw.copy()
    df_base[COL_START_DATE] = pd.to_datetime(df_base[COL_START_DATE], errors="coerce")
    df_base[COL_EXIT_DATE]  = pd.to_datetime(df_base[COL_EXIT_DATE],  errors="coerce")
    df_base[COL_PROGRAM] = df_base[COL_PROGRAM].replace("House of Grace", "GV Renew")
    df_base = df_base[df_base[COL_PROGRAM].isin(RENEW_PROGRAMS)].reset_index(drop=True)

    # =========================================================================
    # 3. Helper Functions
    # =========================================================================
    def categorize_exit_reason(val):
        if pd.isna(val) or str(val).strip() == "": return None
        v = str(val).lower().strip()
        if any(x in v for x in ['graduation', 'completer', 'completion', 'internship ended']): return "Graduation"
        if any(x in v for x in ['relapse', 'drug', 'detox', 'fentanyl', 'paraph', 'blackout', 'deviating']): return "Relapse"
        if any(x in v for x in ['discipline', 'conduct', 'behavior', 'threat', 'violence', 'awol', '30 day review', 'incident']): return "Discipline"
        if any(x in v for x in ['family', 'child', 'wife', 'daughter', 'son', 'mother', 'reunif']): return "Family"
        if any(x in v for x in ['medical', 'mental', 'health', 'hospital', 'cognitive', 'ssi', 'out of scope']): return "Medical/Mental Health"
        if any(x in v for x in ['legal', 'parole', 'probation', 'warrant', 'criminal', 'remanded', 'county', 'immigr']): return "Legal"
        if any(x in v for x in ['transfer', 'another program', 'different program', 'forward', 'other program']): return "Transfer"
        if any(x in v for x in ['housing', 'sle', 'perm housing']): return "Housing"
        if any(x in v for x in ['job', 'work', 'employ']): return "Job"
        if any(x in v for x in ['relationship']): return "Relationship"
        if any(x in v for x in ['not eligible', 'duplicate', 'accident', 'test client', 'mistake', 'lied', 'restart', 'not accepted']): return "Not Eligible/Admin"
        if any(x in v for x in ['self', 'not ready', 'voluntary', 'chose', 'decided', 'walked', 'left', "wasn't ready", 'refused']): return "Self Exit"
        return "Unknown/Other"

    def fmt_date(df, col):
        return pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")

    # =========================================================================
    # 4. Tab Processors
    # =========================================================================
    def process_graduates(df_base):
        df = df_base.copy()
        df["_is_grad"] = df[COL_EXIT_REASON].isin(GRAD_REASONS).astype(int)
        df = (df.sort_values(["_is_grad", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
                .drop(columns=["_is_grad"]))
        df = df[df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(GRAD_REASONS)].copy()

        df["City"]     = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]     = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]  = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]   = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        
        df["Current FY Goal"]        = df[COL_PROGRAM].map(constants.PROGRAM_GOALS)
        df["Next FY Goal"]           = df[COL_PROGRAM].map(constants.NEXT_FY_GRAD_GOALS)
        df["Current FY Projection"]  = df[COL_PROGRAM].map(constants.PROGRAM_PROJECTIONS)
        df["Capacity"]               = df[COL_PROGRAM].map(constants.PROGRAM_CAPACITY)
        df["Theoretical Maximum"]    = df[COL_PROGRAM].map(constants.PROGRAM_THEORETICAL_MAX)

        df["FY27 Capacity"]          = df[COL_PROGRAM].map(FY27_CAPACITY_MAP)
        df["FY27 Occupancy Goal"]    = df[COL_PROGRAM].map(FY27_OCC_GOAL_MAP)

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= pd.Timestamp(w_start)) & (df[COL_EXIT_DATE] <= pd.Timestamp(w_end))).astype(int)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q",
                "Current FY Goal", "Next FY Goal", "Current FY Projection", "Capacity", "Theoretical Maximum",
                "FY27 Capacity", "FY27 Occupancy Goal"
                ] + list(constants.ACTUALS_WINDOWS.keys())
        df = df[[c for c in cols if c in df.columns]].copy()

        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
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
        df[constants.OCC_PRIOR_LABEL]   = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_PRIOR_DATE), axis=1)
        df[constants.OCC_CURRENT_LABEL] = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_CURRENT_DATE), axis=1)

        def count_kids_only(row):
            if row[COL_PROGRAM] != "GV Renew": return 0
            val = row.get(COL_CHILD)
            if pd.isna(val) or str(val).strip() == "": return 0
            import re as _re
            parts = _re.split(r'[,&]|\band\b', str(val).strip(), flags=_re.IGNORECASE)
            return len([p for p in parts if p.strip()])
            
        df["Number of Children"] = df.apply(count_kids_only, axis=1)
        df["Capacity"]           = df[COL_PROGRAM].map(constants.OCCUPANCY_CAPACITY)
        df["Goal"]               = df[COL_PROGRAM].map(constants.OCCUPANCY_GOAL)
        df["Next FY Goal"]       = df[COL_PROGRAM].map(constants.NEXT_FY_OCC_GOALS)
        df["Prior FY Occupancy"] = df[COL_PROGRAM].map(constants.OCCUPANCY_PRIOR_FY)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE,
                "City", constants.OCC_PRIOR_LABEL, constants.OCC_CURRENT_LABEL,
                COL_CHILD, "Number of Children", "Capacity", "Goal", "Next FY Goal", "Prior FY Occupancy"]
        df = df[[c for c in cols if c in df.columns]].copy()

        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        return df.reset_index(drop=True)

    def process_retention(df_base):
        df = df_base.copy()

        df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
        df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE],  errors="coerce")

        df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
        df = (df.sort_values(["_no_exit", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
                .drop(columns=["_no_exit"]))

        # Filter considered cohort: Start date between 5/1/2026 and 8/31/2026
        in_considered_cohort = (
            (df[COL_START_DATE] >= cohort_start) & 
            (df[COL_START_DATE] <= q_end_ts)
        )

        day_31_date = df[COL_START_DATE] + pd.Timedelta(days=31)
        exit_date_filled = df[COL_EXIT_DATE].fillna(far_future)

        # 1. Achieved 31 days (Reached Day 31 within Q4 window)
        reached_31_in_q = (
            in_considered_cohort &
            (day_31_date >= q_start_ts) & 
            (day_31_date <= q_end_ts) & 
            (exit_date_filled >= day_31_date)
        )
        df["Achieved 31 days"] = reached_31_in_q.astype(int)

        # 2. Still in Program (After 31days): Active at Q End AND tenure >= 31 days
        is_active_at_q_end = exit_date_filled > q_end_ts
        tenure_at_q_end = (q_end_ts - df[COL_START_DATE]).dt.days
        still_31 = (
            in_considered_cohort &
            is_active_at_q_end & 
            (tenure_at_q_end >= 31)
        )
        df["Still in Program (After 31days)"] = still_31.astype(int)

        # Additional metadata and indicators
        df["City"]     = df[COL_PROGRAM].apply(constants.assign_city)
        df["Capacity"] = df[COL_PROGRAM].map(constants.PROGRAM_CAPACITY)
        
        df["Year"]    = df[COL_START_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_START_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]  = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        gap = (df[COL_EXIT_DATE].fillna(q_end_ts) - df[COL_START_DATE]).dt.days

        df["Exit Reason Category"]  = df[COL_EXIT_REASON].apply(categorize_exit_reason)
        df["Entered Since Prior FY"] = (df[COL_EXIT_DATE].isna() | (gap > 30)).astype(int)
        df["Exit After 30 Days"]     = (df[COL_EXIT_DATE].notna() & (gap > 30) & (~df[COL_EXIT_REASON].isin(GRAD_REASONS))).astype(int)
        df["Exit Before 30 Days"]  = (
            df[COL_EXIT_DATE].notna() &
            (df[COL_EXIT_DATE] >= q_start_ts) &
            (df[COL_EXIT_DATE] <= q_end_ts) &
            (gap < 30)
        ).astype(int)
        df["Graduated"]              = (df[COL_EXIT_REASON].isin(GRAD_REASONS)).astype(int)
        df["Still in Program"]       = (in_considered_cohort & is_active_at_q_end).astype(int)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
                "Exit Reason Category", "City", "Year", "Quarter", "Year Q", "Capacity",
                "Entered Since Prior FY", "Exit After 30 Days", "Exit Before 30 Days",
                "Graduated", "Still in Program", "Still in Program (After 31days)", "Achieved 31 days"]
        df = df[[c for c in cols if c in df.columns]].copy()

        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        return df.reset_index(drop=True)

    # =========================================================================
    # 5. Execute Tab Processing
    # =========================================================================
    print("Processing Graduates...")
    df_grad = process_graduates(df_base)

    print("Processing Occupancy...")
    df_occ = process_occupancy(df_base)

    print("Processing Retention...")
    df_ret = process_retention(df_base)

    # =========================================================================
    # 6. Google Sheets Export
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

    write_tab(spreadsheet, "Raw Data",        df_raw)
    write_tab(spreadsheet, "Renew Graduates", df_grad)
    write_tab(spreadsheet, "Occupancy",       df_occ)
    write_tab(spreadsheet, "Retention",       df_ret)

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Processing complete!: {spreadsheet.url}")


if __name__ == "__main__":
    run_renew_processing(
        print("🚀 Starting Renew Processing"),
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )