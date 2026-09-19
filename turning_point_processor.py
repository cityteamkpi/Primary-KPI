#!/usr/bin/env python3
"""
This script processes the Turning Point Report
"""

import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_turning_point_processing(
    input_file="Turning Point Report.xlsx",
    output_file="Turning Point Report - Processed",
    input_folder_name=None,
    output_folder_name=None
):
    # =========================================================================
    print("🚀 Starting Turning Point Processing")
    # Spreadsheet-Specific Constants
    # =========================================================================
    RAW_DATA_HEADER_ROW = 3
    RAW_DATA_START_COL  = 1

    COL_RECORD_ID      = "Record Id_102"
    COL_PROGRAM        = "Program Enrolling_2091"
    COL_START_DATE     = "Start Date_2090"
    COL_EXIT_DATE      = "Exit Date_2100"
    COL_EXIT_REASON    = "Primary Reason for Exit_2102"
    COL_HOUSED         = "Successfully Housed (Is Housing Healthy?)_4368"
    COL_INTERN_PROGRAM = "Intern Program_6619"

    PROGRAMS_TO_INCLUDE = [
        "Heritage Home", "Oakland Women Turning Point", "San Jose Men Turning Point",
        "Chester Men Turning Point", "Chester Women Turning Point", "GV Turning Point",
        "Oakland Men Turning Point", "Oakland Youth Collective", "Portland Community of Hope",
        "Portland Youth Collective", "San Jose Youth Collective", "Chester Turning Point",
    ]

    GRAD_REASONS      = constants.GRAD_REASONS
    COMPLETER_REASONS = ["Completer", "Completion of Program"]

    # =========================================================================
    # 1. Initialize Services
    # =========================================================================
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and Load Data
    # =========================================================================
    print(f"Downloading {input_file}...")
    fh, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    df_raw = load_raw(fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # Initial cleanup
    prog_lower_map = {p.lower(): p for p in PROGRAMS_TO_INCLUDE}
    def match_program(val):
        if pd.isna(val): return val
        return prog_lower_map.get(str(val).strip().lower(), str(val).strip())

    df_base = df_raw.copy()
    df_base[COL_PROGRAM]    = df_base[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
    df_base[COL_START_DATE] = pd.to_datetime(df_base[COL_START_DATE], errors="coerce")
    df_base[COL_EXIT_DATE]  = pd.to_datetime(df_base[COL_EXIT_DATE],  errors="coerce")

    # Remap Program Graduate Interns to their Intern Program
    if COL_INTERN_PROGRAM in df_base.columns:
        is_intern = df_base[COL_PROGRAM] == "Program Graduate Intern"
        df_base.loc[is_intern, COL_PROGRAM] = df_base.loc[is_intern, COL_INTERN_PROGRAM].apply(match_program)

    df_base = df_base[df_base[COL_PROGRAM].isin(PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

    # =========================================================================
    # 3. Processing Functions
    # =========================================================================
    def fmt_date(df, col):
        return pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")

    def process_exited_category(df_base, reasons_list=None):
        df = df_base.copy()
        df["_priority"] = df[COL_EXIT_REASON].isin(reasons_list).astype(int) if reasons_list else 0
        df = (df.sort_values(["_priority", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
                .drop(columns=["_priority"]))

        mask = df[COL_EXIT_DATE].notna()
        if reasons_list:
            mask &= df[COL_EXIT_REASON].isin(reasons_list)
        df = df[mask].copy()

        df["City"]        = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]        = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]     = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]      = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        df["Next FY Goal"] = df[COL_PROGRAM].map(constants.NEXT_FY_GRAD_GOALS_TP)

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= w_start) & (df[COL_EXIT_DATE] <= w_end)).astype(int)

        # Explicit column selection
        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q", "Next FY Goal"] + list(constants.ACTUALS_WINDOWS.keys())
        df = df[[c for c in cols if c in df.columns]].copy()
        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        for k in constants.ACTUALS_WINDOWS.keys():
            if k in df.columns: df[k] = df[k].fillna(0).astype(int)
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

        COL_CHILD = "Name of Child_6313"

        # For GV TP: if deduped row has null Name of Child, look up from any row with a value
        if COL_CHILD in df.columns:
            is_gv_tp_null = (df[COL_PROGRAM] == "GV Turning Point") & df["Name of Child_6313"].isna()
            if is_gv_tp_null.any():
                child_lookup = df_base[df_base[COL_PROGRAM] == "GV Turning Point"].dropna(subset=["Name of Child_6313"])
                child_lookup = child_lookup.groupby(COL_RECORD_ID)["Name of Child_6313"].first()
                df.loc[is_gv_tp_null, "Name of Child_6313"] = df.loc[is_gv_tp_null, COL_RECORD_ID].map(child_lookup)



        # Number of Children — counts kids only for women's programs, no active filter
        WOMEN_PROGRAMS = [
            "Chester Women Turning Point", "Oakland Women Turning Point",
            "GV Turning Point", "Heritage Home", "San Jose Youth Collective"
        ]
        def count_kids_only(row):
            if row[COL_PROGRAM] not in WOMEN_PROGRAMS: return 0
            val = row.get(COL_CHILD)
            if pd.isna(val) or str(val).strip() == "": return 0
            import re as _re
            parts = _re.split(r'[,&]|\band\b', str(val).strip(), flags=_re.IGNORECASE)
            return len([p for p in parts if p.strip()])
        df["Number of Children"] = df.apply(count_kids_only, axis=1)

        # Ensure OCC_CURRENT_LABEL is always int
        df[constants.OCC_CURRENT_LABEL] = pd.to_numeric(df[constants.OCC_CURRENT_LABEL], errors="coerce").fillna(0).astype(int)

        # Sync Current Period Actuals with OCC_CURRENT_LABEL (includes child count for GV TP)
        if "Current Period Actuals" in df.columns:
            df["Current Period Actuals"] = df[constants.OCC_CURRENT_LABEL]
        df["Capacity"]           = df[COL_PROGRAM].map(constants.TP_CAPACITY)
        df["Goal"]               = df[COL_PROGRAM].map(constants.TP_GOAL)
        df["Next FY Goal"]       = df[COL_PROGRAM].map(constants.NEXT_FY_OCC_GOALS_TP)
        df["Prior FY Occupancy"] = df[COL_PROGRAM].map(constants.TP_PRIOR_FY)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE,
                "City", constants.OCC_PRIOR_LABEL, constants.OCC_CURRENT_LABEL,
                "Name of Child_6313", "Number of Children", "Capacity", "Goal", "Next FY Goal", "Prior FY Occupancy"]
        df = df[[c for c in cols if c in df.columns]].copy()
        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        for col in [constants.OCC_PRIOR_LABEL, constants.OCC_CURRENT_LABEL,
                    "Capacity", "Goal", "Prior FY Occupancy"]:
            if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df.reset_index(drop=True)

    def process_housed(df_raw):
        # Start from df_raw so intern rows are present before any program filtering
        df = df_raw.copy()
        df[COL_PROGRAM]    = df[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
        df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
        df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE],  errors="coerce")

        # Remap intern rows
        if COL_INTERN_PROGRAM in df.columns:
            is_intern = df[COL_PROGRAM] == "Program Graduate Intern"
            df.loc[is_intern, COL_PROGRAM] = df.loc[is_intern, COL_INTERN_PROGRAM].apply(match_program)
            print(f"   Remapped {is_intern.sum()} Program Graduate Intern rows")

        df = df[df[COL_PROGRAM].isin(PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

        df["_is_housed"]        = df[COL_HOUSED].astype(str).str.contains("Yes", case=False, na=False).astype(int) if COL_HOUSED in df.columns else 0
        df["_is_exited_housed"] = (df[COL_EXIT_DATE].notna() & (df["_is_housed"] == 1)).astype(int)

        df = (df.sort_values(COL_EXIT_DATE, ascending=False)
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM], keep="first")
                .drop(columns=["_is_housed", "_is_exited_housed"]))

        df["City"]                = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]                = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]             = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]              = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        df["Successfully Housed?"] = df[COL_HOUSED].astype(str).str.contains("Yes", case=False, na=False).astype(int) if COL_HOUSED in df.columns else 0
        df["Capacity"]            = df[COL_PROGRAM].map(constants.TP_CAPACITY)

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (
                df[COL_EXIT_DATE].notna() &
                (df[COL_EXIT_DATE] >= w_start) &
                (df[COL_EXIT_DATE] <= w_end) &
                (df["Successfully Housed?"] == 1)
            ).astype(int)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q", "Successfully Housed?", "Capacity"
                ] + list(constants.ACTUALS_WINDOWS.keys())
        df = df[[c for c in cols if c in df.columns]].copy()
        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        for k in constants.ACTUALS_WINDOWS.keys():
            if k in df.columns: df[k] = df[k].fillna(0).astype(int)
        return df.reset_index(drop=True)

    # =========================================================================
    # 4. Execute Processing
    # =========================================================================
    print("Processing Graduates...")
    df_grad = process_exited_category(df_base, reasons_list=["Graduation"])



    print("Processing Occupancy...")
    df_occ = process_occupancy(df_base)

    print("Processing Housed...")
    df_housed = process_housed(df_raw)

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
    write_tab(spreadsheet, "Raw Data",                 df_raw)
    write_tab(spreadsheet, "Turning Point Graduates",  df_grad)
    write_tab(spreadsheet, "Turning Point Occupancy",  df_occ)
    write_tab(spreadsheet, "Successfully Housed",      df_housed)

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Done!: {spreadsheet.url}")


if __name__ == "__main__":
    run_turning_point_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )