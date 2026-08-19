#!/usr/bin/env python3
"""
This script processes @Work KPIs and Barriers
"""

import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def run_atwork_processing(
    barriers_input_file="Renew Client Barriers at Entry.xlsx",
    atwork_input_file="All Programs @Work Fiscal Year Impacts.xlsx",
    output_file="@Work KPIs and Barriers - Processed",
    input_folder_name=None,
    output_folder_name=None
):
    # =========================================================================
    # Spreadsheet-Specific Constants
    # =========================================================================
    RAW_DATA_HEADER_ROW = 3
    RAW_DATA_START_COL  = 1

    COL_ID         = "Record Id_152"
    COL_PROG       = "Program Enrolling_2091"
    COL_INTERN     = "Intern Program_6619"
    COL_DATE       = "Start Date_2090"
    COL_DATE_ENTER = "Date Entered CityTeam @ Work_3842"
    COL_BARRIER    = "Employment Barriers_3870"
    COL_ASSESS     = "All Assessments Complete (Mod 2 Eligible)_3864"
    COL_HSE        = "Actively Pursing HSE_5575"

    VALID_PROGRAMS_ALL           = constants.VALID_PROGRAMS_ALL
    VALID_PROGRAMS_NUM_DEN       = constants.VALID_PROGRAMS_NUM_DEN
    VALID_PROGRAMS_IMPACT        = constants.VALID_PROGRAMS_IMPACT
    VALID_INTERN_PROGRAMS_IMPACT = constants.VALID_INTERN_PROGRAMS_IMPACT
    BARRIER_REMAP                = constants.BARRIER_REMAP
    IC_DATE_COLS                 = constants.IC_DATE_COLS
    EXTRA_WINDOW_COLS            = constants.EXTRA_WINDOW_COLS

    # =========================================================================
    # 1. Initialize Services
    # =========================================================================
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and Load Raw Data
    # =========================================================================
    print(f"Downloading {barriers_input_file}...")
    b_fh, _, _ = download_drive_file(drive_service, barriers_input_file, input_folder_id)
    df_raw_barriers = load_raw(b_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {atwork_input_file}...")
    a_fh, _, _ = download_drive_file(drive_service, atwork_input_file, input_folder_id)
    df_raw_atwork = load_raw(a_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # =========================================================================
    # 3. Build Need DL lookup from Barriers Raw Data
    # =========================================================================
    _dl_mask = df_raw_barriers[COL_BARRIER].astype(str).str.contains(
        "No driver's license|Suspended driver's license", na=False
    )
    df_dl_lookup = df_raw_barriers[[COL_ID]].copy()
    df_dl_lookup["Need DL"] = _dl_mask.astype(int)
    df_dl_lookup = df_dl_lookup.groupby(COL_ID, sort=False)["Need DL"].max().reset_index()

    # =========================================================================
    # 4. Processing Functions
    # =========================================================================
    def process_barriers_cleaned(df_raw):
        df = df_raw.copy()
        df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
        df = df[
            df[COL_DATE].notna() &
            (df[COL_DATE] >= constants.CURRENT_FY_QX_START) &
            (df[COL_DATE] <= constants.CURRENT_Q_END)
        ]
        df = df.sort_values(COL_DATE, ascending=False).drop_duplicates(subset=[COL_ID], keep="first")
        df = df[df[COL_PROG].isin(VALID_PROGRAMS_ALL)].reset_index(drop=True)
        return df

    def process_barriers_numerator(df_raw):
        df_src = df_raw.copy()
        if COL_DATE_ENTER in df_src.columns:
            df_src[COL_DATE_ENTER] = pd.to_datetime(df_src[COL_DATE_ENTER], errors="coerce")
            df_src = df_src[
                df_src[COL_DATE_ENTER].notna() &
                (df_src[COL_DATE_ENTER] >= constants.CURRENT_FY_QX_START) &
                (df_src[COL_DATE_ENTER] <= constants.CURRENT_Q_END)
            ]

        df_src = df_src[df_src[COL_PROG].isin(VALID_PROGRAMS_NUM_DEN)]
        if COL_ASSESS in df_src.columns:
            df_src = df_src[df_src[COL_ASSESS].astype(str).str.strip() == "Yes"]

        # Denominator: dedup by Record Id then Name
        df_den = df_src.drop_duplicates(subset=[COL_ID], keep="first")
        df_den = df_den.drop_duplicates(subset=["Name_2057"], keep="first").reset_index(drop=True)

        # Numerator: expand barriers using vectorized explode
        df_den = df_den.copy()
        df_den["_barriers_list"] = df_den[COL_BARRIER].fillna("").astype(str).apply(
            lambda v: [BARRIER_REMAP.get(b.strip(), b.strip()) for b in v.split("|") if b.strip()] or ["No Barriers"]
        )
        df_den["Barrier Count"] = df_den["_barriers_list"].apply(
            lambda x: 0 if x == ["No Barriers"] else len(x)
        )
        df_num = df_den.explode("_barriers_list").rename(columns={"_barriers_list": "Barrier"})
        df_num = df_num.drop(columns=[COL_BARRIER], errors="ignore").reset_index(drop=True)

        return df_num

    def process_atwork_kpis(df_raw):
        df = df_raw.copy()
        date_cols_all     = [c for c in df.columns if "Date" in c]
        non_date_cols     = [c for c in df.columns if c not in date_cols_all and c != COL_ID]
        existing_ic_cols  = [c for c in IC_DATE_COLS if c in df.columns]
        existing_ext_cols = [c for c in EXTRA_WINDOW_COLS if c in df.columns]

        # Parse all date columns at once
        for col in date_cols_all:
            df[col] = pd.to_datetime(df[col], errors="coerce")

        # Merge client rows — vectorized aggregation
        agg_dict = {col: "first" for col in non_date_cols}
        agg_dict.update({col: "max" for col in date_cols_all})
        df = df.sort_values(COL_DATE, ascending=False).groupby(COL_ID, sort=False).agg(agg_dict).reset_index()

        win_start = constants.CURRENT_FY_QX_START
        win_end   = constants.CURRENT_Q_END

        start_parsed = pd.to_datetime(df[COL_DATE], errors="coerce") if COL_DATE in df.columns else pd.Series(pd.NaT, index=df.index)

        # Vectorized eligibility
        prog_series  = df[COL_PROG].astype(str).str.strip()
        intern_series = df[COL_INTERN].astype(str).str.strip() if COL_INTERN in df.columns else pd.Series("", index=df.index)
        is_pgi       = prog_series == "Program Graduate Intern"
        program_eligible = (
            (~is_pgi & prog_series.isin(VALID_PROGRAMS_IMPACT)) |
            (is_pgi & intern_series.isin(VALID_INTERN_PROGRAMS_IMPACT))
        )
        start_in_window = start_parsed.notna() & (start_parsed >= win_start) & (start_parsed <= win_end)
        need_mask       = program_eligible & start_in_window

        # Need HSE
        if COL_HSE in df.columns:
            df["Need HSE"] = (df[COL_HSE].astype(str).str.strip() == "Yes").astype(int).where(need_mask, other=0)
        else:
            df["Need HSE"] = 0

        # Need DL — from barriers lookup
        df = df.merge(df_dl_lookup, on=COL_ID, how="left")
        df["Need DL"] = df["Need DL"].fillna(0).astype(int).where(need_mask, other=0)

        # IC date cols — null if outside window
        for col in existing_ic_cols + existing_ext_cols:
            if col in df.columns:
                df[col] = df[col].where(
                    df[col].notna() & (df[col] >= win_start) & (df[col] <= win_end),
                    other=pd.NaT
                )

        df["Total IC"]   = df[[c for c in existing_ic_cols if c in df.columns]].notna().sum(axis=1)
        df["At least 1"] = (df["Total IC"] >= 1).astype(int)

        # City — vectorized
        city_source = pd.Series(np.where(
            is_pgi & intern_series.notna() & (intern_series != "nan"),
            intern_series, prog_series
        ), index=df.index)
        df["City"] = city_source.apply(constants.assign_city)

        # Convert date cols to strings
        for col in date_cols_all:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: x.strftime("%m/%d/%Y") if pd.notna(x) else "")

        return df.reset_index(drop=True)

    # =========================================================================
    # 5. Execute Processing
    # =========================================================================
    print("Processing Barriers Cleaned...")
    df_bar_cleaned = process_barriers_cleaned(df_raw_barriers)

    print("Processing Barriers Numerator...")
    df_num = process_barriers_numerator(df_raw_barriers)

    print("Processing @Work KPIs...")
    df_atwork_proc = process_atwork_kpis(df_raw_atwork)

    # =========================================================================
    # 6. Handle Output Spreadsheet
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
    # 7. Export Tabs
    # =========================================================================
    write_tab(spreadsheet, "Barriers Raw Data", df_raw_barriers)

    df_bar_cleaned_write = df_bar_cleaned.copy()
    df_bar_cleaned_write[COL_DATE] = pd.to_datetime(
        df_bar_cleaned_write[COL_DATE], errors="coerce"
    ).dt.strftime("%Y-%m-%d").fillna("")
    write_tab(spreadsheet, "Barriers Cleaned", df_bar_cleaned_write)

    df_num_write = df_num.copy()
    if COL_DATE_ENTER in df_num_write.columns:
        df_num_write[COL_DATE_ENTER] = pd.to_datetime(
            df_num_write[COL_DATE_ENTER], errors="coerce"
        ).dt.strftime("%Y-%m-%d").fillna("")
    write_tab(spreadsheet, "Barriers Numerator", df_num_write)

    write_tab(spreadsheet, "@Work Raw Data", df_raw_atwork)
    write_tab(spreadsheet, "@Work Processed", df_atwork_proc)

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Done!: {spreadsheet.url}")


if __name__ == "__main__":
    run_atwork_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )