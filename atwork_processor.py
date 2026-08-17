# ============================================================
# This script processes @Work KPIs and Barriers
#
# 1. Reads 2 XLSX files from Google Drive
# 2. Creates "@Work KPIs and Barriers - Processed" Google Sheet
#       - with Barriers and @Work processed tabs
# ============================================================

import warnings
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

    COL_ID          = "Record Id_152"
    COL_PROG        = "Program Enrolling_2091"
    COL_INTERN      = "Intern Program_6619"
    COL_DATE        = "Start Date_2090"
    COL_DATE_ENTER  = "Date Entered CityTeam @ Work_3842"
    COL_BARRIER     = "Employment Barriers_3870"
    COL_ASSESS      = "All Assessments Complete (Mod 2 Eligible)_3864"

    VALID_PROGRAMS_ALL = {
        "Chester Renew", "Heritage Home", "Oakland Renew",
        "Oakland Women Turning Point", "Portland Renew",
        "San Jose Men Renew", "San Jose Men Turning Point",
        "Chester Men Turning Point", "Chester Women Turning Point",
        "GV Renew", "GV Turning Point", "House of Grace",
        "Oakland Men Turning Point", "Portland Community of Hope",
        "Program Graduate Intern", "San Jose Youth Collective"
    }

    VALID_PROGRAMS_NUM_DEN = {"San Jose Men Renew", "GV Renew"}

    VALID_PROGRAMS_IMPACT = {
        "Chester Renew", "Heritage Home", "Oakland Renew",
        "Oakland Women Turning Point", "Portland Renew",
        "San Jose Men Renew", "San Jose Men Turning Point",
        "Chester Men Turning Point", "Chester Women Turning Point",
        "GV Renew", "GV Turning Point", "House of Grace",
        "Oakland Men Turning Point", "Portland Community of Hope",
        "Program Graduate Intern", "San Jose Youth Collective"
    }
    VALID_INTERN_PROGRAMS_IMPACT = VALID_PROGRAMS_IMPACT - {"Program Graduate Intern"}

    BARRIER_REMAP = {
        "Criminal History or CASU"   : "Justice System Involved",
        "Suspended driver's licence" : "Suspended or No Driver's License",
        "Suspended driver's license" : "Suspended or No Driver's License",
        "No driver's license"        : "Suspended or No Driver's License",
        "No DL"                      : "Suspended or No Driver's License",
    }

    IC_DATE_COLS = [
        "Electrical Trainee Card Date_5727", "NCCER Date_5728",
        "Blueprint Reading Date_6785", "Forklift Operator Date_5106",
        "OSHA10 Date_5118", "MA Date_5115",
        "Google IT Support Certificate Date_5107", "HAZWOPER Date_5109",
        "HAZMAT Date_5108", "CPR/FA/AED Date_5105", "CADC Date_5789",
        "ServSafe Food Manager Date_5126", "ServSafe Food Handler Date_5125",
        "Material Handling Date_5116",
        "Multi-Core Craft Curriculum Completion Date_5117",
        "RADT Date_5121", "HVAC Date_5113", "CDL Date_6104",
        "CompTIA and/or A+ Date_5104", "OTHER Date_5119",
        "Guard Card Date_6561", "TWIC Date_6631",
        "Google Automate Cybersecurity w/Python Date_6273",
        "CNC Machining Date_5103", "Confined Space Date_6788",
        "DVAT Date_6791", "Fundamentals of Construction Date_6790",
        "HR Associate Date_6789", "Intuit Bookkeeping Date_6105",
        "Massage Therapy Cert Date_7032", "Peer Support Date_7031",
        "Record Expungement Date_6787",
        "Record Expungement Attorney Consultation Date_6879",
        "Traffic Control Date_6786",
    ]

    EXTRA_WINDOW_COLS = [
        "High School Equivalency (HSE) Date_5110",
        "Bank Account Obtained Date_6640",
        "Driver's License Date_5071",
        "Birth Certificate Obtained Date_6638",
        "State ID Obtained Date_6637",
        "Social Security Card Obtained Date_6639",
        "Date Job Acquired_5590",
    ]

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
    # 2. Download and Load Raw Data
    # =========================================================================
    print(f"Downloading {barriers_input_file}...")
    b_fh, _, _ = download_drive_file(drive_service, barriers_input_file, input_folder_id)
    df_raw_barriers = load_raw(b_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {atwork_input_file}...")
    a_fh, _, _ = download_drive_file(drive_service, atwork_input_file, input_folder_id)
    df_raw_atwork = load_raw(a_fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)


    # =========================================================================
    # 3. Data Processing Internal Functions
    # =========================================================================
    def parse_date_flexible(val):
        if pd.isna(val) or str(val).strip() == "": return pd.NaT
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%d/%m/%Y"):
            try: return pd.to_datetime(val, format=fmt)
            except: pass
        return pd.to_datetime(val, errors="coerce")

    def needs_dl(val):
        if pd.isna(val): return 0
        return 1 if ("No driver's license" in str(val) or "Suspended driver's license" in str(val)) else 0

    def process_barriers_cleaned(df_raw):
        df = df_raw.copy()
        df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
        df = df[df[COL_DATE].notna() & (df[COL_DATE] >= constants.CURRENT_FY_QX_START) & (df[COL_DATE] <= constants.CURRENT_Q_END)]
        df = df.sort_values(COL_DATE, ascending=False).drop_duplicates(subset=[COL_ID], keep="first")
        df = df[df[COL_PROG].isin(VALID_PROGRAMS_ALL)]

        df["Need HSE"] = df[COL_BARRIER].apply(lambda v: 1 if pd.notna(v) and "No HSE" in str(v) else 0) if COL_BARRIER in df.columns else 0
        df["Need DL"]  = df[COL_BARRIER].apply(needs_dl) if COL_BARRIER in df.columns else 0
        return df.reset_index(drop=True)

    def process_numerator_denominator(df_raw):
        df_src = df_raw.copy()
        if COL_DATE_ENTER in df_src.columns:
            df_src[COL_DATE_ENTER] = pd.to_datetime(df_src[COL_DATE_ENTER], errors="coerce")
            df_src = df_src[df_src[COL_DATE_ENTER].notna() & (df_src[COL_DATE_ENTER] >= constants.CURRENT_FY_QX_START) & (df_src[COL_DATE_ENTER] <= constants.CURRENT_Q_END)]
        
        df_src = df_src[df_src[COL_PROG].isin(VALID_PROGRAMS_NUM_DEN)]
        if COL_ASSESS in df_src.columns:
            df_src = df_src[df_src[COL_ASSESS].astype(str).str.strip() == "Yes"]

        df_den = df_src.drop_duplicates(subset=[COL_ID], keep="first")
        df_den = df_den.drop_duplicates(subset=["Name_2057"], keep="first").reset_index(drop=True)

        def expand_barriers(row):
            raw = str(row[COL_BARRIER]) if pd.notna(row.get(COL_BARRIER)) else ""
            barriers = [BARRIER_REMAP.get(b.strip(), b.strip()) for b in raw.split("|") if b.strip()]
            if not barriers: return [{"Barrier": "No Barriers", "Barrier Count": 0}]
            return [{"Barrier": b, "Barrier Count": len(barriers)} for b in barriers]

        expanded_rows = []
        for _, row in df_den.iterrows():
            for extra in expand_barriers(row):
                expanded_rows.append({**row.to_dict(), **extra})
        df_num = pd.DataFrame(expanded_rows).reset_index(drop=True)

        return df_num, df_den

    def process_atwork_kpis(df_raw):
        df = df_raw.copy()
        date_cols_all = [c for c in df.columns if "Date" in c]
        non_date_cols = [c for c in df.columns if c not in date_cols_all and c != COL_ID]

        for col in date_cols_all: df[col] = df[col].apply(parse_date_flexible)

        def merge_client_rows(group):
            group = group.sort_values(COL_DATE, ascending=False)
            merged = {}
            for col in non_date_cols: merged[col] = group[col].iloc[0]
            for col in date_cols_all:
                vals = group[col].dropna()
                merged[col] = vals.max() if not vals.empty else pd.NaT
            return pd.Series(merged)

        df = df.groupby(COL_ID, sort=False).apply(merge_client_rows, include_groups=False).reset_index()

        existing_ic_cols  = [c for c in IC_DATE_COLS if c in df.columns]
        existing_ext_cols = [c for c in EXTRA_WINDOW_COLS if c in df.columns]

        start_parsed = df[COL_DATE] if COL_DATE in df.columns else pd.Series(pd.NaT, index=df.index)
        
        def is_eligible(row):
            prog = str(row.get(COL_PROG, "")).strip()
            if prog == "Program Graduate Intern":
                return str(row.get(COL_INTERN, "")).strip() in VALID_INTERN_PROGRAMS_IMPACT
            return prog in VALID_PROGRAMS_IMPACT

        program_eligible = df.apply(is_eligible, axis=1)
        start_in_window  = start_parsed.notna() & (start_parsed >= constants.CURRENT_FY_QX_START) & (start_parsed <= constants.CURRENT_Q_END)
        need_mask        = program_eligible & start_in_window

        COL_HSE = "Actively Pursing HSE_5575"
        df["Need HSE"] = df[COL_HSE].astype(str).str.strip().apply(lambda v: 1 if v == "Yes" else 0).where(need_mask, other=0) if COL_HSE in df.columns else 0
        df["Need DL"] = df[COL_BARRIER].apply(needs_dl).where(need_mask, other=0) if COL_BARRIER in df.columns else 0

        for col in existing_ic_cols + existing_ext_cols:
            df[col] = df[col].where(df[col].notna() & (df[col] >= constants.CURRENT_FY_QX_START) & (df[col] <= constants.CURRENT_Q_END), other=pd.NaT)

        df["Total IC"]   = df[existing_ic_cols].notna().sum(axis=1)
        df["At least 1"] = (df["Total IC"] >= 1).astype(int)

        df["City"] = df.apply(lambda r: str(r.get(COL_INTERN, "")).strip() if str(r.get(COL_PROG, "")).strip() == "Program Graduate Intern" and pd.notna(r.get(COL_INTERN)) else str(r.get(COL_PROG, "")).strip(), axis=1).apply(constants.assign_city)

        for col in date_cols_all:
            if col in df.columns: df[col] = df[col].apply(lambda x: x.strftime("%m/%d/%Y") if pd.notna(x) else "")
        return df.reset_index(drop=True)

    # =========================================================================
    # 4. Execute Processing
    # =========================================================================
    print("Processing Barriers Cleaned...")
    df_bar_cleaned = process_barriers_cleaned(df_raw_barriers)

    print("Processing Numerator/Denominator...")
    df_num, df_den = process_numerator_denominator(df_raw_barriers)

    print("Processing @Work KPIs...")
    df_atwork_proc = process_atwork_kpis(df_raw_atwork)

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
    write_tab(spreadsheet, "Barriers Raw Data", df_raw_barriers)
    
    # Prep dates for writing
    df_bar_cleaned_write = df_bar_cleaned.copy()
    df_bar_cleaned_write[COL_DATE] = pd.to_datetime(df_bar_cleaned_write[COL_DATE], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    write_tab(spreadsheet, "Barriers Cleaned", df_bar_cleaned_write)

    df_num_write = df_num.copy()
    if COL_DATE_ENTER in df_num_write.columns:
        df_num_write[COL_DATE_ENTER] = pd.to_datetime(df_num_write[COL_DATE_ENTER], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    write_tab(spreadsheet, "Barriers Numerator", df_num_write)

    df_den_write = df_den.copy()
    if COL_DATE_ENTER in df_den_write.columns:
        df_den_write[COL_DATE_ENTER] = pd.to_datetime(df_den_write[COL_DATE_ENTER], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    write_tab(spreadsheet, "Barriers Denominator", df_den_write)

    write_tab(spreadsheet, "@Work Raw Data", df_raw_atwork)
    write_tab(spreadsheet, "@Work Processed", df_atwork_proc)


    # Delete lingering first tab (Sheet1)
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except: pass

    print(f"\n🎉 Done!: {spreadsheet.url}")


if __name__ == "__main__":
    run_atwork_processing(
         input_folder_name="Apricot Report Incoming",
         output_folder_name="KPI Processed Data"
    )
