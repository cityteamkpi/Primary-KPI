import warnings
import os
import time
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import (
    resolve_folder_id, 
    download_drive_file, 
    load_raw, 
    write_tab, 
    find_file_id
)
import constants

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ============================================================
#  CONFIGURATION
# ============================================================
SRC_ALUMNI = "All Programs All Alumni KPIs"
OUTPUT_SHEET_NAME = "Alumni KPI - Processed"

RAW_DATA_HEADER_ROW = 3
RAW_DATA_START_COL = 1

COL_NAME = "Name_2057"
COL_PROGRAM = "Program Enrolling_2091"
COL_EXIT_DATE = "Exit Date_2100"
COL_EXIT_REASON = "Primary Reason for Exit_2102"
COL_CHECKIN = "Checkin Date_6583"
COL_WAGE = "Hourly Wage_6586"
COL_HOUSING = "Housing Expense_6587"
COL_SOBER = "Sober?_6585"
COL_HOW_CHECKED = "How Checked_6584"

GRAD_REASONS = ["Graduation", "Completer"]
RENEW_PROGRAMS = {"Chester Renew", "GV Renew", "Oakland Renew", "Portland Renew", "San Jose Men Renew"}
MONTHLY_HOURS = 173.33

def clean_wage(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    clean = str(val).replace("$", "").replace(",", "").strip()
    try: return float(clean)
    except: return 0.0

def max_nonzero(series):
    vals = series.replace(0, np.nan).dropna()
    return vals.max() if not vals.empty else 0.0

def analyze_client(name, program, grad_date, df_raw_parsed, df_merged):
    lw_target = constants.LW_CRITERIA.get(program, 0.0) or 0.0
    
    person_records = df_raw_parsed[
        (df_raw_parsed[COL_NAME] == name) & (df_raw_parsed[COL_PROGRAM] == program)
    ].sort_values(by=COL_CHECKIN)

    # --- Sobriety Logic ---
    sober_records = person_records[person_records[COL_SOBER].isin(["Yes", "No"])] if COL_SOBER in person_records.columns else pd.DataFrame()
    has_sustained_relapse = False
    last_state = None
    first_no_date = None
    last_checkin_sober = False
    
    for _, s_row in sober_records.iterrows():
        cur_state = s_row[COL_SOBER]
        cur_date = s_row[COL_CHECKIN]
        if last_state == "Yes" and cur_state == "No":
            if first_no_date is None: first_no_date = cur_date
        elif last_state == "No" and cur_state == "Yes":
            if first_no_date is not None:
                if (cur_date - first_no_date).days > 30: 
                    has_sustained_relapse = True
            first_no_date = None
        last_state = cur_state
        last_checkin_sober = (cur_state == "Yes")
    
    is_sober = last_checkin_sober and not has_sustained_relapse

    # --- Wage & Housing Logic ---
    grad_record = person_records[
        person_records[COL_HOW_CHECKED].astype(str).str.contains("Graduation", na=False)
    ] if COL_HOW_CHECKED in person_records.columns else pd.DataFrame()
    
    merged_row = df_merged[(df_merged[COL_NAME] == name) & (df_merged[COL_PROGRAM] == program)]
    
    grad_wage = grad_record[COL_WAGE].replace(0, np.nan).dropna().max() if not grad_record.empty else (
        merged_row[COL_WAGE].iloc[0] if not merged_row.empty else 0.0)
    
    recent_wage = merged_row[COL_WAGE].iloc[0] if not merged_row.empty else 0.0
    recent_checkin = person_records.iloc[-1][COL_CHECKIN] if not person_records.empty else pd.NaT

    # Housing Calculation
    grad_under_30 = None
    recent_under_30 = None
    if not merged_row.empty and COL_HOUSING in merged_row.columns:
        income = merged_row["Monthly Gross Income"].iloc[0]
        housing = merged_row[COL_HOUSING].iloc[0]
        grad_under_30 = 1 if (housing == 0) else (1 if (income > 0 and (housing / income) < 0.30) else 0)
        recent_under_30 = grad_under_30 # In this specific report, recent matches merged max

    return {
        COL_NAME: name,
        COL_PROGRAM: program,
        "Graduation Date": grad_date,
        "Most Recent Checkin": recent_checkin,
        "City": constants.assign_city(program),
        "Year": constants.get_fiscal_year(grad_date),
        "Quarter": constants.get_fiscal_quarter(grad_date),
        "Year Q": f"{constants.get_fiscal_year(grad_date) or ''} {constants.get_fiscal_quarter(grad_date) or ''}".strip(),
        "is_sober": is_sober,
        "Sustained Relapse?": "Yes" if has_sustained_relapse else "No",
        "Wage (Grad)": grad_wage,
        "LW Criteria": lw_target,
        "Pays LW? (Grad)": int(grad_wage >= lw_target) if lw_target else 0,
        "Wage (Recent)": recent_wage,
        "Pays LW? (Recent)": int(recent_wage >= lw_target) if lw_target else 0,
        "Housing Under 30% (Grad)": 0 if grad_under_30 is None else int(grad_under_30),
        "Housing Under 30% (Recent)": 0 if recent_under_30 is None else int(recent_under_30),
    }

def run_alum_processing(
    input_file=SRC_ALUMNI + ".xlsx",
    output_file=OUTPUT_SHEET_NAME,
    input_folder_name=None,
    output_folder_name=None
):

    print("🚀 Starting Alumni KPI Processing...")
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    # 1. Download and Load Data
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    fh_alumni, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    df_raw = load_raw(fh_alumni, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # 2. Parsing and Merging
    df_raw_parsed = df_raw.copy()
    df_raw_parsed[COL_EXIT_DATE] = pd.to_datetime(df_raw_parsed[COL_EXIT_DATE], errors="coerce")
    df_raw_parsed[COL_CHECKIN] = pd.to_datetime(df_raw_parsed[COL_CHECKIN], errors="coerce")
    df_raw_parsed[COL_WAGE] = df_raw_parsed[COL_WAGE].apply(clean_wage)
    df_raw_parsed[COL_HOUSING] = df_raw_parsed[COL_HOUSING].apply(clean_wage) if COL_HOUSING in df_raw_parsed.columns else 0.0
    df_raw_parsed["Monthly Gross Income"] = df_raw_parsed[COL_WAGE] * MONTHLY_HOURS
    if COL_SOBER in df_raw_parsed.columns:
        df_raw_parsed[COL_SOBER] = df_raw_parsed[COL_SOBER].astype(str).str.strip()

    df_grads = df_raw_parsed[
        df_raw_parsed[COL_PROGRAM].replace("House of Grace", "GV Renew").isin(RENEW_PROGRAMS) &
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS) & 
        df_raw_parsed[COL_EXIT_DATE].notna()
    ].drop_duplicates(subset=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE]).reset_index(drop=True)

    df_merged = df_raw_parsed[
        df_raw_parsed[COL_PROGRAM].replace("House of Grace", "GV Renew").isin(RENEW_PROGRAMS) &
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS)
    ].groupby([COL_NAME, COL_PROGRAM], sort=False).agg(
        **{COL_EXIT_DATE: (COL_EXIT_DATE, "max"),
           COL_WAGE: (COL_WAGE, max_nonzero),
           COL_HOUSING: (COL_HOUSING, max_nonzero) if COL_HOUSING in df_raw_parsed.columns else (COL_WAGE, "first")}
    ).reset_index()
    df_merged["Monthly Gross Income"] = df_merged[COL_WAGE] * MONTHLY_HOURS

    # 3. Analyze Clients
    all_results = []
    for _, row in df_grads.iterrows():
        all_results.append(analyze_client(row[COL_NAME], row[COL_PROGRAM], row[COL_EXIT_DATE], df_raw_parsed, df_merged))
    results_df = pd.DataFrame(all_results)

    # 4. Filter for specific tabs
    # --- Sobriety Tab ---
    df_sobriety = results_df[
        results_df["Graduation Date"].between(constants.SOBRIETY_START, constants.SOBRIETY_END)
    ].copy()
    df_sobriety["Sobriety 1 Year"] = df_sobriety["is_sober"].astype(int)
    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        df_sobriety[k] = (
            df_sobriety["Graduation Date"].between(w_start, w_end) & (df_sobriety["Sobriety 1 Year"] == 1)
        ).astype(int)

    # --- LW & Housing Tabs ---
    df_base_lw_h = results_df[
        results_df["Graduation Date"].between(constants.LW_HOUSING_START, constants.LW_HOUSING_END) &
        (results_df["is_sober"] == True)
    ].copy()

    df_lw = df_base_lw_h[[
        COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City", 
        "Year", "Quarter", "Year Q", "LW Criteria", "Wage (Grad)", "Pays LW? (Grad)",
        "Wage (Recent)", "Pays LW? (Recent)"
    ]].copy()

    df_housing = df_base_lw_h[[
        COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City", 
        "Year", "Quarter", "Year Q", "Housing Under 30% (Grad)", "Housing Under 30% (Recent)"
    ]].copy()

    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        mask = df_base_lw_h["Graduation Date"].between(w_start, w_end).astype(int)
        df_lw[k] = mask
        df_housing[k] = mask

    # --- Detailed Analysis ---
    detailed_df = results_df.sort_values(by=[COL_PROGRAM, COL_NAME]).copy()
    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        detailed_df[k] = detailed_df["Graduation Date"].between(w_start, w_end).astype(int)

    # 5. Write to Google Sheets
    output_file_res = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")

    if output_file_res:
        ss = gc.open_by_key(output_file_res['id'])
    else:
        file_metadata = {'name': output_file, 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [output_folder_id]}
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        ss = gc.open_by_key(new_sheet['id'])

    print(f"Writing tabs to {ss.url}...")
    
    # Clean up and fill nulls
    for d in [df_sobriety, df_lw, df_housing, detailed_df]:
        d.fillna('', inplace=True)

    write_tab(ss, "Raw Data", df_raw)
    write_tab(ss, "Sobriety", df_sobriety)
    write_tab(ss, "Living Wage", df_lw)
    write_tab(ss, "Housing", df_housing)
    write_tab(ss, "Detailed Analysis", detailed_df)

    # Cleanup Sheet1
    try:
        ss.del_worksheet(ss.worksheet("Sheet1"))
    except:
        pass

    # Star the file
    drive_service.files().update(
        fileId=ss.id,
        body={"starred": True},
        supportsAllDrives=True
    ).execute()

    print("✅ Done!")
    return f"Alumni Processing Complete. Output: {ss.url}"

if __name__ == "__main__":
    run_alum_processing(
         input_folder_name="Apricot Report Incoming",
         output_folder_name="KPI Processed Data"
    )