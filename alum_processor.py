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

GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]
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
        df_raw_parsed[COL_NAME] == name
    ].sort_values(by=COL_CHECKIN)

    # --- Sobriety Logic ---
    sober_records = person_records[
        person_records[COL_SOBER].isin(["Yes", "No"]) &
        (person_records[COL_CHECKIN] >= grad_date)
    ] if COL_SOBER in person_records.columns else pd.DataFrame()
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
    
    if sober_records.empty:
        is_sober = True  # No post-grad sobriety records — no evidence of relapse, assume sober
    else:
        is_sober = last_checkin_sober and not has_sustained_relapse

    # --- Wage & Housing Logic ---
    merged_row = df_merged[
        (df_merged[COL_NAME] == name) &
        (df_merged[COL_PROGRAM] == program) &
        (df_merged[COL_EXIT_DATE] == grad_date)
    ]

    # Grad wage — max across all rows with this graduation date
    grad_wage = float(merged_row[COL_WAGE].apply(clean_wage).replace(0, float("nan")).max()) if not merged_row.empty else 0.0
    grad_wage = 0.0 if pd.isna(grad_wage) else grad_wage

    # Recent wage — from most recent post-grad checkin; fallback to grad wage if no post-grad data
    _recent_wage_val = merged_row["recent_wage"].iloc[0] if not merged_row.empty and "recent_wage" in merged_row.columns else None
    recent_wage = float(_recent_wage_val) if _recent_wage_val is not None and pd.notna(_recent_wage_val) and float(_recent_wage_val) > 0 else grad_wage
    _recent_checkin_val = merged_row["recent_checkin"].iloc[0] if not merged_row.empty and "recent_checkin" in merged_row.columns else None
    recent_checkin = _recent_checkin_val if _recent_checkin_val is not None and pd.notna(_recent_checkin_val) else pd.NaT

    # Housing Calculation
    grad_under_30 = None
    recent_under_30 = None
    if not merged_row.empty:
        # Grad housing — max across all rows with this graduation date
        grad_housing = float(merged_row[COL_HOUSING].apply(clean_wage).replace(0, float("nan")).max()) if COL_HOUSING in merged_row.columns else 0.0
        grad_housing = 0.0 if pd.isna(grad_housing) else grad_housing
        grad_income = grad_wage * MONTHLY_HOURS
        grad_under_30 = 1 if (grad_housing == 0) else (1 if (grad_income > 0 and (grad_housing / grad_income) < 0.30) else 0)

        # Recent housing — from most recent post-grad checkin; fallback to grad housing
        _recent_housing_val = merged_row["recent_housing"].iloc[0] if "recent_housing" in merged_row.columns else None
        recent_housing = float(_recent_housing_val) if _recent_housing_val is not None and pd.notna(_recent_housing_val) and float(_recent_housing_val) > 0 else grad_housing
        recent_income = recent_wage * MONTHLY_HOURS
        recent_under_30 = 1 if (recent_housing == 0) else (1 if (recent_income > 0 and (recent_housing / recent_income) < 0.30) else 0)
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

    # Remap House of Grace once
    df_raw_parsed[COL_PROGRAM] = df_raw_parsed[COL_PROGRAM].replace("House of Grace", "GV Renew")

    df_grads = df_raw_parsed[
        df_raw_parsed[COL_PROGRAM].isin(RENEW_PROGRAMS) &
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS) &
        df_raw_parsed[COL_EXIT_DATE].notna()
    ].drop_duplicates(subset=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE]).reset_index(drop=True)

    # df_merged: max wage/housing across all rows sharing the same graduation date
    _grad_rows = df_raw_parsed[
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS) &
        df_raw_parsed[COL_PROGRAM].isin(RENEW_PROGRAMS) &
        df_raw_parsed[COL_EXIT_DATE].notna()
    ].copy()
    _grad_rows[COL_WAGE] = _grad_rows[COL_WAGE].apply(clean_wage)
    _grad_rows[COL_HOUSING] = _grad_rows[COL_HOUSING].apply(clean_wage) if COL_HOUSING in _grad_rows.columns else 0.0

    df_merged = _grad_rows.groupby([COL_NAME, COL_PROGRAM, COL_EXIT_DATE], sort=False).agg(
        **{COL_WAGE: (COL_WAGE, max_nonzero),
           COL_HOUSING: (COL_HOUSING, max_nonzero) if COL_HOUSING in _grad_rows.columns else (COL_WAGE, "first")}
    ).reset_index()
    df_merged["Monthly Gross Income"] = df_merged[COL_WAGE] * MONTHLY_HOURS

    # df_post_grad: all rows after grad_date per client+program (post-graduation checkins)
    # Used for Wage (Recent) and Housing (Recent)
    _all = df_raw_parsed.copy()  # Look across all programs for post-grad checkins

    # For each client+program+grad_date, find rows where checkin date > grad_date
    # We merge to get grad dates, then filter
    _grad_dates = df_grads[[COL_NAME, COL_PROGRAM, COL_EXIT_DATE]].rename(columns={COL_EXIT_DATE: "grad_date"})
    _all_with_grad = _all.merge(_grad_dates, on=[COL_NAME, COL_PROGRAM], how="inner")
    _post = _all_with_grad[
        _all_with_grad[COL_CHECKIN].notna() &
        (_all_with_grad[COL_CHECKIN] > _all_with_grad["grad_date"])
    ].copy()

    if not _post.empty:
        # Most recent checkin date per client+program+grad_date
        _latest = _post.groupby([COL_NAME, COL_PROGRAM, "grad_date"])[COL_CHECKIN].transform("max")
        _post_latest = _post[_post[COL_CHECKIN] == _latest].copy()
        df_recent = _post_latest.groupby([COL_NAME, COL_PROGRAM, "grad_date"], sort=False).agg(
            recent_wage=(COL_WAGE, max_nonzero),
            recent_housing=(COL_HOUSING, max_nonzero) if COL_HOUSING in _post_latest.columns else (COL_WAGE, max_nonzero),
            recent_checkin=(COL_CHECKIN, "max")
        ).reset_index().rename(columns={"grad_date": COL_EXIT_DATE})
    else:
        df_recent = pd.DataFrame(columns=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE,
                                          "recent_wage", "recent_housing", "recent_checkin"])

    df_merged = df_merged.merge(df_recent, on=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE], how="left")

    # 3. Analyze Clients
    all_results = []
    for _, row in df_grads.iterrows():
        all_results.append(analyze_client(row[COL_NAME], row[COL_PROGRAM], row[COL_EXIT_DATE], df_raw_parsed, df_merged))
    results_df = pd.DataFrame(all_results)

    # 4. Build Tabs
    # --- Sobriety Tab ---
    df_sobriety = results_df[
        results_df["Graduation Date"].notna() &
        results_df["Graduation Date"].between(constants.SOBRIETY_START, constants.SOBRIETY_END)
    ].copy()
    df_sobriety["Sobriety 1 Year"] = df_sobriety["is_sober"].astype(int)
    df_sobriety = df_sobriety[[
        COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin",
        "City", "Year", "Quarter", "Year Q", "Sustained Relapse?", "Sobriety 1 Year"
    ]].copy()

    # --- LW & Housing Tabs ---
    df_base_lw_h = results_df[
        results_df["Graduation Date"].notna() &
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

    # --- Detailed Analysis ---
    detailed_df = results_df.sort_values(by=[COL_PROGRAM, COL_NAME]).copy()

    # 5. Write to Google Sheets
    output_file_res = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")

    if output_file_res:
        ss = gc.open_by_key(output_file_res['id'])
    else:
        file_metadata = {'name': output_file, 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [output_folder_id]}
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        ss = gc.open_by_key(new_sheet['id'])

    print(f"Writing tabs to {ss.url}...")

    # Fill nulls — numeric cols with 0, rest with empty string
    for d in [df_sobriety, df_lw, df_housing, detailed_df]:
        num_cols = d.select_dtypes(include='number').columns
        d[num_cols] = d[num_cols].fillna(0)
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