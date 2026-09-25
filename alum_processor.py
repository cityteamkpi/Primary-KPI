import warnings
import os
import time
import math
import datetime
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import (
    resolve_folder_id, 
    download_drive_file, 
    load_raw, 
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
COL_START_DATE = "Start Date_2090"
COL_EXIT_DATE = "Exit Date_2100"
COL_EXIT_REASON = "Primary Reason for Exit_2102"
COL_CHECKIN = "Checkin Date_6583"
COL_WAGE = "Hourly Wage_6586"
COL_HOUSING = "Housing Expense_6587"
COL_SOBER = "Sober?_6585"
COL_HOW_CHECKED = "How Checked_6584"

GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]
MONTHLY_HOURS = 173.33

INTERN_PROGRAM = "Program Graduate Intern"

# Hardcoded Intern List
INTERNS = {
    "John Au",
    "Anthony Brumley",
    "Christopher Cajero",
    "Shawn Cannon Jr.",
    "Jose Carpio",
    "Isidro Castro Jr.",
    "Kayla Chamberlain",
    "Evans Choi",
    "Jesus Contreras",
    "Shea Davies",
    "Henry Davis",
    "Tommy Khoa Dinh",
    "James Donaldson",
    "April Dorlarque",
    "Kathleen Erfurt",
    "Haley Givens",
    "Pedro Gomez III",
    "Dwayne Gooden",
    "Erik Grier",
    "Jueun Gu",
    "JASON Herr",
    "Dennis Hester",
    "Christopher Hill",
    "Linden Hyman Jr.",
    "Robert Ivens",
    "James Kahoe",
    "Kyllie Lawson",
    "Lawrence LeDee",
    "Baxter Lipscomb",
    "Richard Lira Jr.",
    "Buck Loden",
    "James Lopez",
    "Gabriel Maldonado",
    "Catherine Mayshack",
    "Andrew Murphy",
    "Denny Nguyen",
    "Alexandria Nunez",
    "Johnne Shiheiber",
    "Veronica Silva Martinez",
    "Mario Terrigno",
    "Angelica Tribo",
    "Jose Varguez",
    "Joshua West",
}

def clean_wage(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    clean = str(val).replace("$", "").replace(",", "").strip()
    try: return float(clean)
    except: return 0.0

def max_nonzero(series):
    vals = series.replace(0, np.nan).dropna()
    return vals.max() if not vals.empty else 0.0

def get_grad_wage(group):
    """Prioritizes graduation checkin rows where How Checked is 'Graduation' strictly.
    Returns the wage on that row (even if $0.00). Fallback to max_nonzero only if
    no 'Graduation' checkin row exists at all.
    """
    if COL_HOW_CHECKED in group.columns:
        grad_checked = group[group[COL_HOW_CHECKED] == "Graduation"]
        if not grad_checked.empty:
            return float(grad_checked[COL_WAGE].iloc[0])
    
    # Fallback if no explicit 'Graduation' check-in row exists
    return max_nonzero(group[COL_WAGE])

def write_tab(spreadsheet, tab_name, df):
    """Writes DataFrame to Google Sheets replacing all non-compliant JSON values
    (NaN, NaT, Inf, Timestamps, Python datetimes) with standard Python strings or None.
    """
    needed_rows = max(len(df) + 1, 1)
    needed_cols = max(len(df.columns), 1)

    try:
        worksheet = spreadsheet.worksheet(tab_name)
        if worksheet.row_count < needed_rows or worksheet.col_count < needed_cols:
            worksheet.resize(rows=max(needed_rows, worksheet.row_count),
                              cols=max(needed_cols, worksheet.col_count))
        worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=tab_name, rows=needed_rows, cols=needed_cols)

    if df.empty:
        worksheet.update([df.columns.tolist()])
        return

    df_clean = df.fillna("")
    records = df_clean.to_dict(orient="records")
    clean_rows = []

    for row in records:
        clean_row = []
        for val in row.values():
            if pd.isna(val) or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
                clean_row.append("")
            elif isinstance(val, (pd.Timestamp, pd.DatetimeIndex, datetime.datetime, datetime.date)):
                clean_row.append(str(val)[:10])
            elif isinstance(val, str) and val.lower() in ("nan", "nat", "inf", "-inf"):
                clean_row.append("")
            else:
                clean_row.append(val)
        clean_rows.append(clean_row)

    headers = [str(c) if c is not None else "" for c in df.columns.tolist()]
    values = [headers] + clean_rows
    worksheet.update(values)

def analyze_client(name, program, grad_date, df_raw_parsed, df_merged):
    lw_target = constants.LW_CRITERIA.get(program, 0.0) or 0.0
    
    person_records = df_raw_parsed[
        (df_raw_parsed[COL_NAME] == name) &
        (df_raw_parsed[COL_PROGRAM] == program)
    ].sort_values(by=COL_CHECKIN)

    if COL_START_DATE in person_records.columns:
        person_records = person_records[
            (person_records[COL_START_DATE].isna()) |
            (person_records[COL_START_DATE] <= grad_date) |
            (person_records[COL_CHECKIN] < person_records[COL_START_DATE])
        ]

    # --- Sobriety Logic ---
    sober_records = person_records[
        person_records[COL_SOBER].isin(["Yes", "No"]) &
        (person_records[COL_CHECKIN] > grad_date)
    ] if COL_SOBER in person_records.columns else pd.DataFrame()
    
    has_sustained_relapse = False
    first_no_date = None
    last_checkin_sober = False
    
    for _, s_row in sober_records.iterrows():
        cur_state = s_row[COL_SOBER]
        cur_date = s_row[COL_CHECKIN]
        
        if cur_state == "No":
            if first_no_date is None: 
                first_no_date = cur_date
        elif cur_state == "Yes":
            if first_no_date is not None:
                if (cur_date - first_no_date).days > 30: 
                    has_sustained_relapse = True
                first_no_date = None
                
        last_checkin_sober = (cur_state == "Yes")
    
    if sober_records.empty:
        is_sober = True
    else:
        is_sober = last_checkin_sober and not has_sustained_relapse

    # --- Wage & Housing Logic ---
    merged_row = df_merged[
        (df_merged[COL_NAME] == name) &
        (df_merged[COL_PROGRAM] == program) &
        (df_merged[COL_EXIT_DATE] == grad_date)
    ]

    _grad_wage_val = merged_row[COL_WAGE].iloc[0] if not merged_row.empty else 0.0
    grad_wage = float(_grad_wage_val) if pd.notna(_grad_wage_val) else 0.0

    _recent_wage_val = merged_row["recent_wage"].iloc[0] if not merged_row.empty and "recent_wage" in merged_row.columns else None
    recent_wage = float(_recent_wage_val) if _recent_wage_val is not None and pd.notna(_recent_wage_val) else grad_wage
    _recent_checkin_val = merged_row["recent_checkin"].iloc[0] if not merged_row.empty and "recent_checkin" in merged_row.columns else None
    recent_checkin = _recent_checkin_val if _recent_checkin_val is not None and pd.notna(_recent_checkin_val) else pd.NaT

    wage_recent_out = recent_wage if is_sober else None

    # Housing Expense & Income calculations
    grad_housing = float(merged_row[COL_HOUSING].iloc[0]) if not merged_row.empty and COL_HOUSING in merged_row.columns and pd.notna(merged_row[COL_HOUSING].iloc[0]) else 0.0
    grad_income = grad_wage * MONTHLY_HOURS

    grad_under_30 = None
    recent_under_30 = None
    recent_housing = grad_housing
    recent_income = 0.0

    # Helper function for housing affordability calculation (<= 30%)
    def is_under_30(housing, income):
        if housing == 0:
            return 1
        if income > 0:
            return 1 if (housing / income) <= 0.30 else 0
        return 0  # Income is 0 but Housing > 0

    if not merged_row.empty:
        grad_under_30 = is_under_30(grad_housing, grad_income)

        _recent_housing_val = merged_row["recent_housing"].iloc[0] if "recent_housing" in merged_row.columns else None
        recent_housing = float(_recent_housing_val) if _recent_housing_val is not None and pd.notna(_recent_housing_val) else grad_housing
        
        effective_recent_wage = wage_recent_out if wage_recent_out is not None else 0.0
        recent_income = effective_recent_wage * MONTHLY_HOURS
        recent_under_30 = is_under_30(recent_housing, recent_income)
    
    # --- Living Wage Evaluators ---
    pays_lw_grad = int(grad_wage >= lw_target) if lw_target else 0

    if is_sober:
        pays_lw_recent = int(recent_wage >= lw_target) if lw_target else 0
    else:
        pays_lw_recent = "Not sober"

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
        "Pays LW? (Grad)": pays_lw_grad,
        "Wage (Recent)": wage_recent_out,
        "Pays LW? (Recent)": pays_lw_recent,
        "Monthly Income (Grad)": grad_income,
        "Monthly Income (Recent)": recent_income if is_sober else None,
        "Housing Expense (Grad)": grad_housing,
        "Housing Expense (Recent)": recent_housing,
        "Housing Under 30% (Grad)": 0 if grad_under_30 is None else int(grad_under_30),
        "Housing Under 30% (Recent)": 0 if recent_under_30 is None else int(recent_under_30),
    }

def run_alum_processing(
    input_file=SRC_ALUMNI + ".xlsx",
    output_file=OUTPUT_SHEET_NAME,
    input_folder_name=None,
    output_folder_name=None
):
    print("🚀 Starting Alumni KPI Processing (Strict Renew Graduates)...")
    drive_service, gc, _ = get_services()
    constants.sync_constants()

    # ── Fiscal Year Target Filtering ──────────────────────────────────────────
    target_fy_num = constants.fy_num - 1
    target_fy_str = f"FY{str(target_fy_num)[-2:]}"

    print(f"   KPI Target Fiscal Year: {target_fy_str} (Q1 - Q4)")

    # 1. Download and Load Data
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    fh_alumni, _, _ = download_drive_file(drive_service, input_file, input_folder_id)

    # --- RAW DATA & PARSED DATA ALIGNMENT ---
    fh_alumni.seek(0)
    df_raw_complete = load_raw(fh_alumni, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    df_raw = df_raw_complete.copy()

    # 2. Parsing & Dropping Invalid Client Names
    df_raw_parsed = df_raw.dropna(subset=[COL_NAME]).copy()
    df_raw_parsed[COL_EXIT_DATE] = pd.to_datetime(df_raw_parsed[COL_EXIT_DATE], errors="coerce")
    df_raw_parsed[COL_CHECKIN] = pd.to_datetime(df_raw_parsed[COL_CHECKIN], errors="coerce")
    if COL_START_DATE in df_raw_parsed.columns:
        df_raw_parsed[COL_START_DATE] = pd.to_datetime(df_raw_parsed[COL_START_DATE], errors="coerce")
    df_raw_parsed[COL_WAGE] = df_raw_parsed[COL_WAGE].apply(clean_wage)
    df_raw_parsed[COL_HOUSING] = df_raw_parsed[COL_HOUSING].apply(clean_wage) if COL_HOUSING in df_raw_parsed.columns else 0.0
    df_raw_parsed["Monthly Gross Income"] = df_raw_parsed[COL_WAGE] * MONTHLY_HOURS
    if COL_SOBER in df_raw_parsed.columns:
        df_raw_parsed[COL_SOBER] = df_raw_parsed[COL_SOBER].astype(str).str.strip()

    # Map each client's absolute latest enrolled program across the entire raw dataset
    df_raw_sorted = df_raw_parsed.sort_values(by=[COL_EXIT_DATE, COL_CHECKIN])
    latest_program_map = (
        df_raw_sorted.dropna(subset=[COL_PROGRAM])
        .groupby(COL_NAME)[COL_PROGRAM]
        .last()
        .astype(str).str.strip()
        .to_dict()
    )

    # 3. Pull graduation records EXCLUSIVELY for Renew Graduates
    renew_programs_set = {str(p).strip().lower() for p in constants.RENEW_PROGRAMS}
    
    df_grads_all = df_raw_parsed[
        df_raw_parsed[COL_PROGRAM].astype(str).str.strip().str.lower().isin(renew_programs_set) &
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS) &
        df_raw_parsed[COL_EXIT_DATE].notna()
    ].copy()

    # Deduplicate: Preserve distinct graduation events per Client, Program, and Exit Date
    df_grads = (
        df_grads_all
        .sort_values(by=[COL_EXIT_DATE, COL_EXIT_REASON])
        .drop_duplicates(subset=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE], keep="last")
        .reset_index(drop=True)
    )

    # Calculate prioritized wage & housing for graduation records
    _grad_rows = df_raw_parsed[
        df_raw_parsed[COL_EXIT_REASON].isin(GRAD_REASONS) &
        df_raw_parsed[COL_EXIT_DATE].notna()
    ].copy()
    _grad_rows[COL_EXIT_DATE] = pd.to_datetime(_grad_rows[COL_EXIT_DATE]).dt.normalize()

    def agg_grad_group(g):
        w_val = get_grad_wage(g)
        h_val = max_nonzero(g[COL_HOUSING]) if COL_HOUSING in g.columns else 0.0
        return pd.Series({COL_WAGE: w_val, COL_HOUSING: h_val})

    df_merged = (
        _grad_rows
        .groupby([COL_NAME, COL_PROGRAM, COL_EXIT_DATE], sort=False)
        .apply(agg_grad_group)
        .reset_index()
    )

    # Post-grad checkin processing
    _all = df_raw_parsed.copy()
    _grad_dates = df_grads[[COL_NAME, COL_PROGRAM, COL_EXIT_DATE]].rename(columns={COL_EXIT_DATE: "grad_date"})
    _grad_dates["grad_date"] = pd.to_datetime(_grad_dates["grad_date"]).dt.normalize()
    _all_with_grad = _all.merge(_grad_dates, on=[COL_NAME, COL_PROGRAM], how="inner")
    _post = _all_with_grad[
        _all_with_grad[COL_CHECKIN].notna() &
        (_all_with_grad[COL_CHECKIN] > _all_with_grad["grad_date"])
    ].copy()

    if COL_START_DATE in _post.columns:
        _post = _post[
            (_post[COL_START_DATE].isna()) | 
            (_post[COL_START_DATE] <= _post["grad_date"]) | 
            (_post[COL_CHECKIN] < _post[COL_START_DATE])
        ].copy()

    if not _post.empty:
        _latest = _post.groupby([COL_NAME, COL_PROGRAM, "grad_date"])[COL_CHECKIN].transform("max")
        _post_latest = _post[_post[COL_CHECKIN] == _latest].copy()
        
        # Grab wage and housing strictly from the latest checkin date row(s)
        df_recent = _post_latest.groupby([COL_NAME, COL_PROGRAM, "grad_date"], sort=False).agg(
            recent_wage=(COL_WAGE, max_nonzero),
            recent_housing=(COL_HOUSING, "max"),  # Max as tie-breaker if multiple rows on latest date
            recent_checkin=(COL_CHECKIN, "max")
        ).reset_index().rename(columns={"grad_date": COL_EXIT_DATE})
    else:
        df_recent = pd.DataFrame(columns=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE,
                                          "recent_wage", "recent_housing", "recent_checkin"])

    df_merged = df_merged.merge(df_recent, on=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE], how="left")

    # 4. Analyze Renew Graduates
    all_results = []
    for _, row in df_grads.iterrows():
        all_results.append(analyze_client(row[COL_NAME], row[COL_PROGRAM], row[COL_EXIT_DATE], df_raw_parsed, df_merged))
    
    results_df = pd.DataFrame(all_results) if all_results else pd.DataFrame()

    if not results_df.empty:
        results_df["Latest Program"] = results_df[COL_NAME].map(latest_program_map)
    else:
        results_df["Latest Program"] = None

    fy_filtered_df = results_df[results_df["Year"] == target_fy_str].copy() if not results_df.empty else pd.DataFrame()

    # --- Sobriety Tab ---
    df_sobriety = fy_filtered_df.copy()
    if not df_sobriety.empty:
        df_sobriety["Sobriety 1 Year"] = df_sobriety["is_sober"].astype(int)
        df_sobriety = df_sobriety[[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin",
            "City", "Year", "Quarter", "Year Q", "Sustained Relapse?", "Sobriety 1 Year"
        ]].copy()
    else:
        df_sobriety = pd.DataFrame(columns=[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin",
            "City", "Year", "Quarter", "Year Q", "Sustained Relapse?", "Sobriety 1 Year"
        ])

    # --- Living Wage Tab ---
    df_base_lw = fy_filtered_df.copy() if not fy_filtered_df.empty else pd.DataFrame()
    if not df_base_lw.empty:
        df_lw = df_base_lw[[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City",
            "Year", "Quarter", "Year Q", "LW Criteria", "Wage (Grad)", "Pays LW? (Grad)",
            "Wage (Recent)", "Pays LW? (Recent)"
        ]].copy()
    else:
        df_lw = pd.DataFrame(columns=[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City",
            "Year", "Quarter", "Year Q", "LW Criteria", "Wage (Grad)", "Pays LW? (Grad)",
            "Wage (Recent)", "Pays LW? (Recent)"
        ])

    # --- Housing Tab ---
    if not fy_filtered_df.empty:
        df_base_housing = fy_filtered_df[
            (fy_filtered_df["is_sober"] == True) &
            (~fy_filtered_df[COL_NAME].isin(INTERNS))
        ].copy()
    else:
        df_base_housing = pd.DataFrame()

    if not df_base_housing.empty:
        df_housing = df_base_housing[[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City",
            "Year", "Quarter", "Year Q", "Wage (Grad)", "Wage (Recent)", 
            "Monthly Income (Grad)", "Monthly Income (Recent)",
            "Housing Expense (Grad)", "Housing Expense (Recent)",
            "Housing Under 30% (Grad)", "Housing Under 30% (Recent)"
        ]].copy()
    else:
        df_housing = pd.DataFrame(columns=[
            COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin", "City",
            "Year", "Quarter", "Year Q", "Wage (Grad)", "Wage (Recent)", 
            "Monthly Income (Grad)", "Monthly Income (Recent)",
            "Housing Expense (Grad)", "Housing Expense (Recent)",
            "Housing Under 30% (Grad)", "Housing Under 30% (Recent)"
        ])

    # --- Detailed Analysis ---
    detailed_df = results_df.sort_values(by=[COL_PROGRAM, COL_NAME]).copy() if not results_df.empty else results_df

    # 5. Write Output
    output_file_res = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")

    if output_file_res:
        ss = gc.open_by_key(output_file_res['id'])
    else:
        file_metadata = {'name': output_file, 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [output_folder_id]}
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        ss = gc.open_by_key(new_sheet['id'])

    print(f"Writing tabs to {ss.url}...")

    # Write DataFrames to Sheets
    write_tab(ss, "Raw Data", df_raw_complete)
    write_tab(ss, "Sobriety", df_sobriety)
    write_tab(ss, "Living Wage", df_lw)
    write_tab(ss, "Housing", df_housing)
    write_tab(ss, "Detailed Analysis", detailed_df)

    try:
        ss.del_worksheet(ss.worksheet("Sheet1"))
    except Exception:
        pass

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