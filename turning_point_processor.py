#!/usr/bin/env python3
"""
This script processes the Turning Point Report and outputs key KPI tabs,
including Weekly (FY26+), Monthly, and Quarterly average tabs for Graduates and Successfully Housed.
Includes complete zero-filling across all programs up to the current date.
"""

import time
import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Map programs to high-level categories
TP_PROGRAM_CATEGORY = {
    "Chester Men Turning Point": "Men Turning Point",
    "Oakland Men Turning Point": "Men Turning Point",
    "San Jose Men Turning Point": "Men Turning Point",
    "Portland Men Turning Point": "Men Turning Point",
    "Chester Women Turning Point": "Women Turning Point",
    "GV Turning Point": "Women Turning Point",
    "Heritage Home": "Women Turning Point",
    "Oakland Women Turning Point": "Women Turning Point",
    "Portland Community of Hope": "Women Turning Point",
    "San Jose Youth Collective": "Women Turning Point",
    "Oakland Youth Collective": "Women Turning Point",
    "Portland Youth Collective": "Women Turning Point",
    "Chester Turning Point": "Turning Point",
}

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
    COL_HOUSED_OLD     = "Successfully Housed (Is Housing Healthy?)_4368"
    COL_HOUSED_NEW     = "Successfully Housed_7011"
    COL_HOUSING_CHECK  = "Date of Housing Check_5141"
    COL_INTERN_PROGRAM = "Intern Program_6619"

    PROGRAMS_TO_INCLUDE = [
        "Heritage Home", "Oakland Women Turning Point", "San Jose Men Turning Point",
        "Chester Men Turning Point", "Chester Women Turning Point", "GV Turning Point",
        "Oakland Men Turning Point", "Oakland Youth Collective", "Portland Community of Hope",
        "Portland Youth Collective", "San Jose Youth Collective", "Chester Turning Point",
    ]

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
    if COL_HOUSING_CHECK in df_base.columns:
        df_base[COL_HOUSING_CHECK] = pd.to_datetime(df_base[COL_HOUSING_CHECK], errors="coerce")

    # Remap Program Graduate Interns to their specific Intern Program
    if COL_INTERN_PROGRAM in df_base.columns:
        is_intern = df_base[COL_PROGRAM] == "Program Graduate Intern"
        df_base.loc[is_intern, COL_PROGRAM] = df_base.loc[is_intern, COL_INTERN_PROGRAM].apply(match_program)

    df_base = df_base[df_base[COL_PROGRAM].isin(PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

    # Helper function to get Monday start of week
    def get_week_start(dt_series):
        dt = pd.to_datetime(dt_series, errors="coerce")
        return dt.apply(lambda d: (d - pd.Timedelta(days=d.weekday())).strftime("%Y-%m-%d") if pd.notna(d) else "")

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

        df["City"]         = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]         = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]      = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]       = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        df["Week"]         = get_week_start(df[COL_EXIT_DATE])
        df["Next FY Goal"] = df[COL_PROGRAM].map(constants.NEXT_FY_GRAD_GOALS_TP)

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= w_start) & (df[COL_EXIT_DATE] <= w_end)).astype(int)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q", "Week", "Next FY Goal"] + list(constants.ACTUALS_WINDOWS.keys())
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

        if COL_CHILD in df.columns:
            is_gv_tp_null = (df[COL_PROGRAM] == "GV Turning Point") & df["Name of Child_6313"].isna()
            if is_gv_tp_null.any():
                child_lookup = df_base[df_base[COL_PROGRAM] == "GV Turning Point"].dropna(subset=["Name of Child_6313"])
                child_lookup = child_lookup.groupby(COL_RECORD_ID)["Name of Child_6313"].first()
                df.loc[is_gv_tp_null, "Name of Child_6313"] = df.loc[is_gv_tp_null, COL_RECORD_ID].map(child_lookup)

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

        df[constants.OCC_CURRENT_LABEL] = pd.to_numeric(df[constants.OCC_CURRENT_LABEL], errors="coerce").fillna(0).astype(int)

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
        df = df_raw.copy()
        df[COL_PROGRAM]    = df[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
        df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
        df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE],  errors="coerce")
        if COL_HOUSING_CHECK in df.columns:
            df[COL_HOUSING_CHECK] = pd.to_datetime(df[COL_HOUSING_CHECK], errors="coerce")

        if COL_INTERN_PROGRAM in df.columns:
            is_intern = df[COL_PROGRAM] == "Program Graduate Intern"
            df.loc[is_intern, COL_PROGRAM] = df.loc[is_intern, COL_INTERN_PROGRAM].apply(match_program)

        df = df[df[COL_PROGRAM].isin(PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

        def eval_housed(row):
            val_n = str(row.get(COL_HOUSED_NEW, '')).strip().lower() if COL_HOUSED_NEW in row else ''
            if val_n not in ['', 'nan', 'none', 'nat']:
                return 1 if 'yes' in val_n else 0
            
            val_o = str(row.get(COL_HOUSED_OLD, '')).strip().lower() if COL_HOUSED_OLD in row else ''
            if val_o not in ['', 'nan', 'none', 'nat']:
                return 1 if 'yes' in val_o else 0
                
            return 0

        df["Successfully Housed?"] = df.apply(eval_housed, axis=1)

        actuals_cols = list(constants.ACTUALS_WINDOWS.keys())
        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (
                df[COL_EXIT_DATE].notna() &
                (df[COL_EXIT_DATE] >= w_start) &
                (df[COL_EXIT_DATE] <= w_end) &
                (df["Successfully Housed?"] == 1)
            ).astype(int)

        id_actuals = df.groupby([COL_RECORD_ID, COL_PROGRAM])[actuals_cols].max()

        df = (df.sort_values(["Successfully Housed?", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM], keep="first"))

        df = df.drop(columns=actuals_cols).merge(id_actuals, on=[COL_RECORD_ID, COL_PROGRAM], how="left")

        df["City"]     = df[COL_PROGRAM].apply(constants.assign_city)
        df["Year"]     = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]  = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]   = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        
        # Determine Week based on Date of Housing Check (fallback to Exit Date)
        housing_check_or_exit = df[COL_HOUSING_CHECK].fillna(df[COL_EXIT_DATE]) if COL_HOUSING_CHECK in df.columns else df[COL_EXIT_DATE]
        df["Week"]     = get_week_start(housing_check_or_exit)

        df["Capacity"] = df[COL_PROGRAM].map(constants.TP_CAPACITY)

        housed_goals_map = getattr(constants, "NEXT_FY_HOUSED_GOALS_TP", {})
        df["Goal"] = df[COL_PROGRAM].map(housed_goals_map).fillna(0).astype(int)

        cols = [COL_RECORD_ID, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_HOUSING_CHECK, COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q", "Week", "Successfully Housed?", "Capacity", "Goal"
                ] + actuals_cols

        df = df[[c for c in cols if c in df.columns]].copy()
        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
        if COL_HOUSING_CHECK in df.columns:
            df[COL_HOUSING_CHECK] = fmt_date(df, COL_HOUSING_CHECK)
        for k in actuals_cols:
            if k in df.columns: df[k] = df[k].fillna(0).astype(int)
            
        return df.reset_index(drop=True)

    # =========================================================================
    # Helper Functions for Averages Processing
    # =========================================================================
    def get_fy(date):
        return date.year + 1 if date.month >= 9 else date.year

    def get_fy_label(date):
        fy = get_fy(date)
        return f"FY{str(fy)[-2:]}"

    def get_week_num(date):
        fy = get_fy(date)
        fy_start = pd.Timestamp(f"{fy-1}-09-01")
        delta = (date - fy_start).days
        return min(52, max(1, delta // 7 + 1))

    def get_week_label(date):
        return f"Week {get_week_num(date)}"

    def get_fy_week_label(date):
        return f"{get_fy_label(date)} Week {get_week_num(date)}"

    def get_month(date):
        return date.strftime("%B")

    def get_fy_month(date):
        return date.strftime("%B %Y")

    def get_quarter(date):
        m = date.month
        if m in [9, 10, 11]: return "Q1"
        elif m in [12, 1, 2]: return "Q2"
        elif m in [3, 4, 5]: return "Q3"
        else: return "Q4"

    def get_fy_quarter(date):
        return f"{get_fy_label(date)} {get_quarter(date)}"

    def compute_kpi_averages(df_source, date_col, period_fn, fy_period_fn=None, is_weekly=False, filter_flag_col=None, min_fy=None):
        if df_source.empty:
            return pd.DataFrame()

        df = df_source.copy()

        # Dynamic date resolution with fallback to COL_EXIT_DATE
        if date_col in df.columns:
            df["_Date"] = pd.to_datetime(df[date_col], errors="coerce")
            if COL_EXIT_DATE in df.columns:
                df["_Date"] = df["_Date"].fillna(pd.to_datetime(df[COL_EXIT_DATE], errors="coerce"))
        elif COL_EXIT_DATE in df.columns:
            df["_Date"] = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")
        else:
            return pd.DataFrame()

        df = df.dropna(subset=["_Date"])

        if filter_flag_col and filter_flag_col in df.columns:
            df = df[pd.to_numeric(df[filter_flag_col], errors="coerce") == 1]

        if df.empty:
            return pd.DataFrame()

        # Build timeline up to current date
        min_date = df["_Date"].min()
        max_date = pd.Timestamp.now()
        
        timeline = pd.date_range(start=min_date, end=max_date, freq='D')
        df_timeline = pd.DataFrame({"_Date": timeline})
        df_timeline["_period"] = df_timeline["_Date"].apply(period_fn)
        df_timeline["_fy_period"] = df_timeline["_Date"].apply(fy_period_fn) if fy_period_fn else df_timeline["_period"]
        df_timeline["_fy"] = df_timeline["_Date"].apply(get_fy_label)

        if is_weekly:
            df_timeline["_week_date"] = df_timeline["_Date"].apply(lambda d: (d - pd.Timedelta(days=d.weekday())).strftime("%Y-%m-%d"))
        else:
            df_timeline["_week_date"] = df_timeline["_Date"].dt.strftime("%Y-%m-%d")

        if min_fy:
            df_timeline = df_timeline[df_timeline["_fy"] >= min_fy]

        if df_timeline.empty:
            return pd.DataFrame()

        unique_periods = df_timeline.groupby("_fy_period", sort=False).first().reset_index()

        # Map actual counts using compound _fy_period key
        df["_fy_period"] = df["_Date"].apply(fy_period_fn) if fy_period_fn else df["_Date"].apply(period_fn)
        counts = df.groupby(["_fy_period", COL_PROGRAM]).size().to_dict()

        rows = []
        for _, p_row in unique_periods.iterrows():
            fy_period = p_row["_fy_period"]
            period = p_row["_period"]
            fy = p_row["_fy"]
            date_val = p_row["_week_date"] if is_weekly else p_row["_Date"].strftime("%Y-%m-%d")

            for prog in PROGRAMS_TO_INCLUDE:
                cat_val = TP_PROGRAM_CATEGORY.get(prog, "Turning Point")
                city_val = constants.assign_city(prog)
                
                total_count = float(counts.get((fy_period, prog), 0))

                rows.append({
                    "FY": fy,
                    "Period": period,
                    "FY Period": fy_period,
                    "Date": date_val,
                    "Category": cat_val,
                    "City": city_val,
                    "Program": prog,
                    "Average": total_count
                })

        return pd.DataFrame(rows)

    def write_averages_tab(out_ss, tab_name, df_avg, period_col_name):
        if df_avg.empty:
            return
        
        df_out = df_avg.rename(columns={
            "Period": period_col_name,
            "FY Period": f"FY {period_col_name}"
        })

        cols = ["FY", period_col_name, f"FY {period_col_name}", "Date", "Category", "City", "Program", "Average"]
        df_out = df_out[[c for c in cols if c in df_out.columns]]

        write_tab(out_ss, tab_name, df_out)
        print(f"✅ {tab_name} written")

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
        print(f"📄 Opened existing spreadsheet: '{output_file}'")
    else:
        file_metadata = {
            'name': output_file,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [output_folder_id]
        }
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        spreadsheet = gc.open_by_key(new_sheet['id'])
        print(f"📄 Created new spreadsheet: '{output_file}'")

    # =========================================================================
    # 6. Export Main Tabs
    # =========================================================================
    write_tab(spreadsheet, "Raw Data",                 df_raw)
    write_tab(spreadsheet, "Turning Point Graduates",  df_grad)
    write_tab(spreadsheet, "Turning Point Occupancy",  df_occ)
    write_tab(spreadsheet, "Successfully Housed",      df_housed)

    # =========================================================================
    # 7. Compute & Export KPI Average Tabs
    # =========================================================================
    print("\nComputing Graduates Averages...")
    write_averages_tab(spreadsheet, "Graduates Weekly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_week_label, get_fy_week_label, is_weekly=True, min_fy="FY26"), "Week")
    write_averages_tab(spreadsheet, "Graduates Monthly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_month, get_fy_month), "Month")
    write_averages_tab(spreadsheet, "Graduates Quarterly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_quarter, get_fy_quarter), "Quarter")

    print("\nComputing Successfully Housed Averages...")
    housed_date_col = COL_HOUSING_CHECK if COL_HOUSING_CHECK in df_housed.columns else COL_EXIT_DATE
    write_averages_tab(spreadsheet, "Successfully Housed Weekly Avg", 
                       compute_kpi_averages(df_housed, housed_date_col, get_week_label, get_fy_week_label, is_weekly=True, filter_flag_col="Successfully Housed?", min_fy="FY26"), "Week")
    write_averages_tab(spreadsheet, "Successfully Housed Monthly Avg", 
                       compute_kpi_averages(df_housed, housed_date_col, get_month, get_fy_month, filter_flag_col="Successfully Housed?"), "Month")
    write_averages_tab(spreadsheet, "Successfully Housed Quarterly Avg", 
                       compute_kpi_averages(df_housed, housed_date_col, get_quarter, get_fy_quarter, filter_flag_col="Successfully Housed?"), "Quarter")

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Done! View sheet at: {spreadsheet.url}")


if __name__ == "__main__":
    run_turning_point_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )