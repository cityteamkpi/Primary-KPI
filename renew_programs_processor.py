#!/usr/bin/env python3
"""
Renew Programs Processing Pipeline
Processes Graduates, Occupancy, and Retention tabs for Q4 FY26.
Includes Weekly (FY26+), Monthly, and Quarterly Average tabs for Graduates.
Retention cohort covers clients starting between 5/1/2026 (1 month prior) and 8/31/2026.
Includes prior quarter (Q3) retention tracking metrics.
"""

import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id, load_raw, write_tab
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Category Mapping for Renew Programs
RENEW_PROGRAM_CATEGORY = {
    "Chester Renew": "Renew",
    "GV Renew": "Renew",
    "Oakland Renew": "Renew",
    "Portland Renew": "Renew",
    "San Jose Men Renew": "Renew",
    "San Jose Women Renew": "Renew",
}

def run_renew_processing(
    input_file="Renew Programs Report.xlsx",
    module_file="Renew Module Progress.xlsx",
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
    COL_NAME         = "Name_2057"
    COL_PROGRAM      = "Program Enrolling_2091"
    COL_START_DATE   = "Start Date_2090"
    COL_EXIT_DATE    = "Exit Date_2100"
    COL_EXIT_REASON  = "Primary Reason for Exit_2102"
    COL_CHILD        = "Name of Child_6313"
    COL_SJ_CASE_MGR  = "San Jose Men Case Managers_4668"

    COL_CHECKIN_DATE   = "Checkin Date_6583"
    COL_HOW_CHECKED    = "How Checked_6584"
    COL_ENGAGEMENT     = "Status of Member Engagement_6386"
    VALID_HOW_CHECKED  = [
        "email", "in-person", "jail roster", "telephone", "text", "text message",
        "co-worker", "reliable source", "reliable source/alumni source", "alumni source"
    ]

    # Module Mapping across City Locations
    CITY_MODULE_MAP = {
        "Oakland": [
            (1, "Module 1 Begin Date_3365"),
            (2, "Module 2 Begin Date_3898"),
            (3, "Module 3 Begin Date_3529"),
            (4, "Module 4 Begin Date_4244")
        ],
        "Portland": [
            (1, "Module 1 Begin Date_3702"),
            (2, "Module 2 Begin Date_4014"),
            (3, "Module 3 Begin Date_4082"),
            (4, "Module 4 Begin Date_4160")
        ],
        "Chester": [
            (1, "Module 1 Begin Date_4396"),
            (2, "Module 2 Begin Date_4479"),
            (3, "Module 3 Begin Date_4586"),
            (4, "Module 4 Begin Date_7204")
        ],
        "San Jose Men": [
            (1, "Module 1 Begin Date_2604"),
            (2, "Module 2 Begin Date_2670"),
            (3, "Module 3 Begin Date_2768"),
            (4, "Module 4 Begin Date_2834")
        ],
        "San Jose Women": [
            (1, "Module 1 Begin Date_2604"),
            (2, "Module 2 Begin Date_2670"),
            (3, "Module 3 Begin Date_2768"),
            (4, "Module 4 Begin Date_2834")
        ]
    }

    # Extract all module date columns across all cities for clean handling
    ALL_MODULE_DATE_COLS = list({col for mods in CITY_MODULE_MAP.values() for _, col in mods})

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

    # Q3 Boundaries
    q3_start_ts = q_start_ts - pd.DateOffset(months=3)          # 2026-03-01

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download and Load Raw Data & Module Progress Data
    # =========================================================================
    print(f"Downloading {input_file}...")
    fh, _, _ = download_drive_file(drive_service, input_file, input_folder_id)
    df_raw = load_raw(fh, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {module_file}...")
    fh_mod, _, _ = download_drive_file(drive_service, module_file, input_folder_id)
    df_mod_raw = load_raw(fh_mod, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # Merge Module Progress data into Raw Data using Record Id_102
    dup_cols = [c for c in df_mod_raw.columns if c in df_raw.columns and c != COL_RECORD_ID]
    df_mod_clean = df_mod_raw.drop(columns=dup_cols)
    df_raw = df_raw.merge(df_mod_clean, on=COL_RECORD_ID, how="left")

    # Ensure required checkin & indicator columns exist
    for col in ALL_MODULE_DATE_COLS + [COL_CHECKIN_DATE, COL_HOW_CHECKED, COL_ENGAGEMENT, "Acquired Spiritual Mentor_5136", "Acquired Sponsor_5137"]:
        if col not in df_raw.columns:
            df_raw[col] = np.nan

    df_base = df_raw.copy()
    df_base[COL_START_DATE] = pd.to_datetime(df_base[COL_START_DATE], errors="coerce")
    df_base[COL_EXIT_DATE]  = pd.to_datetime(df_base[COL_EXIT_DATE],  errors="coerce")
    df_base[COL_PROGRAM]   = df_base[COL_PROGRAM].replace("House of Grace", "GV Renew")
    df_base = df_base[df_base[COL_PROGRAM].isin(RENEW_PROGRAMS)].reset_index(drop=True)

    # =========================================================================
    # 3. Helper Functions
    # =========================================================================
    def assign_city_custom(program_name):
        if str(program_name).strip() in ["San Jose Women Renew", "House of Grace", "GV Renew"]:
            return "San Jose Women"
        return constants.assign_city(program_name)

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

    def calc_days_in_program(df):
        today = pd.Timestamp.now().normalize()
        start = pd.to_datetime(df[COL_START_DATE], errors="coerce")
        exit_dt = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")
        end = exit_dt.fillna(today)
        
        days = (end - start).dt.days
        return days.where(start.notna(), np.nan)

    def calc_days_time_window(days):
        if pd.isna(days):
            return np.nan
        if days <= 30:
            return "30days"
        elif days <= 60:
            return "31~60days"
        elif days <= 90:
            return "61~90days"
        else:
            return "90~days"

    def calc_spiritual_mentor(row):
        is_active = pd.isna(row.get(COL_EXIT_DATE))
        if is_active:
            val = row.get("Acquired Spiritual Mentor_5136")
            return 1 if str(val).strip().lower() == "yes" else 0
        return np.nan

    def calc_aa_sponsor(row):
        is_active = pd.isna(row.get(COL_EXIT_DATE))
        if is_active:
            val = row.get("Acquired Sponsor_5137")
            return 1 if str(val).strip().lower() == "yes" else 0
        return np.nan

    def calc_ecm(row):
        val = row.get(COL_ENGAGEMENT)
        if pd.notna(val) and str(val).strip().lower() == "enrolled":
            return 1
        return 0

    def calc_aftercare_window_range(row, min_days, max_days):
        """Checks if a client had a valid checkin reaching them within a rolling day range from TODAY."""
        today = pd.Timestamp.now().normalize()
        checkin_dt = pd.to_datetime(row.get(COL_CHECKIN_DATE), errors="coerce")
        how_checked = str(row.get(COL_HOW_CHECKED, "")).strip().lower()

        if pd.isna(checkin_dt):
            return 0

        valid_method = any(method in how_checked for method in VALID_HOW_CHECKED)
        days_ago = (today - checkin_dt).days

        within_window = (days_ago >= min_days) and (days_ago <= max_days)
        return 1 if (valid_method and within_window) else 0

    def calculate_current_module(row):
        if pd.notna(row.get(COL_EXIT_DATE)):
            return np.nan

        city = row.get("City")
        mod_cols = CITY_MODULE_MAP.get(city)
        if not mod_cols:
            return np.nan

        latest_num = None
        for num, col_name in mod_cols:
            val = row.get(col_name)
            if pd.notna(val):
                d = pd.to_datetime(val, errors="coerce")
                if pd.notna(d):
                    latest_num = num

        return latest_num if latest_num is not None else np.nan

    # =========================================================================
    # Fiscal Date Helpers for Averages
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

        df["_fy_period"] = df["_Date"].apply(fy_period_fn) if fy_period_fn else df["_Date"].apply(period_fn)
        counts = df.groupby(["_fy_period", COL_PROGRAM]).size().to_dict()

        rows = []
        for _, p_row in unique_periods.iterrows():
            fy_period = p_row["_fy_period"]
            period = p_row["_period"]
            fy = p_row["_fy"]
            date_val = p_row["_week_date"] if is_weekly else p_row["_Date"].strftime("%Y-%m-%d")

            for prog in RENEW_PROGRAMS:
                cat_val = RENEW_PROGRAM_CATEGORY.get(prog, "Renew")
                city_val = assign_city_custom(prog)
                
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
    # 4. Tab Processors
    # =========================================================================
    def process_graduates(df_base):
        df = df_base.copy()
        df["_is_grad"] = df[COL_EXIT_REASON].isin(GRAD_REASONS).astype(int)
        df = (df.sort_values(["_is_grad", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
                .drop(columns=["_is_grad"]))
        df = df[df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(GRAD_REASONS)].copy()

        df["City"]            = df[COL_PROGRAM].apply(assign_city_custom)
        df["Year"]            = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
        df["Quarter"]         = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]          = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
        
        # Calculate First Day of Exit Month
        df["Month"]           = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce").dt.strftime("%Y-%m-01").fillna("")

        df["Days in Program"] = calc_days_in_program(df)
        
        # Calculate Rolling Outreach/Engagement Windows from TODAY
        df["Aftercare30"] = df.apply(lambda r: calc_aftercare_window_range(r, 0, 30), axis=1)
        df["Aftercare60"] = df.apply(lambda r: calc_aftercare_window_range(r, 31, 60), axis=1)
        df["Aftercare90"] = df.apply(lambda r: calc_aftercare_window_range(r, 61, 90), axis=1)

        df["Current FY Goal"]        = df[COL_PROGRAM].map(constants.PROGRAM_GOALS)
        df["Next FY Goal"]           = df[COL_PROGRAM].map(constants.NEXT_FY_GRAD_GOALS)
        df["Current FY Projection"]  = df[COL_PROGRAM].map(constants.PROGRAM_PROJECTIONS)
        df["Capacity"]               = df[COL_PROGRAM].map(constants.PROGRAM_CAPACITY)
        df["Theoretical Maximum"]    = df[COL_PROGRAM].map(constants.PROGRAM_THEORETICAL_MAX)

        df["FY27 Capacity"]          = df[COL_PROGRAM].map(FY27_CAPACITY_MAP)
        df["FY27 Occupancy Goal"]    = df[COL_PROGRAM].map(FY27_OCC_GOAL_MAP)

        for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
            df[k] = (df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= pd.Timestamp(w_start)) & (df[COL_EXIT_DATE] <= pd.Timestamp(w_end))).astype(int)

        cols = [COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, "Days in Program", COL_EXIT_REASON,
                "City", "Year", "Quarter", "Year Q", "Month",
                "Aftercare30", "Aftercare60", "Aftercare90",
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

        df["City"] = df[COL_PROGRAM].apply(assign_city_custom)
        df[constants.OCC_PRIOR_LABEL]   = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_PRIOR_DATE), axis=1)
        df[constants.OCC_CURRENT_LABEL] = df.apply(lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_CURRENT_DATE), axis=1)

        def count_kids_only(row):
            if row["City"] != "San Jose Women": return 0
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

        # Days in Program, Time Window Categorical Column, and universal Current Module calculation
        df["Days in Program"]    = calc_days_in_program(df)
        df["Days(Time Window)"]  = df["Days in Program"].apply(calc_days_time_window)
        df["Current Module"]     = df.apply(calculate_current_module, axis=1)

        # Calculate conditional Spiritual Mentor, AA Sponsor & ECM metrics
        df["Spiritual Mentor"] = df.apply(calc_spiritual_mentor, axis=1)
        df["AA Sponsor"]       = df.apply(calc_aa_sponsor, axis=1)
        df["ECM?"]             = df.apply(calc_ecm, axis=1)

        target_cols = [
            COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, "Days in Program",
            "Days(Time Window)", "City", constants.OCC_PRIOR_LABEL, constants.OCC_CURRENT_LABEL,
            COL_CHILD, "Number of Children", COL_SJ_CASE_MGR, "Current Module",
            "Spiritual Mentor", "AA Sponsor", COL_ENGAGEMENT, "ECM?"
        ] + ALL_MODULE_DATE_COLS + ["Capacity", "Goal", "Next FY Goal", "Prior FY Occupancy"]

        df = df[[c for c in target_cols if c in df.columns]].copy()

        df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
        df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)

        for date_col in ALL_MODULE_DATE_COLS:
            if date_col in df.columns:
                df[date_col] = fmt_date(df, date_col)

        return df.reset_index(drop=True)

    def process_retention(df_base):
        df = df_base.copy()

        df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
        df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE],  errors="coerce")

        df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
        df = (df.sort_values(["_no_exit", COL_EXIT_DATE], ascending=[False, False])
                .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
                .drop(columns=["_no_exit"]))

        today_ts = pd.Timestamp.now().normalize()
        day_31_date = df[COL_START_DATE] + pd.Timedelta(days=31)
        exit_date_filled = df[COL_EXIT_DATE].fillna(far_future)

        df["Exit Reason Category"] = df[COL_EXIT_REASON].apply(categorize_exit_reason)
        df["Days in Program"]     = calc_days_in_program(df)
        is_graduated = df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(GRAD_REASONS)

        in_up_to_date_cohort = (
            (df[COL_START_DATE] >= cohort_start) & 
            (df[COL_START_DATE] <= today_ts)
        )

        reached_31_up_to_today = (
            in_up_to_date_cohort &
            (day_31_date >= q_start_ts) & 
            (day_31_date <= today_ts) & 
            (exit_date_filled >= day_31_date)
        )
        df["Achieved 31 days"] = reached_31_up_to_today.astype(int)

        is_active_today = exit_date_filled > today_ts
        tenure_today = (today_ts - df[COL_START_DATE]).dt.days
        still_or_grad_31 = (
            in_up_to_date_cohort &
            (tenure_today >= 31) &
            (is_active_today | is_graduated)
        )
        df["Still in Program(31days)/Graduated"] = still_or_grad_31.astype(int)

        q3_cohort_start = q_start_ts - pd.DateOffset(months=4)
        q3_end_ts       = q_start_ts - pd.Timedelta(days=1)

        in_q3_cohort = (
            (df[COL_START_DATE] >= q3_cohort_start) & 
            (df[COL_START_DATE] <= q3_end_ts)
        )

        reached_31_q3 = (
            in_q3_cohort &
            (day_31_date >= q3_start_ts) & 
            (day_31_date <= q3_end_ts) & 
            (exit_date_filled >= day_31_date)
        )
        df["Achieved 31 days (Last Q)"] = reached_31_q3.astype(int)

        tenure_at_q3_end = (q3_end_ts - df[COL_START_DATE]).dt.days
        still_or_grad_31_q3 = (
            reached_31_q3 &
            (tenure_at_q3_end >= 31) &
            (is_active_today | is_graduated)
        )
        df["Still in Program(31days)/Graduated(Last Q)"] = still_or_grad_31_q3.astype(int)

        in_q4_cohort = (
            (df[COL_START_DATE] >= cohort_start) & 
            (df[COL_START_DATE] <= q_end_ts)
        )

        gap = (df[COL_EXIT_DATE].fillna(q_end_ts) - df[COL_START_DATE]).dt.days
        is_active_at_q_end = exit_date_filled > q_end_ts

        df["City"]     = df[COL_PROGRAM].apply(assign_city_custom)
        df["Capacity"] = df[COL_PROGRAM].map(constants.PROGRAM_CAPACITY)
        
        df["Year"]    = df[COL_START_DATE].apply(constants.get_fiscal_year)
        df["Quarter"] = df[COL_START_DATE].apply(constants.get_fiscal_quarter)
        df["Year Q"]  = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

        df["Entered Since Prior FY"] = (df[COL_EXIT_DATE].isna() | (gap > 30)).astype(int)
        df["Exit After 30 Days"]     = (df[COL_EXIT_DATE].notna() & (gap > 30) & (~is_graduated)).astype(int)
        df["Exit Before 30 Days"]  = (
            df[COL_EXIT_DATE].notna() &
            (df[COL_EXIT_DATE] >= q_start_ts) &
            (df[COL_EXIT_DATE] <= q_end_ts) &
            (gap < 30)
        ).astype(int)
        df["Graduated"]              = is_graduated.astype(int)
        df["Still in Program"]       = (in_q4_cohort & is_active_at_q_end).astype(int)

        cols = [COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, "Days in Program", COL_EXIT_REASON,
                "Exit Reason Category", "City", "Year", "Quarter", "Year Q", "Capacity",
                "Entered Since Prior FY", "Exit After 30 Days", "Exit Before 30 Days",
                "Graduated", "Still in Program", 
                "Still in Program(31days)/Graduated", "Achieved 31 days",
                "Still in Program(31days)/Graduated(Last Q)", "Achieved 31 days (Last Q)"]
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

    # =========================================================================
    # 7. Compute & Export Graduates Average Tabs
    # =========================================================================
    print("\nComputing Renew Graduates Averages...")
    write_averages_tab(spreadsheet, "Graduates Weekly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_week_label, get_fy_week_label, is_weekly=True, min_fy="FY26"), "Week")
    write_averages_tab(spreadsheet, "Graduates Monthly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_month, get_fy_month), "Month")
    write_averages_tab(spreadsheet, "Graduates Quarterly Avg", 
                       compute_kpi_averages(df_grad, COL_EXIT_DATE, get_quarter, get_fy_quarter), "Quarter")

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except:
        pass

    print(f"\n🎉 Processing complete!: {spreadsheet.url}")


if __name__ == "__main__":
    print("🚀 Starting Renew Processing")
    run_renew_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data"
    )