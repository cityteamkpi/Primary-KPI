#!/usr/bin/env python3
"""
Operations Processing Pipeline (ops_processor.py)
Consolidates key operational tabs from Renew, Turning Point, and Alumni datasets
into a single Google Sheet ('Ops Report - Processed') inside:
Shared drives -> CityTeam Impact Data -> Program Ops -> Program Ops Processed Data

Output Tabs:
1. Renew graduates weekly average
2. Renew graduates
3. Turning Point Graduates
4. Turning Points Successfully Housed
5. Renew - Occupancy
6. Renew - Retention
7. Alumni - Aftercare
8. Alumni - Sobriety
9. Renew - ECM clients
"""

import re
import warnings
import pandas as pd
import numpy as np
import gspread
from auth_utils import get_services
from drive_utils import (
    resolve_folder_id,
    resolve_folder_path,
    download_drive_file,
    find_file_id,
    load_raw,
    write_tab,
    parse_numeric_cell,
    load_program_goal_metrics,
)
import constants

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =============================================================================
# Shared Constants & Mappings
# =============================================================================
RAW_DATA_HEADER_ROW = 3
RAW_DATA_START_COL  = 1

COL_RECORD_ID      = "Record Id_102"
COL_NAME           = "Name_2057"
COL_PROGRAM        = "Program Enrolling_2091"
COL_INTERN_PROGRAM = "Intern Program_6619"
COL_POST_GRAD      = "Post Graduate Program Association_6619"
COL_START_DATE     = "Start Date_2090"
COL_EXIT_DATE      = "Exit Date_2100"
COL_EXIT_REASON    = "Primary Reason for Exit_2102"
COL_CHILD          = "Name of Child_6313"
COL_SJ_CASE_MGR    = "San Jose Men Case Managers_4668"

COL_CHECKIN_DATE   = "Checkin Date_6583"
COL_HOW_CHECKED    = "How Checked_6584"
COL_ENGAGEMENT     = "Status of Member Engagement_6386"
COL_SM_START_DATE  = "Start Date for Mentoring_4365"
COL_WAGE           = "Hourly Wage_6586"
COL_HOUSING        = "Housing Expense_6587"
COL_SOBER          = "Sober?_6585"

COL_HOUSED_OLD     = "Successfully Housed (Is Housing Healthy?)_4368"
COL_HOUSED_NEW     = "Successfully Housed_7011"
COL_HOUSING_CHECK  = "Date of Housing Check_5141"

INTERN_PROGRAM_NAMES = {
    "program graduate intern",
    "leadership module",
    "cdac trainee",
}

RENEW_SIX_MONTH_PROGRAMS = {
    "Portland Renew",
    "GV Renew",
}

RENEW_PROGRAM_CATEGORY = {
    "Chester Renew": "Renew",
    "GV Renew": "Renew",
    "Oakland Renew": "Renew",
    "Portland Renew": "Renew",
    "San Jose Men Renew": "Renew",
    "San Jose Women Renew": "Renew",
}

# Mapping of Program Enrolling -> (City KPI Goal Filename, Worksheet Tab Name)
PROGRAM_GOAL_FILES = {
    # Renew
    "Chester Renew":                 ("Chester FY27 KPI Goals",        "Men Renew"),
    "Oakland Renew":                 ("Oakland FY27 KPI Goals",        "Men Renew"),
    "Portland Renew":                ("Portland FY27 KPI Goals",       "Men Renew"),
    "San Jose Men Renew":            ("San Jose Men FY27 KPI Goals",   "Men Renew"),
    "GV Renew":                      ("San Jose Women FY27 KPI Goals", "GV Renew"),
    "San Jose Women Renew":          ("San Jose Women FY27 KPI Goals", "GV Renew"),
    # Turning Point
    "Chester Men Turning Point":     ("Chester FY27 KPI Goals",        "Men Turning Point"),
    "Chester Turning Point":         ("Chester FY27 KPI Goals",        "Men Turning Point"),
    "Chester Women Turning Point":   ("Chester FY27 KPI Goals",        "Women Turning Point"),
    "Chester Women's Turning Point": ("Chester FY27 KPI Goals",        "Women Turning Point"),
    "Oakland Men Turning Point":     ("Oakland FY27 KPI Goals",        "Men Turning Point"),
    "Oakland Youth Collective":      ("Oakland FY27 KPI Goals",        "Men Turning Point"),
    "Oakland Women Turning Point":   ("Oakland FY27 KPI Goals",        "Women Turning Point"),
    "Portland Men Turning Point":    ("Portland FY27 KPI Goals",       "Men Turning Point"),
    "Portland Youth Collective":     ("Portland FY27 KPI Goals",       "Men Turning Point"),
    "Portland Community of Hope":    ("Portland FY27 KPI Goals",       "Women Turning Point"),
    "San Jose Men Turning Point":    ("San Jose Men FY27 KPI Goals",   "Men Turning Point"),
    "GV Turning Point":              ("San Jose Women FY27 KPI Goals", "GV Turning Point"),
    "Heritage Home":                 ("San Jose Women FY27 KPI Goals", "Heritage Home"),
    "San Jose Youth Collective":     ("San Jose Women FY27 KPI Goals", "Youth Collective"),
}

CITY_MODULE_MAP = {
    "Oakland": [
        (1, "Module 1 Begin Date_3365"),
        (2, "Module 2 Begin Date_3898"),
        (3, "Module 3 Begin Date_3529"),
        (4, "Module 4 Begin Date_4244"),
    ],
    "Portland": [
        (1, "Module 1 Begin Date_3702"),
        (2, "Module 2 Begin Date_4014"),
        (3, "Module 3 Begin Date_4082"),
        (4, "Module 4 Begin Date_4160"),
    ],
    "Chester": [
        (1, "Module 1 Begin Date_4396"),
        (2, "Module 2 Begin Date_4479"),
        (3, "Module 3 Begin Date_4586"),
        (4, "Module 4 Begin Date_7204"),
    ],
    "San Jose Men": [
        (1, "Module 1 Begin Date_2604"),
        (2, "Module 2 Begin Date_2670"),
        (3, "Module 3 Begin Date_2768"),
        (4, "Module 4 Begin Date_2834"),
    ],
    "San Jose Women": [
        (1, "Module 1 Begin Date_2604"),
        (2, "Module 2 Begin Date_2670"),
        (3, "Module 3 Begin Date_2768"),
        (4, "Module 4 Begin Date_2834"),
    ],
}
ALL_MODULE_DATE_COLS = list({col for mods in CITY_MODULE_MAP.values() for _, col in mods})

TP_PROGRAMS_TO_INCLUDE = [
    "Heritage Home", "Oakland Women Turning Point", "San Jose Men Turning Point",
    "Chester Men Turning Point", "Chester Women Turning Point", "GV Turning Point",
    "Oakland Men Turning Point", "Oakland Youth Collective", "Portland Community of Hope",
    "Portland Youth Collective", "San Jose Youth Collective", "Chester Turning Point",
]

ALUM_GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]


# =============================================================================
# Shared Helper Functions
# =============================================================================
def load_fy27_kpi_goals(drive_service, gc, active_quarter):
    """
    Reads each city's FY27 KPI Goals sheet using shared drive_utils helpers:
      - Capacity: active quarter point value
      - Occupancy: active quarter point value (non-cumulative)
      - Graduates: cumulative sum Q1..active_quarter
      - Successfully Housed: cumulative sum Q1..active_quarter from row named 'Successfully Housed'
    """
    metrics = load_program_goal_metrics(
        drive_service=drive_service,
        gc=gc,
        program_goal_files=PROGRAM_GOAL_FILES,
        active_quarter=active_quarter,
        fy_prefix="FY27",
        metric_specs={
            "capacity":  {"patterns": r"^capacity$",            "mode": "point",      "default": np.nan},
            "occupancy": {"patterns": r"^occupancy$",           "mode": "point",      "default": np.nan},
            "graduates": {"patterns": r"^graduates",            "mode": "cumulative", "default": 0},
            "housed":    {"patterns": r"^successfully housed$", "mode": "cumulative", "default": 0},
        },
    )
    return metrics["capacity"], metrics["occupancy"], metrics["graduates"], metrics["housed"]


def load_graduate_projections(drive_service, gc):
    """
    Reads annual graduate projections from:
      CityTeam Impact Data -> Apricot Report Incoming -> Fiscal Year Renew Graduate Projections
    """
    proj_folder_id = resolve_folder_path(
        drive_service,
        ["CityTeam Impact Data", "Apricot Report Incoming"]
    )
    proj_file_name = "Fiscal Year Renew Graduate Projections"
    file_info = find_file_id(drive_service, proj_file_name, proj_folder_id)
    if not file_info:
        raise FileNotFoundError(f"Could not find '{proj_file_name}' in Apricot Report Incoming folder.")

    ss = gc.open_by_key(file_info["id"])
    try:
        ws = ss.worksheet("Sheet1")
    except gspread.WorksheetNotFound:
        ws = ss.get_worksheet(0)

    rows = ws.get_all_values()
    projections_map = {}
    for row in rows[1:]:
        if len(row) >= 2 and str(row[0]).strip():
            prog_name = str(row[0]).strip()
            projections_map[prog_name] = parse_numeric_cell(row[1], default=np.nan)

    return projections_map


def get_intern_program_value(row):
    post_grad = row.get(COL_POST_GRAD)
    intern_prog = row.get(COL_INTERN_PROGRAM)

    post_str = str(post_grad).strip() if pd.notna(post_grad) and str(post_grad).strip().lower() != "nan" else None
    intern_str = str(intern_prog).strip() if pd.notna(intern_prog) and str(intern_prog).strip().lower() != "nan" else None

    return post_str if post_str else intern_str


def fmt_date(df, col):
    return pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")


def was_active_on(start_date, exit_date, ref_date):
    ref = pd.Timestamp(ref_date)
    if pd.isnull(start_date) or start_date > ref:
        return 0
    if pd.isnull(exit_date) or exit_date > ref:
        return 1
    return 0


def assign_city_custom(program_name):
    prog_str = str(program_name).strip().lower()
    if any(x in prog_str for x in ["san jose women", "sj wp", "house of grace", "gv renew", "heritage home", "san jose youth"]):
        return "San Jose Women"
    return constants.assign_city(program_name)


def calc_days_in_program(df):
    today = pd.Timestamp.now().normalize()
    start = pd.to_datetime(df[COL_START_DATE], errors="coerce")
    exit_dt = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")
    end = exit_dt.fillna(today)
    days = (end - start).dt.days
    return days.where(start.notna(), np.nan)


def calc_renew_theoretical_max(row):
    cap_num = parse_numeric_cell(row.get("Capacity"), default=np.nan)
    if pd.isna(cap_num):
        return np.nan
    prog = str(row.get(COL_PROGRAM, "")).strip()
    multiplier = 2.0 if prog in RENEW_SIX_MONTH_PROGRAMS else 1.33
    return int(round(float(cap_num) * multiplier))


# =============================================================================
# 1 & 2. Renew Graduates & Renew Graduates Weekly Average
# =============================================================================
def process_renew_graduates(df_base, capacity_map, occ_goal_map, grad_goal_map, projections_map):
    df = df_base.copy()
    GRAD_REASONS = constants.GRAD_REASONS

    df["_is_grad"] = df[COL_EXIT_REASON].isin(GRAD_REASONS).astype(int)
    df = (
        df.sort_values(["_is_grad", COL_EXIT_DATE], ascending=[False, False])
        .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
        .drop(columns=["_is_grad"])
    )

    df = df[df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(GRAD_REASONS)].copy()

    df["City"]                  = df[COL_PROGRAM].apply(assign_city_custom)
    df["Year"]                  = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
    df["Quarter"]               = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
    df["Year Q"]                = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
    df["Month"]                 = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce").dt.strftime("%Y-%m-01").fillna("")
    df["Days in Program"]       = calc_days_in_program(df)

    df["Current FY Q Goal"]     = df[COL_PROGRAM].map(grad_goal_map)
    df["Current FY Projection"] = df[COL_PROGRAM].map(projections_map)
    df["Capacity"]              = df[COL_PROGRAM].map(capacity_map)
    df["Theoretical Maximum"]   = df.apply(calc_renew_theoretical_max, axis=1)
    df["FY27 Capacity"]         = df[COL_PROGRAM].map(capacity_map)
    df["FY27 Occupancy Goal"]   = df[COL_PROGRAM].map(occ_goal_map)

    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        df[k] = (
            df[COL_EXIT_DATE].notna()
            & (df[COL_EXIT_DATE] >= pd.Timestamp(w_start))
            & (df[COL_EXIT_DATE] <= pd.Timestamp(w_end))
        ).astype(int)

    cols = [
        COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE,
        "Days in Program", COL_EXIT_REASON, "City", "Year", "Quarter", "Year Q", "Month",
        "Current FY Q Goal", "Current FY Projection", "Capacity",
        "Theoretical Maximum", "FY27 Capacity", "FY27 Occupancy Goal"
    ] + list(constants.ACTUALS_WINDOWS.keys())

    df = df[[c for c in cols if c in df.columns]].copy()
    df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
    df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
    return df.reset_index(drop=True)


def compute_renew_grad_weekly_avg(df_grad):
    if df_grad.empty:
        return pd.DataFrame(columns=["FY", "Week", "FY Week", "Date", "Category", "City", "Program", "Average"])

    RENEW_PROGRAMS = list(constants.RENEW_PROGRAMS)

    def get_fy(date):
        return date.year + 1 if date.month >= 9 else date.year

    def get_fy_label(date):
        return f"FY{str(get_fy(date))[-2:]}"

    def get_week_num(date):
        fy = get_fy(date)
        fy_start = pd.Timestamp(f"{fy - 1}-09-01")
        delta = (date - fy_start).days
        return min(52, max(1, delta // 7 + 1))

    def get_week_label(date):
        return f"Week {get_week_num(date)}"

    def get_fy_week_label(date):
        return f"{get_fy_label(date)} Week {get_week_num(date)}"

    df = df_grad.copy()
    df["_Date"] = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")
    df = df.dropna(subset=["_Date"])
    if df.empty:
        return pd.DataFrame(columns=["FY", "Week", "FY Week", "Date", "Category", "City", "Program", "Average"])

    min_date = df["_Date"].min()
    max_date = pd.Timestamp.now()

    timeline = pd.date_range(start=min_date, end=max_date, freq="D")
    df_timeline = pd.DataFrame({"_Date": timeline})
    df_timeline["_period"]    = df_timeline["_Date"].apply(get_week_label)
    df_timeline["_fy_period"] = df_timeline["_Date"].apply(get_fy_week_label)
    df_timeline["_fy"]        = df_timeline["_Date"].apply(get_fy_label)
    df_timeline["_week_date"] = df_timeline["_Date"].apply(
        lambda d: (d - pd.Timedelta(days=d.weekday())).strftime("%Y-%m-%d")
    )

    df_timeline = df_timeline[df_timeline["_fy"] >= "FY26"]
    if df_timeline.empty:
        return pd.DataFrame(columns=["FY", "Week", "FY Week", "Date", "Category", "City", "Program", "Average"])

    unique_periods = df_timeline.groupby("_fy_period", sort=False).first().reset_index()

    df["_fy_period"] = df["_Date"].apply(get_fy_week_label)
    counts = df.groupby(["_fy_period", COL_PROGRAM]).size().to_dict()

    rows = []
    for _, p_row in unique_periods.iterrows():
        fy_period = p_row["_fy_period"]
        period    = p_row["_period"]
        fy        = p_row["_fy"]
        date_val  = p_row["_week_date"]

        for prog in RENEW_PROGRAMS:
            rows.append({
                "FY": fy,
                "Week": period,
                "FY Week": fy_period,
                "Date": date_val,
                "Category": RENEW_PROGRAM_CATEGORY.get(prog, "Renew"),
                "City": assign_city_custom(prog),
                "Program": prog,
                "Average": float(counts.get((fy_period, prog), 0)),
            })

    return pd.DataFrame(rows)


# =============================================================================
# 3 & 4. Turning Point Graduates & Turning Points Successfully Housed
# =============================================================================
def process_tp_graduates(df_tp_base, grad_goal_map):
    df = df_tp_base.copy()
    reasons_list = ["Graduation"]

    is_orig_intern = df["_ORIGINAL_PROGRAM"].astype(str).str.strip().str.lower().isin(INTERN_PROGRAM_NAMES)
    df = df[~is_orig_intern].copy()
    df = df[df[COL_PROGRAM].isin(TP_PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

    df["_priority"] = df[COL_EXIT_REASON].isin(reasons_list).astype(int)
    df = (
        df.sort_values(["_priority", COL_EXIT_DATE], ascending=[False, False])
        .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
        .drop(columns=["_priority"])
    )

    df = df[df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(reasons_list)].copy()

    def get_week_start(dt_series):
        dt = pd.to_datetime(dt_series, errors="coerce")
        return dt.apply(lambda d: (d - pd.Timedelta(days=d.weekday())).strftime("%Y-%m-%d") if pd.notna(d) else "")

    df["City"]              = df[COL_PROGRAM].apply(constants.assign_city)
    df["Year"]              = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
    df["Quarter"]           = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
    df["Year Q"]            = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()
    df["Week"]              = get_week_start(df[COL_EXIT_DATE])
    df["Current FY Q Goal"] = df[COL_PROGRAM].map(grad_goal_map).fillna(0).astype(int)

    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        df[k] = (
            df[COL_EXIT_DATE].notna() & (df[COL_EXIT_DATE] >= w_start) & (df[COL_EXIT_DATE] <= w_end)
        ).astype(int)

    cols = [
        COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, COL_EXIT_REASON,
        "City", "Year", "Quarter", "Year Q", "Week", "Current FY Q Goal"
    ] + list(constants.ACTUALS_WINDOWS.keys())

    df = df[[c for c in cols if c in df.columns]].copy()
    df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
    df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
    for k in constants.ACTUALS_WINDOWS.keys():
        if k in df.columns:
            df[k] = df[k].fillna(0).astype(int)
    return df.reset_index(drop=True)


def process_tp_housed(df_tp_raw, capacity_map, housed_goal_map):
    prog_lower_map = {p.lower(): p for p in TP_PROGRAMS_TO_INCLUDE}

    def match_program(val):
        if pd.isna(val):
            return val
        return prog_lower_map.get(str(val).strip().lower(), str(val).strip())

    def get_week_start(dt_series):
        dt = pd.to_datetime(dt_series, errors="coerce")
        return dt.apply(lambda d: (d - pd.Timedelta(days=d.weekday())).strftime("%Y-%m-%d") if pd.notna(d) else "")

    df = df_tp_raw.copy()
    df[COL_PROGRAM]    = df[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
    df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
    df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")
    if COL_HOUSING_CHECK in df.columns:
        df[COL_HOUSING_CHECK] = pd.to_datetime(df[COL_HOUSING_CHECK], errors="coerce")

    df["_INTERN_RESOLVED"] = df.apply(get_intern_program_value, axis=1)

    is_intern = df[COL_PROGRAM].astype(str).str.strip().str.lower().isin(INTERN_PROGRAM_NAMES)
    df.loc[is_intern, COL_PROGRAM] = df.loc[is_intern, "_INTERN_RESOLVED"].apply(match_program)

    df = df[df[COL_PROGRAM].isin(TP_PROGRAMS_TO_INCLUDE)].reset_index(drop=True)

    def eval_housed(row):
        val_n = str(row.get(COL_HOUSED_NEW, "")).strip().lower() if COL_HOUSED_NEW in row else ""
        if val_n not in ["", "nan", "none", "nat"]:
            return 1 if "yes" in val_n else 0
        val_o = str(row.get(COL_HOUSED_OLD, "")).strip().lower() if COL_HOUSED_OLD in row else ""
        if val_o not in ["", "nan", "none", "nat"]:
            return 1 if "yes" in val_o else 0
        return 0

    df["Successfully Housed?"] = df.apply(eval_housed, axis=1)

    actuals_cols = list(constants.ACTUALS_WINDOWS.keys())
    for k, (w_start, w_end) in constants.ACTUALS_WINDOWS.items():
        df[k] = (
            df[COL_EXIT_DATE].notna()
            & (df[COL_EXIT_DATE] >= w_start)
            & (df[COL_EXIT_DATE] <= w_end)
            & (df["Successfully Housed?"] == 1)
        ).astype(int)

    id_actuals = df.groupby([COL_RECORD_ID, COL_PROGRAM])[actuals_cols].max()

    df = (
        df.sort_values(["Successfully Housed?", COL_EXIT_DATE], ascending=[False, False])
        .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM], keep="first")
    )
    df = df.drop(columns=actuals_cols).merge(id_actuals, on=[COL_RECORD_ID, COL_PROGRAM], how="left")

    df["City"]     = df[COL_PROGRAM].apply(constants.assign_city)
    df["Year"]     = df[COL_EXIT_DATE].apply(constants.get_fiscal_year)
    df["Quarter"]  = df[COL_EXIT_DATE].apply(constants.get_fiscal_quarter)
    df["Year Q"]   = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

    housing_check_or_exit = (
        df[COL_HOUSING_CHECK].fillna(df[COL_EXIT_DATE])
        if COL_HOUSING_CHECK in df.columns
        else df[COL_EXIT_DATE]
    )
    df["Week"]     = get_week_start(housing_check_or_exit)
    df["Capacity"] = df[COL_PROGRAM].map(capacity_map)
    df["Goal"]     = df[COL_PROGRAM].map(housed_goal_map).fillna(0).astype(int)

    cols = [
        COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE,
        COL_HOUSING_CHECK, COL_EXIT_REASON, "City", "Year", "Quarter", "Year Q",
        "Week", "Successfully Housed?", "Capacity", "Goal"
    ] + actuals_cols

    df = df[[c for c in cols if c in df.columns]].copy()
    df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
    df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
    if COL_HOUSING_CHECK in df.columns:
        df[COL_HOUSING_CHECK] = fmt_date(df, COL_HOUSING_CHECK)
    for k in actuals_cols:
        if k in df.columns:
            df[k] = df[k].fillna(0).astype(int)

    return df.reset_index(drop=True)


# =============================================================================
# 5, 6, 9. Renew Occupancy, Renew Retention, Renew ECM Clients
# =============================================================================
def process_renew_occupancy(df_base, capacity_map, occ_goal_map):
    df = df_base.copy()
    df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
    df = (
        df.sort_values(["_no_exit", COL_START_DATE], ascending=[False, False])
        .drop_duplicates(subset=[COL_RECORD_ID], keep="first")
        .drop(columns=["_no_exit"])
    )

    df["City"] = df[COL_PROGRAM].apply(assign_city_custom)
    df[constants.OCC_PRIOR_LABEL] = df.apply(
        lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_PRIOR_DATE), axis=1
    )
    df[constants.OCC_CURRENT_LABEL] = df.apply(
        lambda r: was_active_on(r[COL_START_DATE], r[COL_EXIT_DATE], constants.OCC_CURRENT_DATE), axis=1
    )

    def count_kids_only(row):
        if row["City"] != "San Jose Women":
            return 0
        val = row.get(COL_CHILD)
        if pd.isna(val) or str(val).strip() == "":
            return 0
        parts = re.split(r"[,&]|\band\b", str(val).strip(), flags=re.IGNORECASE)
        return len([p for p in parts if p.strip()])

    def calc_days_time_window(days):
        if pd.isna(days):
            return np.nan
        if days <= 7:
            return "7days"
        elif days <= 30:
            return "8~30days"
        elif days <= 60:
            return "31~60days"
        elif days <= 90:
            return "61~90days"
        else:
            return "90~days"

    def calculate_current_module(row):
        if pd.notna(row.get(COL_EXIT_DATE)):
            return np.nan
        mod_cols = CITY_MODULE_MAP.get(row.get("City"))
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

    def calc_spiritual_mentor(row):
        if pd.isna(row.get(COL_EXIT_DATE)):
            val = row.get("Acquired Spiritual Mentor_5136")
            return 1 if str(val).strip().lower() == "yes" else 0
        return np.nan

    def calc_aa_sponsor(row):
        if pd.isna(row.get(COL_EXIT_DATE)):
            val = row.get("Acquired Sponsor_5137")
            return 1 if str(val).strip().lower() == "yes" else 0
        return np.nan

    def calc_sm_on_or_before_mod2(row):
        sm_start_val = row.get(COL_SM_START_DATE)
        if pd.isna(sm_start_val):
            return 0
        sm_start_dt = pd.to_datetime(sm_start_val, errors="coerce")
        if pd.isna(sm_start_dt):
            return 0
        mod_cols = CITY_MODULE_MAP.get(row.get("City"), [])
        mod2_col = next((col_name for num, col_name in mod_cols if num == 2), None)
        if not mod2_col or pd.isna(row.get(mod2_col)):
            return 0
        mod2_dt = pd.to_datetime(row.get(mod2_col), errors="coerce")
        if pd.isna(mod2_dt):
            return 0
        return 1 if sm_start_dt <= mod2_dt else 0

    df["Number of Children"]    = df.apply(count_kids_only, axis=1)
    df["Capacity"]              = df[COL_PROGRAM].map(capacity_map)
    df["Goal"]                  = df[COL_PROGRAM].map(occ_goal_map)
    df["Prior FY Occupancy"]    = np.nan
    df["Days in Program"]       = calc_days_in_program(df)
    df["Days(Time Window)"]     = df["Days in Program"].apply(calc_days_time_window)
    df["Current Module"]        = df.apply(calculate_current_module, axis=1)
    df["Spiritual Mentor"]      = df.apply(calc_spiritual_mentor, axis=1)
    df["SM on or before Mod2?"] = df.apply(calc_sm_on_or_before_mod2, axis=1)
    df["AA Sponsor"]            = df.apply(calc_aa_sponsor, axis=1)

    target_cols = [
        COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, "Days in Program",
        "Days(Time Window)", "City", constants.OCC_PRIOR_LABEL, constants.OCC_CURRENT_LABEL,
        COL_CHILD, "Number of Children", COL_SJ_CASE_MGR, "Current Module",
        COL_SM_START_DATE, "Spiritual Mentor", "SM on or before Mod2?", "AA Sponsor", COL_ENGAGEMENT
    ] + ALL_MODULE_DATE_COLS + ["Capacity", "Goal", "Prior FY Occupancy"]

    df = df[[c for c in target_cols if c in df.columns]].copy()

    df[COL_START_DATE]    = fmt_date(df, COL_START_DATE)
    df[COL_EXIT_DATE]     = fmt_date(df, COL_EXIT_DATE)
    df[COL_SM_START_DATE] = fmt_date(df, COL_SM_START_DATE)
    for date_col in ALL_MODULE_DATE_COLS:
        if date_col in df.columns:
            df[date_col] = fmt_date(df, date_col)

    return df.reset_index(drop=True)


def process_renew_retention(df_base, capacity_map):
    df = df_base.copy()
    GRAD_REASONS = constants.GRAD_REASONS

    q_start_ts   = pd.Timestamp(constants.CURRENT_Q_START)
    q_end_ts     = pd.Timestamp(constants.CURRENT_Q_END)
    cohort_start = q_start_ts - pd.DateOffset(months=1)
    far_future   = pd.Timestamp("2099-12-31")

    df[COL_START_DATE] = pd.to_datetime(df[COL_START_DATE], errors="coerce")
    df[COL_EXIT_DATE]  = pd.to_datetime(df[COL_EXIT_DATE], errors="coerce")

    df["_no_exit"] = df[COL_EXIT_DATE].isna().astype(int)
    df = (
        df.sort_values(["_no_exit", COL_EXIT_DATE], ascending=[False, False])
        .drop_duplicates(subset=[COL_RECORD_ID, COL_PROGRAM, COL_START_DATE], keep="first")
        .drop(columns=["_no_exit"])
    )

    def categorize_exit_reason(val):
        if pd.isna(val) or str(val).strip() == "":
            return None
        v = str(val).lower().strip()
        if any(x in v for x in ["graduation", "completer", "completion", "internship ended"]):
            return "Graduation"
        if any(x in v for x in ["relapse", "drug", "detox", "fentanyl", "paraph", "blackout", "deviating"]):
            return "Relapse"
        if any(x in v for x in ["discipline", "conduct", "behavior", "threat", "violence", "awol", "30 day review", "incident"]):
            return "Discipline"
        if any(x in v for x in ["family", "child", "wife", "daughter", "son", "mother", "reunif"]):
            return "Family"
        if any(x in v for x in ["medical", "mental", "health", "hospital", "cognitive", "ssi", "out of scope"]):
            return "Medical/Mental Health"
        if any(x in v for x in ["legal", "parole", "probation", "warrant", "criminal", "remanded", "county", "immigr"]):
            return "Legal"
        if any(x in v for x in ["transfer", "another program", "different program", "forward", "other program"]):
            return "Transfer"
        if any(x in v for x in ["housing", "sle", "perm housing"]):
            return "Housing"
        if any(x in v for x in ["job", "work", "employ"]):
            return "Job"
        if any(x in v for x in ["relationship"]):
            return "Relationship"
        if any(x in v for x in ["not eligible", "duplicate", "accident", "test client", "mistake", "lied", "restart", "not accepted"]):
            return "Not Eligible/Admin"
        if any(x in v for x in ["self", "not ready", "voluntary", "chose", "decided", "walked", "left", "wasn't ready", "refused"]):
            return "Self Exit"
        return "Unknown/Other"

    today_ts         = pd.Timestamp.now().normalize()
    day_31_date      = df[COL_START_DATE] + pd.Timedelta(days=31)
    exit_date_filled = df[COL_EXIT_DATE].fillna(far_future)

    df["Exit Reason Category"] = df[COL_EXIT_REASON].apply(categorize_exit_reason)
    df["Days in Program"]      = calc_days_in_program(df)
    is_graduated               = df[COL_EXIT_DATE].notna() & df[COL_EXIT_REASON].isin(GRAD_REASONS)

    in_up_to_date_cohort = (df[COL_START_DATE] >= cohort_start) & (df[COL_START_DATE] <= today_ts)
    reached_31_up_to_today = (
        in_up_to_date_cohort
        & (day_31_date >= q_start_ts)
        & (day_31_date <= today_ts)
        & (exit_date_filled >= day_31_date)
    )
    df["Achieved 31 days"] = reached_31_up_to_today.astype(int)

    is_active_today = exit_date_filled > today_ts
    tenure_today = (today_ts - df[COL_START_DATE]).dt.days
    still_or_grad_31 = in_up_to_date_cohort & (tenure_today >= 31) & (is_active_today | is_graduated)
    df["Still in Program(31days)/Graduated"] = still_or_grad_31.astype(int)

    q3_cohort_start = q_start_ts - pd.DateOffset(months=4)
    q3_end_ts       = q_start_ts - pd.Timedelta(days=1)

    in_q3_cohort = (df[COL_START_DATE] >= q3_cohort_start) & (df[COL_START_DATE] <= q3_end_ts)
    reached_31_q3 = (
        in_q3_cohort
        & (day_31_date >= q_start_ts - pd.DateOffset(months=3))
        & (day_31_date <= q3_end_ts)
        & (exit_date_filled >= day_31_date)
    )
    df["Achieved 31 days (Last Q)"] = reached_31_q3.astype(int)

    tenure_at_q3_end = (q3_end_ts - df[COL_START_DATE]).dt.days
    still_or_grad_31_q3 = reached_31_q3 & (tenure_at_q3_end >= 31) & (is_active_today | is_graduated)
    df["Still in Program(31days)/Graduated(Last Q)"] = still_or_grad_31_q3.astype(int)

    in_q4_cohort       = (df[COL_START_DATE] >= cohort_start) & (df[COL_START_DATE] <= q_end_ts)
    gap                = (df[COL_EXIT_DATE].fillna(q_end_ts) - df[COL_START_DATE]).dt.days
    is_active_at_q_end = exit_date_filled > q_end_ts

    df["City"]     = df[COL_PROGRAM].apply(assign_city_custom)
    df["Capacity"] = df[COL_PROGRAM].map(capacity_map)
    df["Year"]     = df[COL_START_DATE].apply(constants.get_fiscal_year)
    df["Quarter"]  = df[COL_START_DATE].apply(constants.get_fiscal_quarter)
    df["Year Q"]   = (df["Year"].fillna("") + " " + df["Quarter"].fillna("")).str.strip()

    df["Entered Since Prior FY"] = (df[COL_EXIT_DATE].isna() | (gap > 30)).astype(int)
    df["Exit After 30 Days"]     = (df[COL_EXIT_DATE].notna() & (gap > 30) & (~is_graduated)).astype(int)
    df["Exit Before 30 Days"]    = (
        df[COL_EXIT_DATE].notna()
        & (df[COL_EXIT_DATE] >= q_start_ts)
        & (df[COL_EXIT_DATE] <= q_end_ts)
        & (gap < 30)
    ).astype(int)
    df["Graduated"]        = is_graduated.astype(int)
    df["Still in Program"] = (in_q4_cohort & is_active_at_q_end).astype(int)

    cols = [
        COL_RECORD_ID, COL_NAME, COL_PROGRAM, COL_START_DATE, COL_EXIT_DATE, "Days in Program",
        COL_EXIT_REASON, "Exit Reason Category", "City", "Year", "Quarter", "Year Q", "Capacity",
        "Entered Since Prior FY", "Exit After 30 Days", "Exit Before 30 Days",
        "Graduated", "Still in Program",
        "Still in Program(31days)/Graduated", "Achieved 31 days",
        "Still in Program(31days)/Graduated(Last Q)", "Achieved 31 days (Last Q)"
    ]
    df = df[[c for c in cols if c in df.columns]].copy()

    df[COL_START_DATE] = fmt_date(df, COL_START_DATE)
    df[COL_EXIT_DATE]  = fmt_date(df, COL_EXIT_DATE)
    return df.reset_index(drop=True)


def process_renew_ecm_clients(df_mod_raw):
    df = df_mod_raw.copy()
    if COL_ENGAGEMENT not in df.columns:
        return pd.DataFrame()

    is_enrolled = df[COL_ENGAGEMENT].fillna("").astype(str).str.strip().str.lower() == "enrolled"
    df_ecm = df[is_enrolled].copy()

    df_ecm["_start_dt"] = pd.to_datetime(df_ecm[COL_START_DATE], errors="coerce")
    df_ecm = (
        df_ecm.sort_values(by="_start_dt", ascending=False)
        .drop_duplicates(subset=[COL_RECORD_ID], keep="first")
        .drop(columns=["_start_dt"])
        .reset_index(drop=True)
    )

    def assign_ecm_city(row):
        prog = str(row.get(COL_PROGRAM, "")).strip()
        intern_prog = str(row.get("_INTERN_RESOLVED", "")).strip()

        prog_lower = prog.lower()
        intern_lower = intern_prog.lower()

        if prog_lower in INTERN_PROGRAM_NAMES or "intern" in prog_lower or (intern_prog and intern_lower != "nan"):
            target_str = intern_lower if (intern_prog and intern_lower != "nan") else prog_lower
            if "chester" in target_str:
                return "Chester"
            elif "oakland" in target_str:
                return "Oakland"
            elif "portland" in target_str:
                return "Portland"
            elif "san jose men" in target_str or "sj men" in target_str or "mp" in target_str:
                return "San Jose Men"
            elif any(x in target_str for x in ["san jose women", "sj women", "wp", "house of grace", "gv", "heritage home", "san jose youth"]):
                return "San Jose Women"

        if any(x in prog_lower for x in ["sj wp", "wp transition", "heritage home", "san jose youth"]):
            return "San Jose Women"

        return assign_city_custom(prog)

    df_ecm["City"] = df_ecm.apply(assign_ecm_city, axis=1)
    return df_ecm


# =============================================================================
# 7 & 8. Alumni Aftercare & Alumni Sobriety
# =============================================================================
def process_alumni_tabs(df_alum_raw):
    target_fy_num = constants.fy_num - 1
    target_fy_str = f"FY{str(target_fy_num)[-2:]}"

    def clean_wage(val):
        if pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        clean = str(val).replace("$", "").replace(",", "").strip()
        try:
            return float(clean)
        except Exception:
            return 0.0

    def max_nonzero(series):
        vals = series.replace(0, np.nan).dropna()
        return vals.max() if not vals.empty else 0.0

    def get_grad_wage(group):
        if COL_HOW_CHECKED in group.columns:
            grad_checked = group[group[COL_HOW_CHECKED] == "Graduation"]
            if not grad_checked.empty:
                return float(grad_checked[COL_WAGE].iloc[0])
        return max_nonzero(group[COL_WAGE])

    df_raw_parsed = df_alum_raw.dropna(subset=[COL_NAME]).copy()
    if COL_RECORD_ID not in df_raw_parsed.columns:
        df_raw_parsed[COL_RECORD_ID] = np.nan

    df_raw_parsed[COL_EXIT_DATE] = pd.to_datetime(df_raw_parsed[COL_EXIT_DATE], errors="coerce")
    df_raw_parsed[COL_CHECKIN_DATE] = pd.to_datetime(df_raw_parsed[COL_CHECKIN_DATE], errors="coerce")
    if COL_START_DATE in df_raw_parsed.columns:
        df_raw_parsed[COL_START_DATE] = pd.to_datetime(df_raw_parsed[COL_START_DATE], errors="coerce")
    df_raw_parsed[COL_WAGE] = df_raw_parsed[COL_WAGE].apply(clean_wage)
    df_raw_parsed[COL_HOUSING] = (
        df_raw_parsed[COL_HOUSING].apply(clean_wage) if COL_HOUSING in df_raw_parsed.columns else 0.0
    )
    if COL_SOBER in df_raw_parsed.columns:
        df_raw_parsed[COL_SOBER] = df_raw_parsed[COL_SOBER].astype(str).str.strip()

    renew_programs_set = {str(p).strip().lower() for p in constants.RENEW_PROGRAMS}

    today_ts = pd.Timestamp.now().normalize()
    df_grads_renew = df_raw_parsed[
        df_raw_parsed[COL_PROGRAM].astype(str).str.strip().str.lower().isin(renew_programs_set)
        & df_raw_parsed[COL_EXIT_REASON].isin(ALUM_GRAD_REASONS)
        & df_raw_parsed[COL_EXIT_DATE].notna()
    ].copy()

    if df_grads_renew.empty:
        df_aftercare = pd.DataFrame()
    else:
        orig_renew_map = (
            df_grads_renew.sort_values(by=COL_EXIT_DATE)
            .groupby(COL_RECORD_ID)[COL_PROGRAM]
            .last()
            .to_dict()
        )
        grad_record_ids = set(orig_renew_map.keys())
        df_client_records = df_raw_parsed[df_raw_parsed[COL_RECORD_ID].isin(grad_record_ids)].copy()
        df_client_records["_checkin_dt"] = pd.to_datetime(df_client_records[COL_CHECKIN_DATE], errors="coerce")

        df_latest = (
            df_client_records[df_client_records["_checkin_dt"].notna()]
            .sort_values(by=["_checkin_dt", COL_EXIT_DATE], ascending=[False, False])
            .drop_duplicates(subset=[COL_RECORD_ID], keep="first")
            .copy()
        )

        if df_latest.empty:
            df_aftercare = pd.DataFrame()
        else:
            df_latest[COL_PROGRAM] = df_latest[COL_RECORD_ID].map(orig_renew_map).fillna(df_latest[COL_PROGRAM])
            df_latest["_days_ago"] = (today_ts - df_latest["_checkin_dt"]).dt.days

            df_latest["Aftercare30"]     = (df_latest["_days_ago"] <= 30).astype(int)
            df_latest["Aftercare60"]     = ((df_latest["_days_ago"] >= 31) & (df_latest["_days_ago"] <= 60)).astype(int)
            df_latest["Aftercare90"]     = ((df_latest["_days_ago"] >= 61) & (df_latest["_days_ago"] <= 90)).astype(int)
            df_latest["Aftercare91Plus"] = (df_latest["_days_ago"] >= 91).astype(int)

            df_latest["FY"]   = df_latest[COL_EXIT_DATE].apply(constants.get_fiscal_year)
            df_latest["FY Q"] = df_latest[COL_EXIT_DATE].apply(
                lambda d: f"{constants.get_fiscal_year(d) or ''} {constants.get_fiscal_quarter(d) or ''}".strip()
                if pd.notna(d)
                else ""
            )

            df_latest[COL_START_DATE]   = fmt_date(df_latest, COL_START_DATE)
            df_latest[COL_EXIT_DATE]    = fmt_date(df_latest, COL_EXIT_DATE)
            df_latest[COL_CHECKIN_DATE] = fmt_date(df_latest, COL_CHECKIN_DATE)

            ac_cols = [
                COL_RECORD_ID, COL_NAME, COL_CHECKIN_DATE, COL_PROGRAM, COL_INTERN_PROGRAM,
                COL_START_DATE, COL_EXIT_DATE, "FY", "FY Q",
                "Aftercare30", "Aftercare60", "Aftercare90", "Aftercare91Plus"
            ]
            df_aftercare = df_latest[[c for c in ac_cols if c in df_latest.columns]].reset_index(drop=True)

    df_grads = (
        df_grads_renew
        .sort_values(by=[COL_EXIT_DATE, COL_EXIT_REASON])
        .drop_duplicates(subset=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE], keep="last")
        .reset_index(drop=True)
    )

    _grad_rows = df_raw_parsed[
        df_raw_parsed[COL_EXIT_REASON].isin(ALUM_GRAD_REASONS) & df_raw_parsed[COL_EXIT_DATE].notna()
    ].copy()
    _grad_rows[COL_EXIT_DATE] = pd.to_datetime(_grad_rows[COL_EXIT_DATE]).dt.normalize()

    def agg_grad_group(g):
        return pd.Series({
            COL_WAGE: get_grad_wage(g),
            COL_HOUSING: max_nonzero(g[COL_HOUSING]) if COL_HOUSING in g.columns else 0.0,
        })

    df_merged = (
        _grad_rows
        .groupby([COL_NAME, COL_PROGRAM, COL_EXIT_DATE], sort=False)
        .apply(agg_grad_group)
        .reset_index()
    )

    _grad_dates = df_grads[[COL_NAME, COL_PROGRAM, COL_EXIT_DATE]].rename(columns={COL_EXIT_DATE: "grad_date"})
    _grad_dates["grad_date"] = pd.to_datetime(_grad_dates["grad_date"]).dt.normalize()
    _all_with_grad = df_raw_parsed.merge(_grad_dates, on=[COL_NAME, COL_PROGRAM], how="inner")
    _post = _all_with_grad[
        _all_with_grad[COL_CHECKIN_DATE].notna() & (_all_with_grad[COL_CHECKIN_DATE] > _all_with_grad["grad_date"])
    ].copy()

    if COL_START_DATE in _post.columns:
        _post = _post[
            (_post[COL_START_DATE].isna())
            | (_post[COL_START_DATE] <= _post["grad_date"])
            | (_post[COL_CHECKIN_DATE] < _post[COL_START_DATE])
        ].copy()

    if not _post.empty:
        _latest = _post.groupby([COL_NAME, COL_PROGRAM, "grad_date"])[COL_CHECKIN_DATE].transform("max")
        _post_latest = _post[_post[COL_CHECKIN_DATE] == _latest].copy()
        df_recent = (
            _post_latest.groupby([COL_NAME, COL_PROGRAM, "grad_date"], sort=False)
            .agg(
                recent_wage=(COL_WAGE, max_nonzero),
                recent_housing=(COL_HOUSING, "max"),
                recent_checkin=(COL_CHECKIN_DATE, "max"),
            )
            .reset_index()
            .rename(columns={"grad_date": COL_EXIT_DATE})
        )
    else:
        df_recent = pd.DataFrame(
            columns=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE, "recent_wage", "recent_housing", "recent_checkin"]
        )

    df_merged = df_merged.merge(df_recent, on=[COL_NAME, COL_PROGRAM, COL_EXIT_DATE], how="left")

    sobriety_rows = []
    for _, row in df_grads.iterrows():
        name = row[COL_NAME]
        program = row[COL_PROGRAM]
        grad_date = row[COL_EXIT_DATE]

        person_records = df_raw_parsed[
            (df_raw_parsed[COL_NAME] == name) & (df_raw_parsed[COL_PROGRAM] == program)
        ].sort_values(by=COL_CHECKIN_DATE)

        if COL_START_DATE in person_records.columns:
            person_records = person_records[
                (person_records[COL_START_DATE].isna())
                | (person_records[COL_START_DATE] <= grad_date)
                | (person_records[COL_CHECKIN_DATE] < person_records[COL_START_DATE])
            ]

        sober_records = (
            person_records[person_records[COL_SOBER].isin(["Yes", "No"]) & (person_records[COL_CHECKIN_DATE] > grad_date)]
            if COL_SOBER in person_records.columns
            else pd.DataFrame()
        )

        has_sustained_relapse = False
        first_no_date = None
        last_checkin_sober = False

        for _, s_row in sober_records.iterrows():
            cur_state = s_row[COL_SOBER]
            cur_date = s_row[COL_CHECKIN_DATE]
            if cur_state == "No":
                if first_no_date is None:
                    first_no_date = cur_date
            elif cur_state == "Yes":
                if first_no_date is not None:
                    if (cur_date - first_no_date).days > 30:
                        has_sustained_relapse = True
                    first_no_date = None
            last_checkin_sober = cur_state == "Yes"

        is_sober = True if sober_records.empty else (last_checkin_sober and not has_sustained_relapse)

        merged_row = df_merged[
            (df_merged[COL_NAME] == name)
            & (df_merged[COL_PROGRAM] == program)
            & (df_merged[COL_EXIT_DATE] == grad_date)
        ]
        _recent_checkin_val = (
            merged_row["recent_checkin"].iloc[0]
            if not merged_row.empty and "recent_checkin" in merged_row.columns
            else None
        )
        recent_checkin = (
            _recent_checkin_val if _recent_checkin_val is not None and pd.notna(_recent_checkin_val) else pd.NaT
        )

        fy_val = constants.get_fiscal_year(grad_date)
        q_val  = constants.get_fiscal_quarter(grad_date)

        sobriety_rows.append({
            COL_NAME: name,
            COL_PROGRAM: program,
            "Graduation Date": grad_date,
            "Most Recent Checkin": recent_checkin,
            "City": constants.assign_city(program),
            "Year": fy_val,
            "Quarter": q_val,
            "Year Q": f"{fy_val or ''} {q_val or ''}".strip(),
            "Sustained Relapse?": "Yes" if has_sustained_relapse else "No",
            "Sobriety 1 Year": int(is_sober),
        })

    sobriety_cols = [
        COL_NAME, COL_PROGRAM, "Graduation Date", "Most Recent Checkin",
        "City", "Year", "Quarter", "Year Q", "Sustained Relapse?", "Sobriety 1 Year"
    ]

    if sobriety_rows:
        results_df = pd.DataFrame(sobriety_rows)
        df_sobriety = results_df[results_df["Year"] == target_fy_str][sobriety_cols].copy().reset_index(drop=True)
    else:
        df_sobriety = pd.DataFrame(columns=sobriety_cols)

    return df_aftercare, df_sobriety


# =============================================================================
# Main Orchestrator
# =============================================================================
def run_ops_processing(
    renew_input_file  = "Renew Programs Report.xlsx",
    module_input_file = "Renew Module Progress.xlsx",
    tp_input_file     = "Turning Point Report.xlsx",
    alumni_input_file = "All Programs All Alumni KPIs.xlsx",
    output_file       = "Ops Report - Processed",
    input_folder_name = "Apricot Report Incoming",
    output_folder_path = ["CityTeam Impact Data", "Program Ops", "Program Ops Processed Data"],
):
    print("🚀 Starting Operations Processing (ops_processor)...")

    drive_service, gc, _ = get_services()
    constants.sync_constants()

    q_start_ts     = pd.Timestamp(constants.CURRENT_Q_START)
    active_quarter = constants.get_fiscal_quarter(q_start_ts)
    print(f"📅 Active Reporting Quarter: {constants.get_fiscal_year(q_start_ts)} {active_quarter}")

    capacity_map, occ_goal_map, grad_goal_map, housed_goal_map = load_fy27_kpi_goals(
        drive_service, gc, active_quarter
    )
    projections_map = load_graduate_projections(drive_service, gc)

    input_folder_id  = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_path(drive_service, output_folder_path)

    print(f"Downloading {renew_input_file}...")
    fh_renew, _, _ = download_drive_file(drive_service, renew_input_file, input_folder_id)
    df_renew_raw = load_raw(fh_renew, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {module_input_file}...")
    fh_mod, _, _ = download_drive_file(drive_service, module_input_file, input_folder_id)
    df_mod_raw = load_raw(fh_mod, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {tp_input_file}...")
    fh_tp, _, _ = download_drive_file(drive_service, tp_input_file, input_folder_id)
    df_tp_raw = load_raw(fh_tp, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    print(f"Downloading {alumni_input_file}...")
    fh_alum, _, _ = download_drive_file(drive_service, alumni_input_file, input_folder_id)
    df_alum_raw = load_raw(fh_alum, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    for df_target in [df_renew_raw, df_mod_raw]:
        df_target["_INTERN_RESOLVED"] = df_target.apply(get_intern_program_value, axis=1)

    dup_cols = [c for c in df_mod_raw.columns if c in df_renew_raw.columns and c != COL_RECORD_ID]
    df_mod_clean = df_mod_raw.drop(columns=dup_cols)
    df_renew_merged = df_renew_raw.merge(df_mod_clean, on=COL_RECORD_ID, how="left")

    for col in ALL_MODULE_DATE_COLS + [
        COL_CHECKIN_DATE, COL_HOW_CHECKED, COL_ENGAGEMENT, COL_SM_START_DATE,
        "Acquired Spiritual Mentor_5136", "Acquired Sponsor_5137"
    ]:
        if col not in df_renew_merged.columns:
            df_renew_merged[col] = np.nan

    df_renew_base = df_renew_merged.copy()
    df_renew_base[COL_START_DATE] = pd.to_datetime(df_renew_base[COL_START_DATE], errors="coerce")
    df_renew_base[COL_EXIT_DATE]  = pd.to_datetime(df_renew_base[COL_EXIT_DATE], errors="coerce")
    df_renew_base[COL_PROGRAM]    = df_renew_base[COL_PROGRAM].replace({
        "House of Grace": "GV Renew",
        "San Jose Women Renew": "GV Renew",
    })
    df_renew_base = df_renew_base[df_renew_base[COL_PROGRAM].isin(list(constants.RENEW_PROGRAMS))].reset_index(drop=True)

    prog_lower_map = {p.lower(): p for p in TP_PROGRAMS_TO_INCLUDE}
    def match_tp_program(val):
        if pd.isna(val):
            return val
        return prog_lower_map.get(str(val).strip().lower(), str(val).strip())

    df_tp_base = df_tp_raw.copy()
    df_tp_base[COL_PROGRAM]    = df_tp_base[COL_PROGRAM].replace("Chester Women's Turning Point", "Chester Women Turning Point")
    df_tp_base[COL_START_DATE] = pd.to_datetime(df_tp_base[COL_START_DATE], errors="coerce")
    df_tp_base[COL_EXIT_DATE]  = pd.to_datetime(df_tp_base[COL_EXIT_DATE], errors="coerce")
    if COL_HOUSING_CHECK in df_tp_base.columns:
        df_tp_base[COL_HOUSING_CHECK] = pd.to_datetime(df_tp_base[COL_HOUSING_CHECK], errors="coerce")

    df_tp_base["_INTERN_RESOLVED"]  = df_tp_base.apply(get_intern_program_value, axis=1)
    df_tp_base["_ORIGINAL_PROGRAM"] = df_tp_base[COL_PROGRAM]

    is_tp_intern = df_tp_base[COL_PROGRAM].astype(str).str.strip().str.lower().isin(INTERN_PROGRAM_NAMES)
    df_tp_base.loc[is_tp_intern, COL_PROGRAM] = df_tp_base.loc[is_tp_intern, "_INTERN_RESOLVED"].apply(match_tp_program)

    print("Processing 1 & 2: Renew Graduates & Weekly Average...")
    df_renew_grads      = process_renew_graduates(df_renew_base, capacity_map, occ_goal_map, grad_goal_map, projections_map)
    df_renew_grad_wkavg = compute_renew_grad_weekly_avg(df_renew_grads)

    print("Processing 3 & 4: Turning Point Graduates & Successfully Housed...")
    df_tp_grads  = process_tp_graduates(df_tp_base, grad_goal_map)
    df_tp_housed = process_tp_housed(df_tp_raw, capacity_map, housed_goal_map)

    print("Processing 5, 6, 9: Renew Occupancy, Retention & ECM Clients...")
    df_renew_occ = process_renew_occupancy(df_renew_base, capacity_map, occ_goal_map)
    df_renew_ret = process_renew_retention(df_renew_base, capacity_map)
    df_renew_ecm = process_renew_ecm_clients(df_mod_raw)

    print("Processing 7 & 8: Alumni Aftercare & Sobriety...")
    df_alum_aftercare, df_alum_sobriety = process_alumni_tabs(df_alum_raw)

    processed_file = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")
    if processed_file:
        spreadsheet = gc.open_by_key(processed_file["id"])
        print(f"📄 Opened existing spreadsheet: '{output_file}'")
    else:
        file_metadata = {
            "name": output_file,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [output_folder_id],
        }
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        spreadsheet = gc.open_by_key(new_sheet["id"])
        print(f"📄 Created new spreadsheet: '{output_file}'")

    tabs_to_write = [
        ("Renew graduates weekly average",     df_renew_grad_wkavg),
        ("Renew graduates",                    df_renew_grads),
        ("Turning Point Graduates",            df_tp_grads),
        ("Turning Points Successfully Housed", df_tp_housed),
        ("Renew - Occupancy",                  df_renew_occ),
        ("Renew - Retention",                  df_renew_ret),
        ("Alumni - Aftercare",                 df_alum_aftercare),
        ("Alumni - Sobriety",                  df_alum_sobriety),
        ("Renew - ECM clients",                df_renew_ecm),
    ]

    for tab_name, df_out in tabs_to_write:
        print(f"Writing '{tab_name}' ({len(df_out)} rows)...")
        write_tab(spreadsheet, tab_name, df_out)

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except Exception:
        pass

    print(f"\n🎉 Operations Processing Complete! View sheet at: {spreadsheet.url}")
    return spreadsheet.url


if __name__ == "__main__":
    run_ops_processing()