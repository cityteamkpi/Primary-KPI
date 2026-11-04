import io
import sys
import time
import socket
import warnings
import gspread
import openpyxl
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from googleapiclient.http import MediaFileUpload
from concurrent.futures import ThreadPoolExecutor
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, find_file_id

socket.setdefaulttimeout(120)
warnings.filterwarnings("ignore", category=UserWarning)

# ============================================================
#  GLOBAL CONSTANTS & LISTS
# ============================================================
INTERN_PROGRAM_NAMES = {
    "program graduate intern",
    "leadership module",
    "cdac trainee",
}

INTERN_AND_FORWARD_PROGRAM_NAMES = {
    "program graduate intern",
    "leadership module",
    "cdac trainee",
    "forward program",
    "forward",
}

GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]
MONTHLY_HOURS = 173.33

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
    "Traffic Control Date_6786", "Certificate/License Date"
]

DESTINATION_SHEET_ID = "1bvcXczlUgCpdY5jPJopjJ6lT5J03g1qrGo9zQgc93Eg"

FY_CONFIGS = {
    "FY26": {
        "start": pd.Timestamp("2025-09-01"),
        "end": pd.Timestamp("2026-08-31"),
        "sheet_name": "FY2026 Shell"
    },
    "FY27": {
        "start": pd.Timestamp("2026-09-01"),
        "end": pd.Timestamp.now().normalize(),  # FY27 To-Date
        "sheet_name": "Current Fiscal Year To-Date"
    }
}

FILE_OCCUPANCY_NAME   = "Clients in All Programs Average"
FILE_RENEW_NAME       = "Renew Programs Report.xlsx"
FILE_TP_NAME          = "Turning Point Report.xlsx"
FILE_BAPTISMS_NAME    = "CityTeam Baptisms.xlsx"
FILE_MENTOR_NAME      = "Mentorship.xlsx"
FILE_REUNIF_NAME      = "Family Reunifications.xlsx"
FILE_ALUMNI_NAME      = "All Programs All Alumni KPIs.xlsx"
FILE_ATWORK_NAME      = "All Programs @Work Fiscal Year Impacts.xlsx"

COLUMN_INDICES = {
    "San Jose Men": 2,      # Col B
    "San Jose Women": 3,    # Col C
    "Oakland Men": 4,       # Col D
    "Oakland Women": 5,     # Col E
    "Portland Men": 6,      # Col F
    "Portland Women": 7,    # Col G
    "Chester Men": 8,       # Col H
    "Chester Women": 9      # Col I
}

PROGRAM_MAP = {
    "San Jose Men": {
        "renew": ["San Jose Men Renew"],
        "tp": ["San Jose Men Turning Point"],
        "fwd_intern": ["GV Men's Forward"]
    },
    "San Jose Women": {
        "renew": ["GV Renew", "House of Grace"],
        "tp": ["GV Turning Point", "Heritage Home", "San Jose Youth Collective"],
        "fwd_intern": ["GV Women's Forward", "San Jose Women's House of Light"]
    },
    "Oakland Men": {
        "renew": ["Oakland Renew"],
        "tp": ["Oakland Men Turning Point"],
        "fwd_intern": ["Oakland Forward Men", "Oakland Forwad Men"]
    },
    "Oakland Women": {
        "renew": [],
        "tp": ["Oakland Women Turning Point"],
        "fwd_intern": ["Oakland Forward Women"]
    },
    "Portland Men": {
        "renew": ["Portland Renew"],
        "tp": ["Portland Men Turning Point"],
        "fwd_intern": ["Portland Forward"]
    },
    "Portland Women": {
        "renew": [],
        "tp": ["Portland Community of Hope"],
        "fwd_intern": []
    },
    "Chester Men": {
        "renew": ["Chester Renew"],
        "tp": ["Chester Men Turning Point"],
        "fwd_intern": ["Chester Forward"]
    },
    "Chester Women": {
        "renew": [],
        "tp": ["Chester Women Turning Point"],
        "fwd_intern": []
    }
}

ROW_RES_NEW_ADULT_ENROLLMENTS      = 5
ROW_RES_NEW_CHILD_ENROLLMENTS      = 6
ROW_RENEW_ADULT_AVG_OCCUPANCY      = 7
ROW_RENEW_CHILD_AVG_OCCUPANCY      = 8
ROW_TP_ADULT_AVG_OCCUPANCY         = 9
ROW_TP_CHILD_AVG_OCCUPANCY         = 10
ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY = 11
ROW_FWD_INTERN_CHILD_AVG_OCCUPANCY = 12

ROW_RENEW_ADULT_NIGHTS             = 14
ROW_TP_ADULT_NIGHTS                = 15
ROW_FWD_INTERN_NIGHTS              = 16

ROW_RENEW_GRADUATIONS              = 18
ROW_TP_GRADUATIONS                 = 19
ROW_TP_SUCCESSFULLY_HOUSED         = 20

ROW_FAMILIES_REUNIFIED             = 22
ROW_SPIRITUAL_MENTORSHIPS          = 23
ROW_BAPTISMS                       = 24
ROW_SOBRIETY_RATE                  = 25
ROW_CASE_MANAGEMENT_HOURS          = 26

ROW_JOBS_ATTAINED                  = 28
ROW_CREDENTIALS_ATTAINED           = 29
ROW_IDS_CARDS_DOCS                 = 30
ROW_DRIVERS_LICENSES               = 31
ROW_HSE_ATTAINED                   = 32
ROW_BANK_ACCOUNTS                  = 33
ROW_HOUSING_BURDEN_1YR             = 34
ROW_WAGE_INCREASE                  = 35

ROWS_KPI_SPEC = [
    ("Row 5", "Res. New Adult Enrollments", "Renew Programs Report / Turning Point Report", "Count of unique Record IDs starting Renew/TP in FY"),
    ("Row 6", "Res. New Child Enrollments", "Renew Programs Report / Turning Point Report", "Sum of 'Number of Children' for Renew/TP enrollments in FY"),
    ("Row 7", "Renew Adult Avg Occupancy", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Count' for Category=='Renew' programs (FY27)"),
    ("Row 8", "Renew Child Avg Occupancy", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Children Count' for Category=='Renew' programs (FY27)"),
    ("Row 9", "Turning Point Adult Avg Occupancy", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Count' for Category=='Men/Women Turning Point' (FY27)"),
    ("Row 10", "Turning Point Child Avg Occupancy", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Children Count' for TP Category (FY27)"),
    ("Row 11", "Forward - Intern Adult Avg Occ", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Count' for Category in ['Forward', 'Intern'] (FY27)"),
    ("Row 12", "Forward - Intern Child Avg Occ", "Clients in All Programs Average ('Latest Occupancy')", "Sum of 'Children Count' for Forward/Intern Category (FY27)"),
    ("Row 14", "Renew Adult Nights of Shelter", "Renew Programs Report.xlsx", "Sum of participant stay days overlapping FY window"),
    ("Row 15", "Turning Point Adult Nights", "Turning Point Report.xlsx", "Sum of participant stay days overlapping FY window"),
    ("Row 16", "Forward - Intern Nights", "Renew Programs Report.xlsx", "Sum of stay days for Forward/Intern variations overlapping FY"),
    ("Row 18", "Renew Graduations", "Renew Programs Report.xlsx", "Count of unique graduates with primary reason in Graduation list"),
    ("Row 19", "Turning Point Graduations", "Turning Point Report.xlsx", "Count of unique graduates with reason 'Graduation'"),
    ("Row 20", "Turning Point Housed", "Turning Point Report.xlsx", "Count of unique clients with Exit Date in FY & Housed=='Yes'"),
    ("Row 22", "Families Reunified", "Family Reunifications.xlsx", "Count of unique client reunifications occurring in FY"),
    ("Row 23", "Spiritual Mentorships", "Mentorship.xlsx / Latest Mentorship", "FY26: Raw start dates in FY; FY27: Count from 'Latest Mentorship' tab"),
    ("Row 24", "Baptisms", "CityTeam Baptisms.xlsx", "Count of unique client baptism dates occurring in FY"),
    ("Row 25", "Sobriety Rate", "All Programs All Alumni KPIs.xlsx", "Proportion of prior FY Renew graduates sustaining post-grad sobriety"),
    ("Row 26", "Case Management Hours", "Renew / Turning Point Reports", "Sum of stay weeks x location multiplier (2.0 Chester, 1.0 others)"),
    ("Row 28", "Jobs Attained", "All Programs @Work Fiscal Year Impacts.xlsx", "Count of unique clients acquiring job in FY"),
    ("Row 29", "Credentials Attained", "All Programs @Work Fiscal Year Impacts.xlsx", "Total count of 35 industry cert dates earned in FY"),
    ("Row 30", "Essential Documents", "All Programs @Work Fiscal Year Impacts.xlsx", "Combined count of State ID, Birth Cert, and SS Card dates in FY"),
    ("Row 31", "Driver's Licenses", "All Programs @Work Fiscal Year Impacts.xlsx", "Count of unique driver's licenses obtained in FY"),
    ("Row 32", "High School Equivalency (HSE)", "All Programs @Work Fiscal Year Impacts.xlsx", "Count of unique HSE dates obtained in FY"),
    ("Row 33", "Bank Accounts Obtained", "All Programs @Work Fiscal Year Impacts.xlsx", "Count of unique bank account dates obtained in FY"),
    ("Row 34", "Housing Burden Under 30%", "All Programs All Alumni KPIs.xlsx", "Proportion of prior FY Renew graduates with housing cost <= 30% of income"),
    ("Row 35", "Hourly Wage Increase", "All Programs All Alumni KPIs.xlsx", "Average wage difference ($/hr) from graduation to recent check-in for Renew grads")
]

def log(msg):
    print(msg, flush=True)

# ============================================================
#  HELPER & REMAPPING FUNCTIONS
# ============================================================
def clean_df_columns(df):
    """Strips leading/trailing whitespaces from DataFrame column names to prevent KeyErrors."""
    if not df.empty:
        df.columns = df.columns.astype(str).str.strip()
    return df

def get_intern_program_value(row, col_map=None):
    if col_map:
        col_pg = next((col_map[c] for c in col_map if "post graduate program association" in c), None)
        col_intern = next((col_map[c] for c in col_map if "intern program" in c), None)
        val_pg = str(row[col_pg]).strip() if col_pg and col_pg in row and pd.notna(row[col_pg]) else ""
        val_intern = str(row[col_intern]).strip() if col_intern and col_intern in row and pd.notna(row[col_intern]) else ""
    else:
        post_grad = row.get("Post Graduate Program Association_6619") or row.get("Post Graduate Program Association")
        intern_prog = row.get("Intern Program_6619") or row.get("Intern Program")
        val_pg = str(post_grad).strip() if pd.notna(post_grad) else ""
        val_intern = str(intern_prog).strip() if pd.notna(intern_prog) else ""

    invalid_tokens = {"", "nan", "none", "nat", "<na>"}
    pg_valid = val_pg if val_pg.lower() not in invalid_tokens else None
    intern_valid = val_intern if val_intern.lower() not in invalid_tokens else None

    if pg_valid:
        return pg_valid
    return intern_valid if intern_valid else ""

def is_matching_program(row, prog_col, target_progs_lower, col_map):
    p_val = str(row[prog_col]).strip().lower() if pd.notna(row[prog_col]) else ""
    if p_val in target_progs_lower:
        return True
    if p_val in INTERN_PROGRAM_NAMES:
        secondary_val = get_intern_program_value(row, col_map).lower()
        if secondary_val in target_progs_lower:
            return True
    return False

def clean_wage(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    clean = str(val).replace("$", "").replace(",", "").strip()
    try: return float(clean)
    except: return 0.0

def combine_columns_with_priority(df, col_new, col_old, target_col, win_start, win_end):
    invalid_tokens = ["", "nan", "none", "nat", "<na>"]

    s_new = df[col_new].astype(str).str.strip().str.lower() if col_new in df.columns else pd.Series("", index=df.index)
    s_new_clean = df[col_new].where(~s_new.isin(invalid_tokens), np.nan) if col_new in df.columns else pd.Series(np.nan, index=df.index)
    dt_new = pd.to_datetime(s_new_clean, errors="coerce")
    dt_new_valid = dt_new.where((dt_new >= win_start) & (dt_new <= win_end), pd.NaT)

    s_old = df[col_old].astype(str).str.strip().str.lower() if col_old in df.columns else pd.Series("", index=df.index)
    s_old_clean = df[col_old].where(~s_old.isin(invalid_tokens), np.nan) if col_old in df.columns else pd.Series(np.nan, index=df.index)
    dt_old = pd.to_datetime(s_old_clean, errors="coerce")
    dt_old_valid = dt_old.where((dt_old >= win_start) & (dt_old <= win_end), pd.NaT)

    df[target_col] = dt_new_valid.fillna(dt_old_valid)
    return df

# ============================================================
#  NATIVE LOGIC FUNCTIONS FOR ALL METRICS
# ============================================================

# --- ROWS 5 & 6: ENROLLMENTS ---
def count_residential_enrollments(df, fy_start, fy_end, prog_list):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in ["record id_102", "record id_2086", "record id_152", "record id"] if cand in col_map), None)
    date_col = next((col_map[cand] for cand in ["start date_2090", "start date"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not date_col or not prog_col:
        return 0

    starts = pd.to_datetime(df[date_col], errors="coerce")
    df_fy = df[(starts >= fy_start) & (starts <= fy_end)].copy()
    if df_fy.empty:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    mask = df_fy.apply(lambda row: is_matching_program(row, prog_col, prog_set_lower, col_map), axis=1)
    df_res = df_fy[mask]

    if df_res.empty:
        return 0

    if id_col and id_col in df_res.columns:
        return int(df_res[id_col].dropna().nunique())

    return len(df_res)

# --- ROWS 14–16: NIGHTS OF SHELTER ---
def calculate_nights_of_shelter(df, prog_list, fy_start, fy_end):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in ["record id_102", "record id_2086", "record id"] if cand in col_map), None)
    start_col = next((col_map[cand] for cand in ["start date_2090", "start date"] if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in ["exit date_2100", "exit date"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not start_col or not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    mask = df.apply(lambda row: is_matching_program(row, prog_col, prog_set_lower, col_map), axis=1)
    df_sub = df[mask].copy()
    if df_sub.empty:
        return 0

    if id_col and exit_col:
        df_sub = df_sub.dropna(subset=[id_col, start_col, prog_col]).drop_duplicates(subset=[id_col, start_col, prog_col])

    today = pd.Timestamp.now().normalize()
    effective_end_limit = min(today, fy_end)

    starts = pd.to_datetime(df_sub[start_col], errors="coerce")
    exits = pd.to_datetime(df_sub[exit_col], errors="coerce").fillna(today) if exit_col else pd.Series(today, index=df_sub.index)

    overlap_start = starts.clip(lower=fy_start)
    overlap_end = exits.clip(upper=effective_end_limit)

    nights = (overlap_end - overlap_start).dt.days
    nights = nights.where((nights > 0) & (starts <= effective_end_limit) & (exits >= fy_start), 0)

    return int(nights.sum())

# --- ROWS 18 & 19: NATIVE GRADUATIONS ---
def count_native_graduations(df, fy_start, fy_end, prog_list, valid_reasons):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in ["record id_102", "record id_2086", "record id"] if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in ["exit date_2100", "exit date"] if cand in col_map), None)
    reason_col = next((col_map[cand] for cand in ["primary reason for exit_2102", "primary reason for exit", "exit reason"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not exit_col or not prog_col or not reason_col:
        return 0

    exits = pd.to_datetime(df[exit_col], errors="coerce")
    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    reasons_lower = {str(r).strip().lower() for r in valid_reasons}

    mask = (
        (exits >= fy_start) & 
        (exits <= fy_end) & 
        df.apply(lambda row: is_matching_program(row, prog_col, prog_set_lower, col_map), axis=1) &
        (df[reason_col].astype(str).str.strip().str.lower().isin(reasons_lower))
    )
    df_sub = df[mask].copy()

    if df_sub.empty:
        return 0

    if id_col and id_col in df_sub.columns:
        return int(df_sub[id_col].dropna().nunique())

    return len(df_sub)

# --- ROW 20: TURNING POINT SUCCESSFULLY HOUSED ---
def count_native_tp_housed(df_tp_raw, fy_start, fy_end, tp_progs):
    if df_tp_raw.empty or not tp_progs:
        return 0

    df = df_tp_raw.copy()
    col_map = {str(c).lower().strip(): c for c in df.columns}

    id_col = next((col_map[cand] for cand in ["record id_102", "record id"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in ["exit date_2100", "exit date"] if cand in col_map), None)
    h_new_col = next((col_map[cand] for cand in ["successfully housed_7011"] if cand in col_map), None)
    h_old_col = next((col_map[cand] for cand in ["successfully housed (is housing healthy?)_4368"] if cand in col_map), None)

    if not prog_col or not exit_col:
        return 0

    def eval_housed(row):
        val_n = str(row.get(h_new_col, '')).strip().lower() if h_new_col else ''
        if val_n in ['yes', '1', 'true']:
            return 1
        val_o = str(row.get(h_old_col, '')).strip().lower() if h_old_col else ''
        if val_o in ['yes', '1', 'true']:
            return 1
        return 0

    df["_is_housed"] = df.apply(eval_housed, axis=1)
    df_h = df[df["_is_housed"] == 1].copy()
    if df_h.empty:
        return 0

    tp_lower = {str(p).strip().lower() for p in tp_progs}
    mask = df_h.apply(lambda row: is_matching_program(row, prog_col, tp_lower, col_map), axis=1)
    df_sub = df_h[mask].copy()
    if df_sub.empty:
        return 0

    exits = pd.to_datetime(df_sub[exit_col], errors="coerce")
    df_valid = df_sub[(exits >= fy_start) & (exits <= fy_end)]
    if df_valid.empty:
        return 0

    if id_col and id_col in df_valid.columns:
        return int(df_valid[id_col].dropna().nunique())

    return len(df_valid)

# --- ROW 23: SPIRITUAL MENTORSHIP (FY27 FROM LATEST MENTORSHIP TAB) ---
def count_latest_mentorship_occ(df_ment_occ, all_progs):
    if df_ment_occ.empty or not all_progs:
        return 0

    df = df_ment_occ.copy()
    if "Program" not in df.columns or "Count" not in df.columns:
        return 0

    progs_lower = {str(p).strip().lower() for p in all_progs}
    df_sub = df[df["Program"].astype(str).str.strip().str.lower().isin(progs_lower)]
    
    if df_sub.empty:
        return 0

    return int(pd.to_numeric(df_sub["Count"], errors="coerce").fillna(0).sum())

# --- ROWS 22–24: NATIVE SPIRITUAL METRICS ---
def count_native_spiritual_metric(df_raw, metric_label, fy_start, fy_end, all_progs):
    if df_raw.empty or not all_progs:
        return 0

    df = df_raw.copy()
    col_map = {str(c).lower().strip(): c for c in df.columns}

    id_col = next((col_map[cand] for cand in ["record id_102", "record id"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not prog_col:
        return 0

    if metric_label == "Mentorship":
        date_new = col_map.get("start date for mentoring_7199")
        date_old = col_map.get("start date for mentoring_4365")
        dt_new = pd.to_datetime(df[date_new], errors="coerce") if date_new else pd.Series(pd.NaT, index=df.index)
        dt_old = pd.to_datetime(df[date_old], errors="coerce") if date_old else pd.Series(pd.NaT, index=df.index)
        df["_event_date"] = dt_new.fillna(dt_old)

        flag_new = col_map.get("acquired spiritual mentor_7196")
        flag_old = col_map.get("acquired spiritual mentor_5136")
        v_new = df[flag_new].astype(str).str.strip() if flag_new else pd.Series("", index=df.index)
        v_old = df[flag_old].astype(str).str.strip() if flag_old else pd.Series("", index=df.index)
        df = df[(v_new == "Yes") | (v_old == "Yes")].copy()

    elif metric_label == "Reunifications":
        date_new = col_map.get("date of reunification_7018")
        date_old = col_map.get("date of reunification_5140")
        dt_new = pd.to_datetime(df[date_new], errors="coerce") if date_new else pd.Series(pd.NaT, index=df.index)
        dt_old = pd.to_datetime(df[date_old], errors="coerce") if date_old else pd.Series(pd.NaT, index=df.index)
        df["_event_date"] = dt_new.fillna(dt_old)

    elif metric_label == "Baptisms":
        date_new = col_map.get("date of baptism_7016")
        date_old = col_map.get("baptism date_4378")
        dt_new = pd.to_datetime(df[date_new], errors="coerce") if date_new else pd.Series(pd.NaT, index=df.index)
        dt_old = pd.to_datetime(df[date_old], errors="coerce") if date_old else pd.Series(pd.NaT, index=df.index)
        df["_event_date"] = dt_new.fillna(dt_old)

    if df.empty:
        return 0

    all_lower = {str(p).strip().lower() for p in all_progs}
    mask = df.apply(lambda row: is_matching_program(row, prog_col, all_lower, col_map), axis=1)
    df_sub = df[mask].copy()
    if df_sub.empty:
        return 0

    dates = pd.to_datetime(df_sub["_event_date"], errors="coerce")
    df_valid = df_sub[(dates >= fy_start) & (dates <= fy_end)]
    if df_valid.empty:
        return 0

    if id_col and id_col in df_valid.columns:
        return int(df_valid[id_col].dropna().nunique())

    return len(df_valid)

# --- ROW 26: CASE MANAGEMENT HOURS ---
def calculate_case_management_hours(df, prog_list, fy_start, fy_end, is_chester=False):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in ["record id_102", "record id_2086", "record id"] if cand in col_map), None)
    start_col = next((col_map[cand] for cand in ["start date_2090", "start date"] if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in ["exit date_2100", "exit date"] if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not start_col or not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    mask = df.apply(lambda row: is_matching_program(row, prog_col, prog_set_lower, col_map), axis=1)
    df_sub = df[mask].copy()
    if df_sub.empty:
        return 0

    if id_col and exit_col:
        df_sub = df_sub.dropna(subset=[id_col, start_col, prog_col]).drop_duplicates(subset=[id_col, start_col, prog_col])

    today = pd.Timestamp.now().normalize()
    effective_end_limit = min(today, fy_end)

    starts = pd.to_datetime(df_sub[start_col], errors="coerce")
    exits = pd.to_datetime(df_sub[exit_col], errors="coerce").fillna(today) if exit_col else pd.Series(today, index=df_sub.index)

    overlap_start = starts.clip(lower=fy_start)
    overlap_end = exits.clip(upper=effective_end_limit)

    days = (overlap_end - overlap_start).dt.days
    days = days.where((days > 0) & (starts <= effective_end_limit) & (exits >= fy_start), 0)

    weeks = days / 7.0
    multiplier = 2.0 if is_chester else 1.0

    return int(round((weeks * multiplier).sum()))

# --- ROWS 28–33: NATIVE @WORK METRICS ---
def process_atwork_raw_to_kpis(df_raw, fy_start, fy_end):
    if df_raw.empty:
        return pd.DataFrame()

    df = df_raw.copy()
    col_map = {str(c).lower().strip(): c for c in df.columns}
    
    actual_id_col = next((col_map[cand] for cand in ["record id_152", "record id_102", "record id_2086", "record id"] if cand in col_map), None)
    actual_prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not actual_id_col or not actual_prog_col:
        return df_raw

    col_job_new = next((col_map[c] for c in col_map if "7009" in c), None)
    col_job_old = next((col_map[c] for c in col_map if "5590" in c), None)

    if col_job_new or col_job_old:
        df = combine_columns_with_priority(
            df,
            col_new=col_job_new if col_job_new else "_dummy_new",
            col_old=col_job_old if col_job_old else "_dummy_old",
            target_col="Date Job Acquired",
            win_start=fy_start,
            win_end=fy_end
        )

    date_cols = [c for c in df.columns if "date" in str(c).lower()]
    non_date_cols = [c for c in df.columns if c not in date_cols and c != actual_id_col]

    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    agg_dict = {c: "first" for c in non_date_cols}
    agg_dict.update({c: "max" for c in date_cols})

    df_consolidated = df.groupby(actual_id_col, as_index=False, sort=False).agg(agg_dict)
    return df_consolidated

def count_atwork_dates_from_raw(df_raw, date_col_names, prog_list, fy_start, fy_end, event_based=False):
    if df_raw.empty or not prog_list:
        return 0

    df_sub_consolidated = process_atwork_raw_to_kpis(df_raw, fy_start, fy_end)
    if df_sub_consolidated.empty:
        return 0

    col_map = {str(c).lower().strip(): c for c in df_sub_consolidated.columns}
    prog_col = next((col_map[cand] for cand in ["program enrolling_2091", "program enrolling", "program"] if cand in col_map), None)

    if not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    mask = df_sub_consolidated.apply(lambda row: is_matching_program(row, prog_col, prog_set_lower, col_map), axis=1)
    df_sub = df_sub_consolidated[mask].copy()
    if df_sub.empty:
        return 0

    if isinstance(date_col_names, str):
        date_col_names = [date_col_names]

    matched_cols = [col_map[c.lower().strip()] for c in date_col_names if c.lower().strip() in col_map]
    if not matched_cols:
        return 0

    if event_based:
        total_events = 0
        for col in matched_cols:
            dates = pd.to_datetime(df_sub[col], errors="coerce")
            total_events += int(((dates >= fy_start) & (dates <= fy_end)).sum())
        return total_events

    in_window_mask = pd.Series(False, index=df_sub.index)
    for col in matched_cols:
        dates = pd.to_datetime(df_sub[col], errors="coerce")
        in_window_mask |= ((dates >= fy_start) & (dates <= fy_end))

    df_valid = df_sub[in_window_mask]
    actual_id_col = next((col_map[cand] for cand in ["record id_152", "record id_102", "record id_2086", "record id"] if cand in col_map), None)

    if actual_id_col and actual_id_col in df_valid.columns:
        return int(df_valid[actual_id_col].dropna().nunique())

    return len(df_valid)

# --- ROWS 25 & 34–35: NATIVE ALUMNI METRICS (RENEW ONLY) ---
def compute_native_alumni_metrics(df_alumni_raw, target_fy_str, renew_progs):
    if df_alumni_raw.empty or not renew_progs:
        return {"sobriety": "", "housing": "", "wage": ""}

    df = clean_df_columns(df_alumni_raw.copy())
    col_map = {str(c).lower().strip(): c for c in df.columns}

    name_col = next((col_map[c] for c in col_map if "name" in c), None)
    exit_col = next((col_map[c] for c in col_map if "exit date" in c), None)
    prog_col = next((col_map[c] for c in col_map if "program enrolling" in c or "program" in c), None)
    reason_col = next((col_map[c] for c in col_map if "primary reason for exit" in c or "exit reason" in c), None)

    if not name_col or not exit_col or not prog_col or not reason_col:
        return {"sobriety": "", "housing": "", "wage": ""}

    df = df.dropna(subset=[name_col]).copy()

    checkin_col = next((col_map[c] for c in col_map if "checkin date" in c or "check-in date" in c), None)
    wage_col = next((col_map[c] for c in col_map if "hourly wage" in c or "wage" in c), None)
    housing_col = next((col_map[c] for c in col_map if "housing expense" in c or "housing" in c), None)
    sober_col = next((col_map[c] for c in col_map if "sober?" in c or "sober" in c), None)
    start_col = next((col_map[c] for c in col_map if "start date" in c), None)

    df[exit_col] = pd.to_datetime(df[exit_col], errors="coerce")
    if checkin_col and checkin_col in df.columns:
        df[checkin_col] = pd.to_datetime(df[checkin_col], errors="coerce")
    if start_col and start_col in df.columns:
        df[start_col] = pd.to_datetime(df[start_col], errors="coerce")

    if wage_col and wage_col in df.columns:
        df[wage_col] = df[wage_col].apply(clean_wage)
    if housing_col and housing_col in df.columns:
        df[housing_col] = df[housing_col].apply(clean_wage)

    renew_lower = {str(p).strip().lower() for p in renew_progs}

    df_grads = df[
        df[prog_col].astype(str).str.strip().str.lower().isin(renew_lower) &
        df[reason_col].astype(str).str.strip().isin(GRAD_REASONS) &
        df[exit_col].notna()
    ].copy()

    def get_fiscal_year_label(dt):
        if pd.isna(dt): return ""
        fy = dt.year + 1 if dt.month >= 9 else dt.year
        return f"FY{str(fy)[-2:]}"

    df_grads["FY_Label"] = df_grads[exit_col].apply(get_fiscal_year_label)
    df_grads_target = df_grads[df_grads["FY_Label"] == target_fy_str].copy()

    if df_grads_target.empty:
        return {"sobriety": "", "housing": "", "wage": ""}

    df_grads_dedup = df_grads_target.sort_values(by=[exit_col, reason_col]).drop_duplicates(subset=[name_col, prog_col, exit_col], keep="last")

    sobriety_vals = []
    housing_under_30_vals = []
    wage_diff_vals = []

    for _, row in df_grads_dedup.iterrows():
        c_name = row[name_col]
        c_prog = row[prog_col]
        g_date = row[exit_col]

        person_records = df[(df[name_col] == c_name) & (df[prog_col] == c_prog)]
        if checkin_col and checkin_col in person_records.columns:
            person_records = person_records.sort_values(by=checkin_col)
            if start_col and start_col in person_records.columns:
                person_records = person_records[(person_records[start_col].isna()) | (person_records[start_col] <= g_date) | (person_records[checkin_col] < person_records[start_col])]

        sober_recs = person_records[person_records[sober_col].isin(["Yes", "No"]) & (person_records[checkin_col] > g_date)] if (sober_col and checkin_col and sober_col in person_records.columns and checkin_col in person_records.columns) else pd.DataFrame()
        has_sustained_relapse = False
        first_no_date = None
        last_checkin_sober = False

        for _, s_row in sober_recs.iterrows():
            c_state = s_row[sober_col]
            c_date = s_row[checkin_col]
            if c_state == "No":
                if first_no_date is None: first_no_date = c_date
            elif c_state == "Yes":
                if first_no_date is not None:
                    if (c_date - first_no_date).days > 30: has_sustained_relapse = True
                    first_no_date = None
                last_checkin_sober = True

        is_sober = True if sober_recs.empty else (last_checkin_sober and not has_sustained_relapse)
        sobriety_vals.append(1 if is_sober else 0)

        grad_wage = float(row[wage_col]) if (wage_col and wage_col in row and pd.notna(row[wage_col])) else 0.0
        grad_housing = float(row[housing_col]) if (housing_col and housing_col in row and pd.notna(row[housing_col])) else 0.0

        post_recs = person_records[person_records[checkin_col].notna() & (person_records[checkin_col] > g_date)] if (checkin_col and checkin_col in person_records.columns) else pd.DataFrame()
        if not post_recs.empty:
            latest_checkin = post_recs[checkin_col].max()
            latest_row = post_recs[post_recs[checkin_col] == latest_checkin].iloc[0]
            recent_wage = float(latest_row[wage_col]) if (wage_col and wage_col in latest_row and pd.notna(latest_row[wage_col])) else grad_wage
            recent_housing = float(latest_row[housing_col]) if (housing_col and housing_col in latest_row and pd.notna(latest_row[housing_col])) else grad_housing
        else:
            recent_wage = grad_wage
            recent_housing = grad_housing

        if is_sober:
            recent_income = recent_wage * MONTHLY_HOURS
            if recent_housing == 0:
                h_under_30 = 1
            elif recent_income > 0:
                h_under_30 = 1 if (recent_housing / recent_income) <= 0.30 else 0
            else:
                h_under_30 = 0
            housing_under_30_vals.append(h_under_30)

            wage_diff_vals.append(recent_wage - grad_wage)

    val_r25 = float(np.mean(sobriety_vals)) if sobriety_vals else ""
    val_r34 = float(np.mean(housing_under_30_vals)) if housing_under_30_vals else ""
    val_r35 = float(np.mean(wage_diff_vals)) if wage_diff_vals else ""

    return {
        "sobriety": round(val_r25, 4) if isinstance(val_r25, float) else "",
        "housing": round(val_r34, 4) if isinstance(val_r34, float) else "",
        "wage": round(val_r35, 2) if isinstance(val_r35, float) else ""
    }

# ============================================================
#  AUDIT SPREADSHEET GENERATION & UPLOAD ENGINE
# ============================================================
def sanitize_sheet_title(r_num, name):
    clean_name = name.replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "").replace(":", "")
    title = f"{r_num} - {clean_name}"
    return title[:31]

def generate_and_upload_audit_spreadsheet(main_drive_service, target_folder_id):
    log("\n📊 Generating 'Marketing Data Impact Spreadsheet.xlsx' workbook...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_subhdr = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    thin_border_side = Side(style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # 1. Add Summary Overview Tab
    ws_summary = wb.create_sheet(title="Audit Summary & Index")
    ws_summary.views.sheetView[0].showGridLines = True
    ws_summary.append(["CityTeam Marketing Impact Data - Tab Audit Index"])
    ws_summary.cell(1, 1).font = Font(name="Calibri", size=14, bold=True, color="1F4E78")

    headers_summary = ["Destination Row", "Metric Name", "Primary Source Report File", "Audit Tab Name", "Calculation Logic & Filter Rule"]
    ws_summary.append(headers_summary)

    for col_idx in range(1, len(headers_summary) + 1):
        cell = ws_summary.cell(row=2, column=col_idx)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center

    for idx, (r_num, name, src, logic) in enumerate(ROWS_KPI_SPEC, start=3):
        tab_title = sanitize_sheet_title(r_num, name)
        ws_summary.append([r_num, name, src, tab_title, logic])
        for col_idx in range(1, 6):
            c = ws_summary.cell(row=idx, column=col_idx)
            c.font = font_regular
            c.border = thin_border
            if idx % 2 == 0:
                c.fill = fill_zebra
            if col_idx in [1, 4]:
                c.alignment = align_center
                if col_idx == 1:
                    c.font = font_bold
            else:
                c.alignment = align_left

    # 2. Add individual tabs for each destination row
    for r_num, name, src, logic in ROWS_KPI_SPEC:
        tab_title = sanitize_sheet_title(r_num, name)
        ws = wb.create_sheet(title=tab_title)
        ws.views.sheetView[0].showGridLines = True

        ws.cell(row=1, column=1, value="Metric Name:").font = font_bold
        ws.cell(row=1, column=2, value=name).font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        ws.cell(row=2, column=1, value="Destination Row:").font = font_bold
        ws.cell(row=2, column=2, value=r_num).font = font_regular
        ws.cell(row=3, column=1, value="Source File:").font = font_bold
        ws.cell(row=3, column=2, value=src).font = font_regular
        ws.cell(row=4, column=1, value="Calculation Rule:").font = font_bold
        ws.cell(row=4, column=2, value=logic).font = font_regular

        headers = [
            "Record ID", "Client Name", "Program Name", 
            "Location / Division", "Start Date", "Exit / Event Date", 
            "Audit Status (FY26)", "Audit Status (FY27)", 
            "San Jose Men", "San Jose Women", "Oakland Men", "Oakland Women",
            "Portland Men", "Portland Women", "Chester Men", "Chester Women",
            "Notes / Validation Formula"
        ]
        ws.append([])
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=6, column=col_idx)
            c.font = font_header
            c.fill = fill_subhdr
            c.alignment = align_center
            c.border = thin_border

        sample_rows = [
            ["1001", "Client Sample A", "San Jose Men Renew", "San Jose Men", "2025-09-15", "2026-03-10", "Included FY26", "Excluded", 1, 0, 0, 0, 0, 0, 0, 0, "Verified within FY26 window"],
            ["1002", "Client Sample B", "Oakland Men Turning Point", "Oakland Men", "2026-01-10", "2026-07-20", "Included FY26", "Excluded", 0, 0, 1, 0, 0, 0, 0, 0, "Verified within FY26 window"],
            ["1003", "Client Sample C", "GV Renew", "San Jose Women", "2026-09-05", "2027-02-14", "Excluded", "Included FY27", 0, 1, 0, 0, 0, 0, 0, 0, "Verified within FY27 To-Date window"]
        ]
        for r_idx, row_data in enumerate(sample_rows, start=7):
            ws.append(row_data)
            for col_idx in range(1, len(row_data) + 1):
                c = ws.cell(row=r_idx, column=col_idx)
                c.font = font_regular
                c.border = thin_border
                if col_idx in [1, 5, 6, 7, 8]:
                    c.alignment = align_center
                elif col_idx >= 9 and col_idx <= 16:
                    c.alignment = align_right
                else:
                    c.alignment = align_left

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row <= 4 and col_letter == 'B':
                    continue
                max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    local_filename = "Marketing_Data_Impact_Spreadsheet.xlsx"
    wb.save(local_filename)
    log(f"   ✓ Local audit workbook saved as '{local_filename}'.")

    # Upload to target Drive folder
    log("🚀 Uploading audit workbook to 'Quarterly KPI Processed Data' folder...")
    file_metadata = {
        'name': 'Marketing Data Impact Spreadsheet.xlsx',
        'parents': [target_folder_id]
    }
    media = MediaFileUpload(local_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    existing_file = find_file_id(main_drive_service, 'Marketing Data Impact Spreadsheet.xlsx', target_folder_id)
    if existing_file:
        updated_file = main_drive_service.files().update(
            fileId=existing_file['id'],
            media_body=media,
            fields='id, name',
            supportsAllDrives=True
        ).execute()
        log(f"   ✅ Updated existing audit spreadsheet in Drive (ID: {updated_file.get('id')})!")
    else:
        uploaded_file = main_drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name',
            supportsAllDrives=True
        ).execute()
        log(f"   ✅ Created new audit spreadsheet in Drive (ID: {uploaded_file.get('id')})!")

# ============================================================
#  EXECUTION ENGINE & DRIVE INTEGRATION WITH RETRY LOGIC
# ============================================================
def thread_safe_download(task):
    key, fname, folder_id = task
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            drive_service, _, _ = get_services()
            fh, file_id, _ = download_drive_file(drive_service, fname, folder_id)
            return key, fh, file_id
        except Exception as e:
            if attempt == max_retries:
                raise e
            log(f"   ⚠️ Network error downloading '{fname}' (Attempt {attempt}/{max_retries}). Retrying in 2s...")
            time.sleep(2)

def process_fiscal_year(fy_key, fy_info, gc, dfs):
    fy_start = fy_info["start"]
    fy_end = fy_info["end"]
    sheet_name = fy_info["sheet_name"]

    log(f"\n📊 Processing Metrics for {fy_key} ({fy_start.strftime('%Y-%m-%d')} to {fy_end.strftime('%Y-%m-%d')})...")
    log(f"   Connecting to worksheet: '{sheet_name}'...")
    doc = gc.open_by_key(DESTINATION_SHEET_ID)
    ws = doc.worksheet(sheet_name)

    cell_updates = []
    fy_num = int(fy_key.replace("FY", ""))
    target_grad_fy = f"FY{fy_num - 1}"

    for loc_name, col_idx in COLUMN_INDICES.items():
        renew_progs = PROGRAM_MAP[loc_name]["renew"]
        tp_progs = PROGRAM_MAP[loc_name]["tp"]
        fwd_intern_progs = PROGRAM_MAP[loc_name]["fwd_intern"]

        fwd_intern_all = fwd_intern_progs
        all_progs = renew_progs + tp_progs + fwd_intern_all
        res_progs = renew_progs + tp_progs

        df_res_raw_combined = pd.concat([dfs["renew_raw"], dfs["tp_raw"]], ignore_index=True)

        # ----------------------------------------------------
        # ROWS 5-6: RESIDENTIAL ENROLLMENTS (RENEW + TURNING POINT)
        # ----------------------------------------------------
        val_r5 = count_residential_enrollments(df_res_raw_combined, fy_start, fy_end, res_progs)
        cell_updates.append({"row": ROW_RES_NEW_ADULT_ENROLLMENTS, "col": col_idx, "val": val_r5})

        val_r6 = 0
        col_map_res = {str(c).lower().strip(): c for c in df_res_raw_combined.columns}
        if "number of children" in col_map_res:
            n_kids_col = col_map_res["number of children"]
            d_col = next((col_map_res[c] for c in ["start date_2090", "start date"] if c in col_map_res), None)
            p_col = next((col_map_res[c] for c in ["program enrolling_2091", "program enrolling"] if c in col_map_res), None)
            if d_col and p_col:
                starts = pd.to_datetime(df_res_raw_combined[d_col], errors="coerce")
                res_progs_lower = {str(p).strip().lower() for p in res_progs}
                mask = df_res_raw_combined.apply(lambda r: is_matching_program(r, p_col, res_progs_lower, col_map_res), axis=1)
                in_fy = df_res_raw_combined[mask & (starts >= fy_start) & (starts <= fy_end)]
                val_r6 = int(in_fy[n_kids_col].sum()) if not in_fy.empty else 0
        cell_updates.append({"row": ROW_RES_NEW_CHILD_ENROLLMENTS, "col": col_idx, "val": val_r6})

        # ----------------------------------------------------
        # ROWS 7-12: AVERAGE OCCUPANCY (FIXED CATEGORY + PROGRAM MATCHING)
        # ----------------------------------------------------
        if fy_key == "FY27":
            renew_lower = {str(p).strip().lower() for p in renew_progs}
            tp_lower = {str(p).strip().lower() for p in tp_progs}
            all_progs_lower = {str(p).strip().lower() for p in all_progs}

            if not dfs["occ"].empty and "Program" in dfs["occ"].columns and "Category" in dfs["occ"].columns:
                occ_df = dfs["occ"]
                occ_prog_series = occ_df["Program"].astype(str).str.strip().str.lower()
                occ_cat_series = occ_df["Category"].astype(str).str.strip()

                # Rows 7 & 8: Renew Occupancy (Category MUST be 'Renew')
                df_renew_occ = occ_df[(occ_cat_series == "Renew") & (occ_prog_series.isin(renew_lower))]
                val_r7 = int(df_renew_occ["Count"].sum()) if not df_renew_occ.empty else 0
                val_r8 = int(df_renew_occ["Children Count"].sum()) if ("Children Count" in df_renew_occ.columns and not df_renew_occ.empty) else 0

                # Rows 9 & 10: Turning Point Occupancy (Category MUST be 'Men Turning Point' or 'Women Turning Point')
                df_tp_occ = occ_df[(occ_cat_series.isin(["Men Turning Point", "Women Turning Point"])) & (occ_prog_series.isin(tp_lower))]
                val_r9 = int(df_tp_occ["Count"].sum()) if not df_tp_occ.empty else 0
                val_r10 = int(df_tp_occ["Children Count"].sum()) if ("Children Count" in df_tp_occ.columns and not df_tp_occ.empty) else 0

                # Rows 11 & 12: Forward + Intern Occupancy (Category MUST be 'Forward' or 'Intern')
                df_fwd_occ = occ_df[(occ_cat_series.isin(["Forward", "Intern"])) & (occ_prog_series.isin(all_progs_lower))]
                val_r11 = int(df_fwd_occ["Count"].sum()) if not df_fwd_occ.empty else 0
                val_r12 = int(df_fwd_occ["Children Count"].sum()) if ("Children Count" in df_fwd_occ.columns and not df_fwd_occ.empty) else 0
            else:
                val_r7 = val_r8 = val_r9 = val_r10 = val_r11 = val_r12 = 0
        else:
            val_r7 = val_r8 = val_r9 = val_r10 = val_r11 = val_r12 = ""

        cell_updates.append({"row": ROW_RENEW_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r7})
        cell_updates.append({"row": ROW_RENEW_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r8})
        cell_updates.append({"row": ROW_TP_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r9})
        cell_updates.append({"row": ROW_TP_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r10})
        cell_updates.append({"row": ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r11})
        cell_updates.append({"row": ROW_FWD_INTERN_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r12})

        # ----------------------------------------------------
        # ROWS 14-16: NIGHTS OF SHELTER
        # ----------------------------------------------------
        val_r14 = calculate_nights_of_shelter(dfs["renew_raw"], renew_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_RENEW_ADULT_NIGHTS, "col": col_idx, "val": val_r14})

        val_r15 = calculate_nights_of_shelter(dfs["tp_raw"], tp_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_TP_ADULT_NIGHTS, "col": col_idx, "val": val_r15})

        val_r16 = calculate_nights_of_shelter(dfs["renew_raw"], fwd_intern_all, fy_start, fy_end)
        cell_updates.append({"row": ROW_FWD_INTERN_NIGHTS, "col": col_idx, "val": val_r16})

        # ----------------------------------------------------
        # ROWS 18-20: GRADUATIONS & HOUSING
        # ----------------------------------------------------
        val_r18 = count_native_graduations(dfs["renew_raw"], fy_start, fy_end, renew_progs, GRAD_REASONS)
        cell_updates.append({"row": ROW_RENEW_GRADUATIONS, "col": col_idx, "val": val_r18})

        val_r19 = count_native_graduations(dfs["tp_raw"], fy_start, fy_end, tp_progs, ["Graduation"])
        cell_updates.append({"row": ROW_TP_GRADUATIONS, "col": col_idx, "val": val_r19})

        val_r20 = count_native_tp_housed(dfs["tp_raw"], fy_start, fy_end, tp_progs)
        cell_updates.append({"row": ROW_TP_SUCCESSFULLY_HOUSED, "col": col_idx, "val": val_r20})

        # ----------------------------------------------------
        # ROWS 22-26: SPIRITUAL & PROGRAM METRICS
        # ----------------------------------------------------
        val_r22 = count_native_spiritual_metric(dfs["reun_raw"], "Reunifications", fy_start, fy_end, all_progs)
        cell_updates.append({"row": ROW_FAMILIES_REUNIFIED, "col": col_idx, "val": val_r22})

        # Row 23: Mentorship (FY27 pulls from 'Latest Mentorship' tab, FY26 pulls from raw dates)
        if fy_key == "FY27":
            val_r23 = count_latest_mentorship_occ(dfs["ment_occ"], all_progs)
        else:
            val_r23 = count_native_spiritual_metric(dfs["ment_raw"], "Mentorship", fy_start, fy_end, all_progs)
        cell_updates.append({"row": ROW_SPIRITUAL_MENTORSHIPS, "col": col_idx, "val": val_r23})

        val_r24 = count_native_spiritual_metric(dfs["bap_raw"], "Baptisms", fy_start, fy_end, all_progs)
        cell_updates.append({"row": ROW_BAPTISMS, "col": col_idx, "val": val_r24})

        # Row 25: Sobriety Rate (Evaluates Renew Program Graduates Only)
        alum_metrics_renew = compute_native_alumni_metrics(dfs["alumni_raw"], target_grad_fy, renew_progs)
        cell_updates.append({"row": ROW_SOBRIETY_RATE, "col": col_idx, "val": alum_metrics_renew["sobriety"]})

        is_chester_loc = "Chester" in loc_name
        val_r26 = calculate_case_management_hours(df_res_raw_combined, res_progs, fy_start, fy_end, is_chester=is_chester_loc)
        cell_updates.append({"row": ROW_CASE_MANAGEMENT_HOURS, "col": col_idx, "val": val_r26})

        # ----------------------------------------------------
        # ROWS 28-35: @WORK & AFTERCARE METRICS
        # ----------------------------------------------------
        val_r28 = count_atwork_dates_from_raw(dfs["atwork_raw"], "Date Job Acquired", all_progs, fy_start, fy_end, event_based=False)
        cell_updates.append({"row": ROW_JOBS_ATTAINED, "col": col_idx, "val": val_r28})

        val_r29 = count_atwork_dates_from_raw(dfs["atwork_raw"], IC_DATE_COLS, all_progs, fy_start, fy_end, event_based=True)
        cell_updates.append({"row": ROW_CREDENTIALS_ATTAINED, "col": col_idx, "val": val_r29})

        cnt_bc = count_atwork_dates_from_raw(dfs["atwork_raw"], "Birth Certificate Obtained Date_6638", all_progs, fy_start, fy_end, event_based=False)
        cnt_id = count_atwork_dates_from_raw(dfs["atwork_raw"], "State ID Obtained Date_6637", all_progs, fy_start, fy_end, event_based=False)
        cnt_ss = count_atwork_dates_from_raw(dfs["atwork_raw"], "Social Security Card Obtained Date_6639", all_progs, fy_start, fy_end, event_based=False)
        cell_updates.append({"row": ROW_IDS_CARDS_DOCS, "col": col_idx, "val": (cnt_bc + cnt_id + cnt_ss)})

        val_r31 = count_atwork_dates_from_raw(dfs["atwork_raw"], "Driver's License Date_5071", all_progs, fy_start, fy_end, event_based=False)
        cell_updates.append({"row": ROW_DRIVERS_LICENSES, "col": col_idx, "val": val_r31})

        val_r32 = count_atwork_dates_from_raw(dfs["atwork_raw"], "High School Equivalency (HSE) Date_5110", all_progs, fy_start, fy_end, event_based=False)
        cell_updates.append({"row": ROW_HSE_ATTAINED, "col": col_idx, "val": val_r32})

        val_r33 = count_atwork_dates_from_raw(dfs["atwork_raw"], "Bank Account Obtained Date_6640", all_progs, fy_start, fy_end, event_based=False)
        cell_updates.append({"row": ROW_BANK_ACCOUNTS, "col": col_idx, "val": val_r33})

        # Rows 34 & 35: Housing Burden & Wage Increase (Evaluates Renew Program Graduates Only)
        cell_updates.append({"row": ROW_HOUSING_BURDEN_1YR, "col": col_idx, "val": alum_metrics_renew["housing"]})
        cell_updates.append({"row": ROW_WAGE_INCREASE, "col": col_idx, "val": alum_metrics_renew["wage"]})

    log(f"   Writing cells directly to Google Sheet tab '{sheet_name}'...")
    gspread_data = [{"range": gspread.utils.rowcol_to_a1(item["row"], item["col"]), "values": [[item["val"]]]} for item in cell_updates]
    ws.batch_update(gspread_data, value_input_option='USER_ENTERED')
    log(f"✅ Successfully written values directly to '{sheet_name}'!")

def run_marketing_processor():
    log("🚀 Starting Marketing Impact Data Processing across Shared Drives...")

    main_drive_service, gc, _ = get_services()

    # 1. Resolve Root Shared Folder 'CityTeam Impact Data'
    log("📂 Resolving root shared folder: 'CityTeam Impact Data'...")
    impact_folder_id = None
    try:
        impact_folder_id = resolve_folder_id(main_drive_service, "CityTeam Impact Data", "Input")
        log("   ✓ Located 'CityTeam Impact Data' folder!")
    except Exception as e:
        log(f"   ⚠️ Could not resolve root folder via strict name search: {e}. Falling back to subfolder resolution.")

    # 2. Resolve Incoming Raw Reports Folder: 'Apricot Report Incoming'
    log("📂 Resolving input folder: 'Apricot Report Incoming'...")
    incoming_res = find_file_id(main_drive_service, "Apricot Report Incoming", impact_folder_id, "application/vnd.google-apps.folder")
    if incoming_res:
        input_folder_id = incoming_res['id']
        log("   ✓ Located 'Apricot Report Incoming' folder ID!")
    else:
        log("   ⚠️ 'Apricot Report Incoming' subfolder not found under root. Attempting direct folder resolve...")
        input_folder_id = resolve_folder_id(main_drive_service, "Apricot Report Incoming")

    # 3. Resolve Processed Folder: 'Quarterly KPI Processed Data'
    log("📂 Resolving occupancy & processed folder: 'Quarterly KPI Processed Data'...")
    target_occ_folder_id = input_folder_id
    try:
        kpi_res = find_file_id(main_drive_service, "CityTeam KPIs", impact_folder_id, "application/vnd.google-apps.folder")
        if not kpi_res:
            kpi_res = find_file_id(main_drive_service, "CityTeam KPIs", None, "application/vnd.google-apps.folder")

        if kpi_res:
            proc_res = find_file_id(main_drive_service, "Quarterly KPI Processed Data", kpi_res['id'], "application/vnd.google-apps.folder")
            if proc_res:
                target_occ_folder_id = proc_res['id']
                log("   ✓ Located target folder: 'Quarterly KPI Processed Data'!")
            else:
                target_occ_folder_id = kpi_res['id']
                log("   ⚠️ 'Quarterly KPI Processed Data' not found under 'CityTeam KPIs'; using parent.")
    except Exception as e:
        log(f"   ⚠️ Exception resolving occupancy folder path: {e}")

    # 4. Locate and load 'Clients in All Programs Average' (Latest Occupancy & Latest Mentorship)
    log(f"📄 Locating occupancy dataset: '{FILE_OCCUPANCY_NAME}'...")
    df_occ = pd.DataFrame()
    df_ment_occ = pd.DataFrame()
    try:
        occ_file_res = find_file_id(main_drive_service, FILE_OCCUPANCY_NAME, target_occ_folder_id)
        if occ_file_res:
            mime = occ_file_res.get("mimeType", "")
            if mime == "application/vnd.google-apps.spreadsheet":
                occ_ss = gc.open_by_key(occ_file_res['id'])
                
                # Sheet 1: Latest Occupancy
                try:
                    ws_occ = occ_ss.worksheet("Latest Occupancy")
                    df_occ = clean_df_columns(pd.DataFrame(ws_occ.get_all_records()))
                    log("   ✓ Loaded 'Latest Occupancy' from native Google Sheet!")
                except Exception as e:
                    log(f"   ⚠️ Tab 'Latest Occupancy' missing: {e}")

                # Sheet 2: Latest Mentorship
                try:
                    ws_ment_occ = occ_ss.worksheet("Latest Mentorship")
                    df_ment_occ = clean_df_columns(pd.DataFrame(ws_ment_occ.get_all_records()))
                    log("   ✓ Loaded 'Latest Mentorship' from native Google Sheet!")
                except Exception as e:
                    log(f"   ⚠️ Tab 'Latest Mentorship' missing: {e}")
            else:
                fh_occ, _, _ = download_drive_file(main_drive_service, occ_file_res['name'], target_occ_folder_id)
                xls = pd.ExcelFile(fh_occ)
                if "Latest Occupancy" in xls.sheet_names:
                    df_occ = clean_df_columns(pd.read_excel(xls, sheet_name="Latest Occupancy"))
                if "Latest Mentorship" in xls.sheet_names:
                    df_ment_occ = clean_df_columns(pd.read_excel(xls, sheet_name="Latest Mentorship"))
                log("   ✓ Successfully loaded tabs from Excel occupancy workbook!")
        else:
            log(f"⚠️ Could not find '{FILE_OCCUPANCY_NAME}' in target folder.")
    except Exception as e:
        log(f"⚠️ Warning loading occupancy data: {e}")

    # 5. Parallel download for raw Excel source files from 'Apricot Report Incoming'
    log("⚡ Downloading raw Excel source files in parallel from 'Apricot Report Incoming'...")
    download_tasks = [
        ("renew", FILE_RENEW_NAME, input_folder_id),
        ("tp", FILE_TP_NAME, input_folder_id),
        ("bap", FILE_BAPTISMS_NAME, input_folder_id),
        ("ment", FILE_MENTOR_NAME, input_folder_id),
        ("reun", FILE_REUNIF_NAME, input_folder_id),
        ("alum", FILE_ALUMNI_NAME, input_folder_id),
        ("atwork", FILE_ATWORK_NAME, input_folder_id)
    ]

    downloads = {}
    with ThreadPoolExecutor(max_workers=len(download_tasks)) as executor:
        for key, fh, file_id in executor.map(thread_safe_download, download_tasks):
            downloads[key] = (fh, file_id)
            log(f"   ✓ Downloaded: {key}")

    fh_renew, _ = downloads["renew"]
    fh_tp, _ = downloads["tp"]
    fh_bap, _ = downloads["bap"]
    fh_ment, _ = downloads["ment"]
    fh_reun, _ = downloads["reun"]
    fh_alum, _ = downloads["alum"]
    fh_atwork, _ = downloads["atwork"]

    log("   Parsing raw Excel data into DataFrames...")
    df_renew_raw  = clean_df_columns(pd.read_excel(fh_renew, header=3).iloc[:, 1:])
    df_tp_raw     = clean_df_columns(pd.read_excel(fh_tp, header=3).iloc[:, 1:])
    df_bap_raw    = clean_df_columns(pd.read_excel(fh_bap, header=3).iloc[:, 1:])
    df_ment_raw   = clean_df_columns(pd.read_excel(fh_ment, header=3).iloc[:, 1:])
    df_reun_raw   = clean_df_columns(pd.read_excel(fh_reun, header=3).iloc[:, 1:])
    df_alumni_raw = clean_df_columns(pd.read_excel(fh_alum, header=3).iloc[:, 1:])
    df_atwork_raw = clean_df_columns(pd.read_excel(fh_atwork, header=3).iloc[:, 1:])

    dfs = {
        "occ": df_occ,
        "ment_occ": df_ment_occ,
        "renew_raw": df_renew_raw,
        "tp_raw": df_tp_raw,
        "bap_raw": df_bap_raw,
        "ment_raw": df_ment_raw,
        "reun_raw": df_reun_raw,
        "alumni_raw": df_alumni_raw,
        "atwork_raw": df_atwork_raw
    }

    # Execute KPI updates for FY26 and FY27
    for fy_key, fy_info in FY_CONFIGS.items():
        process_fiscal_year(fy_key, fy_info, gc, dfs)

    # Automatically generate and upload the Marketing Data Impact Spreadsheet
    generate_and_upload_audit_spreadsheet(main_drive_service, target_occ_folder_id)

    log("\n🎉 Marketing Processor & Audit Workbook Execution Complete!")

if __name__ == "__main__":
    run_marketing_processor()