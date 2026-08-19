# ============================================================
#  constants.py — Zero-Redundancy KPI Constants Module
#  Call sync_constants() once at startup to populate all values.
# ============================================================

import calendar
import pandas as pd
from googleapiclient.discovery import build
from auth_utils import get_services
from drive_utils import find_file_id, resolve_folder_id

# ============================================================
#  STATIC CONFIGURATION
# ============================================================

MONTHLY_HOURS = 173.33
GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]

# Core Program Sets (Single source of truth)
RENEW_PROGRAMS = {
    "Chester Renew", "GV Renew", "Oakland Renew", "Portland Renew", "San Jose Men Renew"
}
TP_PROGRAMS = [
    "Chester Men Turning Point", "Chester Women Turning Point",
    "Oakland Men Turning Point", "Oakland Women Turning Point",
    "Portland Community of Hope", "GV Turning Point",
    "Heritage Home", "San Jose Men Turning Point", "San Jose Youth Collective"
]

VALID_PROGRAMS_ALL = RENEW_PROGRAMS | set(TP_PROGRAMS) | {"House of Grace", "Program Graduate Intern"}
VALID_PROGRAMS_NUM_DEN = {"San Jose Men Renew", "GV Renew"}
VALID_PROGRAMS_IMPACT = VALID_PROGRAMS_ALL
VALID_INTERN_PROGRAMS_IMPACT = VALID_PROGRAMS_ALL - {"Program Graduate Intern"}

BARRIER_REMAP = {
    "Criminal History or CASU": "Justice System Involved",
    "Suspended driver's licence": "Suspended or No Driver's License",
    "Suspended driver's license": "Suspended or No Driver's License",
    "No driver's license": "Suspended or No Driver's License",
    "No DL": "Suspended or No Driver's License",
}

IC_DATE_COLS = [
    "Electrical Trainee Card Date_5727", "NCCER Date_5728", "Blueprint Reading Date_6785",
    "Forklift Operator Date_5106", "OSHA10 Date_5118", "MA Date_5115",
    "Google IT Support Certificate Date_5107", "HAZWOPER Date_5109", "HAZMAT Date_5108",
    "CPR/FA/AED Date_5105", "CADC Date_5789", "ServSafe Food Manager Date_5126",
    "ServSafe Food Handler Date_5125", "Material Handling Date_5116",
    "Multi-Core Craft Curriculum Completion Date_5117", "RADT Date_5121", "HVAC Date_5113",
    "CDL Date_6104", "CompTIA and/or A+ Date_5104", "OTHER Date_5119", "Guard Card Date_6561",
    "TWIC Date_6631", "Google Automate Cybersecurity w/Python Date_6273", "CNC Machining Date_5103",
    "Confined Space Date_6788", "DVAT Date_6791", "Fundamentals of Construction Date_6790",
    "HR Associate Date_6789", "Intuit Bookkeeping Date_6105", "Massage Therapy Cert Date_7032",
    "Peer Support Date_7031", "Record Expungement Date_6787",
    "Record Expungement Attorney Consultation Date_6879", "Traffic Control Date_6786"
]

EXTRA_WINDOW_COLS = [
    "High School Equivalency (HSE) Date_5110", "Bank Account Obtained Date_6640",
    "Driver's License Date_5071", "Birth Certificate Obtained Date_6638",
    "State ID Obtained Date_6637", "Social Security Card Obtained Date_6639",
    "Date Job Acquired_5590"
]

# Generate Forward Map variations programmatically to avoid key repetition
_FORWARD_BASE = {
    "GV Men's Forward": "San Jose Men Renew",
    "GV Women's Forward": "GV Renew",
    "House of Grace": "GV Renew",
    "San Jose Women's House of Light": "GV Renew",
    "San Jose women's House of light": "GV Renew",
    "Chester Forward": "Chester Renew",
    "Oakland Forward": "Oakland Renew",
    "Oakland Forward Men": "Oakland Renew",
    "Oakland Men's Forward": "Oakland Renew",
    "Portland Forward": "Portland Renew"
}
FORWARD_MAP = {**_FORWARD_BASE, **{k.lower(): v for k, v in _FORWARD_BASE.items()}}


# ============================================================
#  DYNAMIC GLOBALS INITIALIZATION
# ============================================================

fy_num = 2026
FY_LABEL, Q_LABEL = "FY26", "Q1"
PRIOR_FY_LABEL = PP_FY_LABEL = PPP_FY_LABEL = ""

CURRENT_Q_START = CURRENT_Q_END = CURRENT_FY_QX_START = CURRENT_FY_QX_END = None
PRIOR_FY_QX_START = PRIOR_FY_QX_END = None
PP_FY_START = PP_FY_END = PP_FY_QX_END = None
PPP_FY_START = PPP_FY_END = PPP_FY_QX_END = None
FY_START = FY_END = RETENTION_START_CUTOFF = None
OCC_CURRENT_DATE = OCC_PRIOR_DATE = None
SOBRIETY_START = SOBRIETY_END = LW_HOUSING_START = LW_HOUSING_END = None
FY_QUARTERS = []

OCC_CURRENT_LABEL, OCC_PRIOR_LABEL = "Current Period Actuals", "Prior Period Actuals"

ACTUALS_WINDOWS, PROGRAM_GOALS, PROGRAM_PROJECTIONS = {}, {}, {}
PROGRAM_CAPACITY, PROGRAM_THEORETICAL_MAX = {}, {}
OCCUPANCY_CAPACITY, OCCUPANCY_GOAL, OCCUPANCY_PRIOR_FY = {}, {}, {}
LW_CRITERIA, TP_CAPACITY, TP_GOAL, TP_PRIOR_FY = {}, {}, {}, {}


# ============================================================
#  LOAD CONSTANTS FROM GOOGLE SHEET
# ============================================================

def sync_constants():
    drive_service, _, creds = get_services()
    folder_path = "CityTeam Impact Data/CityTeam KPIs/FY26/FY26 Q3"
    sheet_name = "Constants Update"

    folder_id = resolve_folder_id(drive_service, folder_path.split("/")[-1], "Constants")
    target_file = find_file_id(drive_service, sheet_name, folder_id)
    if not target_file:
        raise FileNotFoundError(f"❌ '{sheet_name}' not found in {folder_path}")

    sheets_service = build("sheets", "v4", credentials=creds)
    sheet_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=target_file['id'], range="A1:O42"
    ).execute().get("values", [])

    def get_val(r_idx, c_idx):
        try:
            val = sheet_data[r_idx][c_idx]
            return float(str(val).replace("$", "").replace(",", "").strip())
        except (IndexError, ValueError, TypeError):
            return None

    # Report Period Setup
    row3 = sheet_data[2] if len(sheet_data) > 2 else []
    raw_fy, raw_q = (row3[1] if len(row3) > 1 else None), (row3[2] if len(row3) > 2 else None)
    if raw_fy is None:
        raise ValueError("❌ Fiscal Year missing at B3.")

    fy = int(str(raw_fy).strip())
    fy = fy + 2000 if fy < 100 else fy
    fy_lbl = f"FY{str(fy)[-2:]}"
    q_lbl = f"Q{str(raw_q).strip()}" if str(raw_q).strip().isdigit() else str(raw_q).strip()

    # Dates Calculation
    q_months = {"Q1": (9, 11), "Q2": (12, 2), "Q3": (3, 5), "Q4": (6, 8)}
    q_start_m, q_end_m = q_months[q_lbl]
    fy_start_yr = fy - 1
    q_start_yr = fy_start_yr if q_lbl in ["Q1", "Q2"] else fy
    q_end_yr = fy_start_yr if q_lbl == "Q1" else fy
    q_end_day = calendar.monthrange(q_end_yr, q_end_m)[1]

    cur_q_start = pd.Timestamp(f"{q_start_yr}-{q_start_m:02d}-01")
    cur_q_end   = pd.Timestamp(f"{q_end_yr}-{q_end_m:02d}-{q_end_day:02d}")
    fy_start    = pd.Timestamp(f"{fy_start_yr}-09-01")

    prior_fy_num = fy - 1
    prior_fy_qx_end = cur_q_end - pd.DateOffset(years=1)

    # Matrix Loading for Renew Programs (Cols B, C, D, E, G, H, I)
    sorted_renew = sorted(RENEW_PROGRAMS)
    renew_col_map = [
        (PROGRAM_GOALS, 1), (PROGRAM_PROJECTIONS, 2), (PROGRAM_CAPACITY, 3),
        (PROGRAM_THEORETICAL_MAX, 4), (OCCUPANCY_CAPACITY, 6),
        (OCCUPANCY_GOAL, 7), (OCCUPANCY_PRIOR_FY, 8)
    ]
    for idx, prog in enumerate(sorted_renew):
        row = 9 + (idx * 4)
        for target_dict, col in renew_col_map:
            target_dict[prog] = get_val(row, col)
        LW_CRITERIA[prog] = get_val(8 + (idx * 2), 10)  # Col K

    # Matrix Loading for Turning Point Programs (Cols M, N, O)
    tp_col_map = [(TP_CAPACITY, 12), (TP_GOAL, 13), (TP_PRIOR_FY, 14)]
    for idx, prog in enumerate(TP_PROGRAMS):
        row = 9 + (idx * 4)
        for target_dict, col in tp_col_map:
            target_dict[prog] = get_val(row, col)

    # Shift Back Helper for Alumni Windows
    def _shift_back(ts, months):
        m = ts.month - months
        y = ts.year + (m - 1) // 12
        m = ((m - 1) % 12) + 1
        return pd.Timestamp(f"{y}-{m:02d}-{min(ts.day, calendar.monthrange(y, m)[1]):02d}")

    sobriety_start = _shift_back(cur_q_start, 21)
    if get_fiscal_quarter(sobriety_start) == "Q4":
        sobriety_start = _shift_back(sobriety_start, 3)

    # Global State Update
    globals().update({
        "fy_num": fy, "FY_LABEL": fy_lbl, "Q_LABEL": q_lbl,
        "PRIOR_FY_LABEL": f"FY{str(prior_fy_num)[-2:]}",
        "PP_FY_LABEL": f"FY{str(fy - 2)[-2:]}",
        "PPP_FY_LABEL": f"FY{str(fy - 3)[-2:]}",
        "CURRENT_Q_START": cur_q_start, "CURRENT_Q_END": cur_q_end,
        "CURRENT_FY_QX_START": fy_start, "CURRENT_FY_QX_END": cur_q_end,
        "PRIOR_FY_QX_START": pd.Timestamp(f"{prior_fy_num - 1}-09-01"),
        "PRIOR_FY_QX_END": prior_fy_qx_end,
        "PP_FY_START": pd.Timestamp(f"{fy - 3}-09-01"),
        "PP_FY_END": pd.Timestamp(f"{fy - 2}-08-31"),
        "PP_FY_QX_END": prior_fy_qx_end - pd.DateOffset(years=1),
        "PPP_FY_START": pd.Timestamp(f"{fy - 4}-09-01"),
        "PPP_FY_END": pd.Timestamp(f"{fy - 3}-08-31"),
        "PPP_FY_QX_END": prior_fy_qx_end - pd.DateOffset(years=2),
        "FY_START": fy_start, "FY_END": pd.Timestamp(f"{fy}-08-31"),
        "RETENTION_START_CUTOFF": pd.Timestamp(f"{fy - 2}-09-01"),
        "OCC_CURRENT_DATE": cur_q_end, "OCC_PRIOR_DATE": prior_fy_qx_end,
        "SOBRIETY_START": sobriety_start,
        "SOBRIETY_END": _shift_back(cur_q_end, 21),
        "LW_HOUSING_START": pd.Timestamp(f"{fy - 2}-09-01") if q_lbl == "Q4" else cur_q_start - pd.DateOffset(years=1),
        "LW_HOUSING_END": pd.Timestamp(f"{fy - 1}-08-31") if q_lbl == "Q4" else cur_q_end - pd.DateOffset(years=1),
        "FY_QUARTERS": generate_fy_quarters(fy)
    })

    ACTUALS_WINDOWS.update({
        "Current Period Actuals": (FY_START, CURRENT_Q_END),
        "Prior Period Actuals": (PRIOR_FY_QX_START, PRIOR_FY_QX_END),
        "2 Years Ago Period Actuals": (PP_FY_START, PP_FY_QX_END),
        "3 Years Ago Period Actuals": (PPP_FY_START, PPP_FY_QX_END),
        "Prior Year Actuals": (pd.Timestamp(f"{fy - 2}-09-01"), pd.Timestamp(f"{fy - 1}-08-31")),
        "2 Years Ago Actuals": (PP_FY_START, PP_FY_END),
        "3 Years Ago Actuals": (PPP_FY_START, PPP_FY_END),
    })

    print(f"✅ All dynamic constants loaded for {FY_LABEL} {Q_LABEL}.")


# ============================================================
#  HELPER FUNCTIONS
# ============================================================

def get_fiscal_year(date):
    if pd.isnull(date): return None
    ts = pd.to_datetime(date)
    for offset in range(-5, 6):
        fy = fy_num + offset
        if pd.Timestamp(f"{fy - 1}-09-01") <= ts <= pd.Timestamp(f"{fy}-08-31"):
            return f"FY{str(fy)[-2:]}"
    return None

def get_fiscal_quarter(date):
    if pd.isnull(date): return None
    return f"Q{(pd.to_datetime(date).month - 9) % 12 // 3 + 1}"

def generate_fy_quarters(center_fy, spread=3):
    quarters = []
    for offset in range(-spread, spread + 1):
        fy = center_fy + offset
        fl, sy = f"FY{str(fy)[-2:]}", fy - 1
        feb = 29 if calendar.isleap(fy) else 28
        quarters.extend([
            (f"{fl} Q1", f"{sy}-09-01", f"{sy}-11-30"),
            (f"{fl} Q2", f"{sy}-12-01", f"{fy}-02-{feb:02d}"),
            (f"{fl} Q3", f"{fy}-03-01", f"{fy}-05-31"),
            (f"{fl} Q4", f"{fy}-06-01", f"{fy}-08-31"),
        ])
    return quarters

def assign_city(program):
    if pd.isna(program): return "Unknown"
    p = str(program).strip()
    for city in ["Chester", "Portland", "Oakland"]:
        if city in p: return city
    if "GV Men" in p or "San Jose Men" in p: return "San Jose Men"
    if any(k in p for k in ["San Jose", "GV", "Heritage", "House of Grace"]):
        return "San Jose Women"
    return "Unknown"

print("✅ Static configuration loaded.")