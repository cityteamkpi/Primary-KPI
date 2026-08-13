# ============================================================
#  This script parses a Constants file
# ============================================================

from googleapiclient.discovery import build
import re
import pandas as pd
import calendar
from auth_utils import get_services
from drive_utils import find_file_id, resolve_folder_id


# --- FY Quarter ranges for summary tabs ---
### UNUSED ###
# Edit these lists to change which quarters are compared
#FY25_QUARTERS = ["FY25 Q1", "FY25 Q2", "FY25 Q3"]   # ← edit here
#FY26_QUARTERS = ["FY26 Q1", "FY26 Q2", "FY26 Q3"]   # ← edit here

# --- FY range to include (covers FY23–FY26) ---

# --- Forward program → Renew program mapping ---
FORWARD_MAP = {
    "GV Men's Forward"               : "San Jose Men Renew",
    "GV Men's forward"               : "San Jose Men Renew",
    "GV Women's Forward"             : "GV Renew",
    "GV Women's forward"             : "GV Renew",
    "House of Grace"                 : "GV Renew",
    "San Jose women's House of light": "GV Renew",
    "San Jose Women's House of Light": "GV Renew",
    "Chester Forward"                : "Chester Renew",
    "Chester forward"                : "Chester Renew",
    "Oakland Forward"                : "Oakland Renew",
    "Oakland forward men"            : "Oakland Renew",
    "Oakland Forward Men"            : "Oakland Renew",
    "Oakland Men's Forward"          : "Oakland Renew",
    "Portland Forward"               : "Portland Renew",
    "Portland forward"               : "Portland Renew",
}

# --- Dynamic Constants (Populated by sync_constants) ---
fy_num = 2026 # Default fallback
FY_LABEL = "FY26"
Q_LABEL = "Q1"
CURRENT_Q_START = None
CURRENT_Q_END = None
ACTUALS_WINDOWS = {}
PROGRAM_GOALS = {}
PROGRAM_PROJECTIONS = {}
PROGRAM_CAPACITY = {}
PROGRAM_THEORETICAL_MAX = {}
OCCUPANCY_CAPACITY = {}
OCCUPANCY_GOAL = {}
OCCUPANCY_PRIOR_FY = {}
LW_CRITERIA = {}
TP_CAPACITY = {}
TP_GOAL = {}
TP_PRIOR_FY = {}
FY_START = None
FY_END = None
RETENTION_START_CUTOFF = None
OCC_CURRENT_DATE = None
OCC_PRIOR_DATE = None
OCC_CURRENT_LABEL = "Current Period Actuals"
OCC_PRIOR_LABEL = "Prior Period Actuals"
PRIOR_FY_LABEL = ""

# ============================================================
#  LOAD CONSTANTS FROM GOOGLE SHEET
# ============================================================

def sync_constants():
    global fy_num, FY_LABEL, Q_LABEL, CURRENT_Q_START, CURRENT_Q_END
    global ACTUALS_WINDOWS, PROGRAM_GOALS, PROGRAM_PROJECTIONS, PROGRAM_CAPACITY
    global PROGRAM_THEORETICAL_MAX, OCCUPANCY_CAPACITY, OCCUPANCY_GOAL
    global OCCUPANCY_PRIOR_FY, LW_CRITERIA, TP_CAPACITY, TP_GOAL, TP_PRIOR_FY
    global FY_START, FY_END, RETENTION_START_CUTOFF, OCC_CURRENT_DATE, OCC_PRIOR_DATE, PRIOR_FY_LABEL

    drive_service, _, creds = get_services()

    # --- Locate Constants Update spreadsheet ---
    CONSTANTS_FOLDER_PATH = "CityTeam Impact Data/CityTeam KPIs/FY26/FY26 Q3"
    CONSTANTS_SHEET_NAME  = "Constants Update"

    # Using the leaf folder name to resolve the ID via drive_utils
    leaf_folder_name = CONSTANTS_FOLDER_PATH.split("/")[-1]
    constants_folder_id = resolve_folder_id(drive_service, leaf_folder_name, "Constants")

    target_file = find_file_id(drive_service, CONSTANTS_SHEET_NAME, constants_folder_id)
    if not target_file:
        raise FileNotFoundError(f"❌ '{CONSTANTS_SHEET_NAME}' not found in {CONSTANTS_FOLDER_PATH}")

    constants_sheet_id = target_file['id']
    print(f"✅ Constants sheet found: {CONSTANTS_SHEET_NAME}")

    # --- Read values from sheet ---
    sheets_service = build("sheets", "v4", credentials=creds)

    # Batch fetch relevant range
    sheet_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=constants_sheet_id,
        range="A1:O42"
    ).execute().get("values", [])

    def get_val_from_batch(cell_ref):
        match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
        col_str, row_str = match.groups()
        col_idx = ord(col_str) - ord('A')
        row_idx = int(row_str) - 1
        try:
            val = sheet_data[row_idx][col_idx]
            return float(str(val).replace("$","").replace(",","").strip())
        except: return None

    # ── Fiscal Year & Quarter ─────────────────────────────────────
    raw_fy = sheet_data[2][1] if len(sheet_data) > 2 else None # B3
    raw_q  = sheet_data[2][2] if len(sheet_data) > 2 else None # C3

    fy_num = int(str(raw_fy).strip())
    if fy_num < 100: fy_num += 2000
    FY_LABEL = f"FY{str(fy_num)[-2:]}"
    _raw_q = str(raw_q).strip()
    Q_LABEL = f"Q{_raw_q}" if _raw_q.isdigit() else _raw_q
    print(f"✅ Report period: {FY_LABEL} {Q_LABEL}")

    # ── Date Calculations ─────────────────────────────────────────
    Q_END_MONTH = {"Q1": 11, "Q2": 2, "Q3": 5, "Q4": 8}
    Q_END_DAY   = {"Q1": 30, "Q2": 28, "Q3": 31, "Q4": 31}
    q_num = int(Q_LABEL[1])

    fy_start_year = fy_num - 1
    fy_start = pd.Timestamp(f"{fy_start_year}-09-01")
    FY_START = fy_start
    FY_END = pd.Timestamp(f"{fy_num}-08-31")

    if Q_LABEL == "Q1": q_end_year = fy_start_year
    else: q_end_year = fy_num

    q_end_month = Q_END_MONTH[Q_LABEL]
    q_end_day   = Q_END_DAY[Q_LABEL]
    if q_end_month == 2: q_end_day = 29 if calendar.isleap(q_end_year) else 28

    Q_START_MONTH = {"Q1": 9, "Q2": 12, "Q3": 3, "Q4": 6}
    q_start_month = Q_START_MONTH[Q_LABEL]
    q_start_year  = fy_start_year if Q_LABEL in ["Q1", "Q2"] else fy_num
    
    CURRENT_Q_START = pd.Timestamp(f"{q_start_year}-{q_start_month:02d}-01")
    CURRENT_Q_END   = pd.Timestamp(f"{q_end_year}-{q_end_month:02d}-{q_end_day:02d}")

    # Prior FY
    prior_fy_num        = fy_num - 1
    PRIOR_FY_LABEL      = f"FY{str(prior_fy_num)[-2:]}"
    PRIOR_FY_QX_START   = pd.Timestamp(f"{prior_fy_num - 1}-09-01")
    PRIOR_FY_QX_END     = CURRENT_Q_END.replace(year=CURRENT_Q_END.year - 1)

    # Actuals windows
    pp_fy_num = fy_num - 2
    PP_FY_START = pd.Timestamp(f"{pp_fy_num - 1}-09-01")
    PP_FY_END = pd.Timestamp(f"{pp_fy_num}-08-31")
    
    ppp_fy_num = fy_num - 3
    PPP_FY_START = pd.Timestamp(f"{ppp_fy_num - 1}-09-01")
    PPP_FY_END = pd.Timestamp(f"{ppp_fy_num}-08-31")

    ACTUALS_WINDOWS = {
        "Current Period Actuals"  : (fy_start, CURRENT_Q_END),
        "Prior Period Actuals"    : (PRIOR_FY_QX_START, PRIOR_FY_QX_END),
        "2 Years Ago Actuals"     : (PP_FY_START, PP_FY_END),
        "3 Years Ago Actuals"     : (PPP_FY_START, PPP_FY_END),
    }

    RETENTION_START_CUTOFF = pd.Timestamp(f"{fy_num - 2}-09-01")
    OCC_CURRENT_DATE = CURRENT_Q_END
    OCC_PRIOR_DATE   = PRIOR_FY_QX_END

    # ── Mappings ──────────────────────────────────────────────────
    PROGRAM_GOALS.update({"Chester Renew": get_val_from_batch("B10"), "GV Renew": get_val_from_batch("B14"), "Oakland Renew": get_val_from_batch("B18"), "Portland Renew": get_val_from_batch("B22"), "San Jose Men Renew": get_val_from_batch("B26")})
    PROGRAM_PROJECTIONS.update({"Chester Renew": get_val_from_batch("C10"), "GV Renew": get_val_from_batch("C14"), "Oakland Renew": get_val_from_batch("C18"), "Portland Renew": get_val_from_batch("C22"), "San Jose Men Renew": get_val_from_batch("C26")})
    PROGRAM_CAPACITY.update({"Chester Renew": get_val_from_batch("D10"), "GV Renew": get_val_from_batch("D14"), "Oakland Renew": get_val_from_batch("D18"), "Portland Renew": get_val_from_batch("D22"), "San Jose Men Renew": get_val_from_batch("D26")})
    PROGRAM_THEORETICAL_MAX.update({"Chester Renew": get_val_from_batch("E10"), "GV Renew": get_val_from_batch("E14"), "Oakland Renew": get_val_from_batch("E18"), "Portland Renew": get_val_from_batch("E22"), "San Jose Men Renew": get_val_from_batch("E26")})
    OCCUPANCY_CAPACITY.update({"Chester Renew": get_val_from_batch("G10"), "GV Renew": get_val_from_batch("G14"), "Oakland Renew": get_val_from_batch("G18"), "Portland Renew": get_val_from_batch("G22"), "San Jose Men Renew": get_val_from_batch("G26")})
    OCCUPANCY_GOAL.update({"Chester Renew": get_val_from_batch("H10"), "GV Renew": get_val_from_batch("H14"), "Oakland Renew": get_val_from_batch("H18"), "Portland Renew": get_val_from_batch("H22"), "San Jose Men Renew": get_val_from_batch("H26")})
    OCCUPANCY_PRIOR_FY.update({"Chester Renew": get_val_from_batch("I10"), "GV Renew": get_val_from_batch("I14"), "Oakland Renew": get_val_from_batch("I18"), "Portland Renew": get_val_from_batch("I22"), "San Jose Men Renew": get_val_from_batch("I26")})
    LW_CRITERIA.update({"Chester Renew": get_val_from_batch("K9"), "Portland Renew": get_val_from_batch("K11"), "Oakland Renew": get_val_from_batch("K13"), "San Jose Men Renew": get_val_from_batch("K15"), "GV Renew": get_val_from_batch("K17")})
    TP_CAPACITY.update({"Chester Men Turning Point": get_val_from_batch("M10"), "Chester Women Turning Point": get_val_from_batch("M14"), "Oakland Men Turning Point": get_val_from_batch("M18"), "Oakland Women Turning Point": get_val_from_batch("M22"), "Portland Community of Hope": get_val_from_batch("M26"), "GV Turning Point": get_val_from_batch("M30"), "Heritage Home": get_val_from_batch("M34"), "San Jose Men Turning Point": get_val_from_batch("M38"), "San Jose Youth Collective": get_val_from_batch("M42")})
    TP_GOAL.update({"Chester Men Turning Point": get_val_from_batch("N10"), "Chester Women Turning Point": get_val_from_batch("N14"), "Oakland Men Turning Point": get_val_from_batch("N18"), "Oakland Women Turning Point": get_val_from_batch("N22"), "Portland Community of Hope": get_val_from_batch("N26"), "GV Turning Point": get_val_from_batch("N30"), "Heritage Home": get_val_from_batch("N34"), "San Jose Men Turning Point": get_val_from_batch("N38"), "San Jose Youth Collective": get_val_from_batch("N42")})
    TP_PRIOR_FY.update({"Chester Men Turning Point": get_val_from_batch("O10"), "Chester Women Turning Point": get_val_from_batch("O14"), "Oakland Men Turning Point": get_val_from_batch("O18"), "Oakland Women Turning Point": get_val_from_batch("O22"), "Portland Community of Hope": get_val_from_batch("O26"), "GV Turning Point": get_val_from_batch("O30"), "Heritage Home": get_val_from_batch("O34"), "San Jose Men Turning Point": get_val_from_batch("O38"), "San Jose Youth Collective": get_val_from_batch("O42")})

    print(f"✅ All dynamic constants loaded for {FY_LABEL} {Q_LABEL}.")

print("✅ Static configuration loaded.")



# 4. Helper Functions

# ── Dynamic fiscal year/quarter helpers (based on constants sheet) ────────────
def get_fiscal_year(date):
    if pd.isnull(date): return None
    for offset in range(-5, 6):
        fy = fy_num + offset
        start = pd.Timestamp(f"{fy - 1}-09-01")
        end   = pd.Timestamp(f"{fy}-08-31")
        if start <= date <= end:
            return f"FY{str(fy)[-2:]}"
    return None

def get_fiscal_quarter(date):
    if pd.isnull(date): return None
    m = date.month
    if m in [9, 10, 11]:  return "Q1"
    if m in [12, 1, 2]:   return "Q2"
    if m in [3, 4, 5]:    return "Q3"
    if m in [6, 7, 8]:    return "Q4"
    return None

def generate_fy_quarters(center_fy, spread=3):
    quarters = []
    for offset in range(-spread, spread + 1):
        fy = center_fy + offset
        fl = f"FY{str(fy)[-2:]}"
        sy = fy - 1
        feb = 29 if calendar.isleap(fy) else 28
        quarters += [
            (f"{fl} Q1", f"{sy}-09-01", f"{sy}-11-30"),
            (f"{fl} Q2", f"{sy}-12-01", f"{fy}-02-{feb:02d}"),
            (f"{fl} Q3", f"{fy}-03-01", f"{fy}-05-31"),
            (f"{fl} Q4", f"{fy}-06-01", f"{fy}-08-31"),
        ]
    return quarters

FY_QUARTERS = generate_fy_quarters(fy_num)
FY_START    = pd.Timestamp(f"{fy_num - 4}-09-01")
FY_END      = pd.Timestamp(f"{fy_num + 2}-08-31")

def assign_city(program):
    if pd.isna(program): return "Unknown"
    p = str(program).strip()
    if "Chester" in p:   return "Chester"
    if "Portland" in p:  return "Portland"
    if "Oakland" in p:   return "Oakland"
    if "GV Men" in p or "San Jose Men" in p: return "San Jose Men"
    if p in ["San Jose Women's House of Light", "San Jose Youth Collective",
             "SJ WP Transition Phase", "GV Renew", "GV Turning Point",
             "Heritage Home", "House of Grace", "GV APH",
             "GV Transition Phase", "GV Women's Forward"] or "GV" in p or "Heritage" in p:
        return "San Jose Women"
    if "San Jose" in p:  return "San Jose Women"
    return "Unknown"

print(f"✅ Dynamic fiscal helpers defined (centered on {FY_LABEL})")
