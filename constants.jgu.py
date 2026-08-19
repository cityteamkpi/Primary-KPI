# ============================================================
#  constants.py — Shared constants for all 5 KPI scripts
#  Call sync_constants() once at startup to populate all values.
# ============================================================

from googleapiclient.discovery import build
import re
import pandas as pd
import calendar
from auth_utils import get_services
from drive_utils import find_file_id, resolve_folder_id


# ============================================================
#  STATIC CONFIGURATION (script-independent)
# ============================================================

MONTHLY_HOURS = 173.33

GRAD_REASONS = ["Graduation", "Completer", "Completion of Program"]

RENEW_PROGRAMS = {
    "Chester Renew", "GV Renew", "Oakland Renew",
    "Portland Renew", "San Jose Men Renew"
}

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


# ============================================================
#  DYNAMIC CONSTANTS (populated by sync_constants())
# ============================================================

fy_num                  = 2026
FY_LABEL                = "FY26"
Q_LABEL                 = "Q1"
PRIOR_FY_LABEL          = ""
PP_FY_LABEL             = ""
PPP_FY_LABEL            = ""

CURRENT_Q_START         = None
CURRENT_Q_END           = None
CURRENT_FY_QX_START     = None
CURRENT_FY_QX_END       = None

PRIOR_FY_QX_START       = None
PRIOR_FY_QX_END         = None

PP_FY_START             = None
PP_FY_END               = None
PP_FY_QX_END            = None

PPP_FY_START            = None
PPP_FY_END              = None
PPP_FY_QX_END           = None

FY_START                = None
FY_END                  = None
FY_QUARTERS             = []

RETENTION_START_CUTOFF  = None
OCC_CURRENT_DATE        = None
OCC_PRIOR_DATE          = None
OCC_CURRENT_LABEL       = "Current Period Actuals"
OCC_PRIOR_LABEL         = "Prior Period Actuals"

# Alumni-specific windows (Script 5)
SOBRIETY_START          = None
SOBRIETY_END            = None
LW_HOUSING_START        = None
LW_HOUSING_END          = None

ACTUALS_WINDOWS         = {}
PROGRAM_GOALS           = {}
PROGRAM_PROJECTIONS     = {}
PROGRAM_CAPACITY        = {}
PROGRAM_THEORETICAL_MAX = {}
OCCUPANCY_CAPACITY      = {}
OCCUPANCY_GOAL          = {}
OCCUPANCY_PRIOR_FY      = {}
LW_CRITERIA             = {}
TP_CAPACITY             = {}
TP_GOAL                 = {}
TP_PRIOR_FY             = {}


# ============================================================
#  LOAD CONSTANTS FROM GOOGLE SHEET
# ============================================================

def sync_constants():
    global fy_num, FY_LABEL, Q_LABEL, PRIOR_FY_LABEL, PP_FY_LABEL, PPP_FY_LABEL
    global CURRENT_Q_START, CURRENT_Q_END, CURRENT_FY_QX_START, CURRENT_FY_QX_END
    global PRIOR_FY_QX_START, PRIOR_FY_QX_END
    global PP_FY_START, PP_FY_END, PP_FY_QX_END
    global PPP_FY_START, PPP_FY_END, PPP_FY_QX_END
    global FY_START, FY_END, FY_QUARTERS
    global RETENTION_START_CUTOFF, OCC_CURRENT_DATE, OCC_PRIOR_DATE
    global SOBRIETY_START, SOBRIETY_END, LW_HOUSING_START, LW_HOUSING_END
    global ACTUALS_WINDOWS, PROGRAM_GOALS, PROGRAM_PROJECTIONS, PROGRAM_CAPACITY
    global PROGRAM_THEORETICAL_MAX, OCCUPANCY_CAPACITY, OCCUPANCY_GOAL
    global OCCUPANCY_PRIOR_FY, LW_CRITERIA, TP_CAPACITY, TP_GOAL, TP_PRIOR_FY

    drive_service, _, creds = get_services()

    # --- Locate Constants Update spreadsheet ---
    CONSTANTS_FOLDER_PATH = "CityTeam Impact Data/CityTeam KPIs/FY26/FY26 Q3"
    CONSTANTS_SHEET_NAME  = "Constants Update"

    leaf_folder_name = CONSTANTS_FOLDER_PATH.split("/")[-1]
    constants_folder_id = resolve_folder_id(drive_service, leaf_folder_name, "Constants")

    target_file = find_file_id(drive_service, CONSTANTS_SHEET_NAME, constants_folder_id)
    if not target_file:
        raise FileNotFoundError(f"❌ '{CONSTANTS_SHEET_NAME}' not found in {CONSTANTS_FOLDER_PATH}")

    constants_sheet_id = target_file['id']
    print(f"✅ Constants sheet found: {CONSTANTS_SHEET_NAME}")

    # --- Batch fetch sheet data ---
    sheets_service = build("sheets", "v4", credentials=creds)
    sheet_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=constants_sheet_id,
        range="A1:O42"
    ).execute().get("values", [])

    def get_num(cell_ref):
        match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
        col_str, row_str = match.groups()
        col_idx = ord(col_str) - ord('A')
        row_idx = int(row_str) - 1
        try:
            val = sheet_data[row_idx][col_idx]
            return float(str(val).replace("$", "").replace(",", "").strip())
        except:
            return None

    # ── Fiscal Year & Quarter ─────────────────────────────────────
    raw_fy = sheet_data[2][1] if len(sheet_data) > 2 else None  # B3
    raw_q  = sheet_data[2][2] if len(sheet_data) > 2 else None  # C3

    fy_num = int(str(raw_fy).strip())
    if fy_num < 100: fy_num += 2000
    FY_LABEL = f"FY{str(fy_num)[-2:]}"
    _raw_q = str(raw_q).strip()
    Q_LABEL = f"Q{_raw_q}" if _raw_q.isdigit() else _raw_q
    print(f"✅ Report period: {FY_LABEL} {Q_LABEL}")

    # ── Date calculations ─────────────────────────────────────────
    Q_END_MONTH   = {"Q1": 11, "Q2": 2,  "Q3": 5,  "Q4": 8}
    Q_END_DAY     = {"Q1": 30, "Q2": 28, "Q3": 31, "Q4": 31}
    Q_START_MONTH = {"Q1": 9,  "Q2": 12, "Q3": 3,  "Q4": 6}

    fy_start_year = fy_num - 1
    fy_start      = pd.Timestamp(f"{fy_start_year}-09-01")
    FY_START      = fy_start
    FY_END        = pd.Timestamp(f"{fy_num}-08-31")

    q_end_year    = fy_start_year if Q_LABEL == "Q1" else fy_num
    q_end_month   = Q_END_MONTH[Q_LABEL]
    q_end_day     = Q_END_DAY[Q_LABEL]
    if q_end_month == 2:
        q_end_day = 29 if calendar.isleap(q_end_year) else 28

    q_start_month = Q_START_MONTH[Q_LABEL]
    q_start_year  = fy_start_year if Q_LABEL in ["Q1", "Q2"] else fy_num

    CURRENT_Q_START     = pd.Timestamp(f"{q_start_year}-{q_start_month:02d}-01")
    CURRENT_Q_END       = pd.Timestamp(f"{q_end_year}-{q_end_month:02d}-{q_end_day:02d}")
    CURRENT_FY_QX_START = fy_start
    CURRENT_FY_QX_END   = CURRENT_Q_END

    # Prior FY
    prior_fy_num        = fy_num - 1
    PRIOR_FY_LABEL      = f"FY{str(prior_fy_num)[-2:]}"
    prior_fy_start_year = prior_fy_num - 1
    PRIOR_FY_QX_START   = pd.Timestamp(f"{prior_fy_start_year}-09-01")
    PRIOR_FY_QX_END     = CURRENT_Q_END.replace(year=CURRENT_Q_END.year - 1)

    # 2 years ago
    pp_fy_num    = fy_num - 2
    PP_FY_LABEL  = f"FY{str(pp_fy_num)[-2:]}"
    PP_FY_START  = pd.Timestamp(f"{pp_fy_num - 1}-09-01")
    PP_FY_END    = pd.Timestamp(f"{pp_fy_num}-08-31")
    PP_FY_QX_END = PRIOR_FY_QX_END.replace(year=PRIOR_FY_QX_END.year - 1)

    # 3 years ago
    ppp_fy_num    = fy_num - 3
    PPP_FY_LABEL  = f"FY{str(ppp_fy_num)[-2:]}"
    PPP_FY_START  = pd.Timestamp(f"{ppp_fy_num - 1}-09-01")
    PPP_FY_END    = pd.Timestamp(f"{ppp_fy_num}-08-31")
    PPP_FY_QX_END = PRIOR_FY_QX_END.replace(year=PRIOR_FY_QX_END.year - 2)

    # ── Actuals windows ───────────────────────────────────────────
    ACTUALS_WINDOWS.clear()
    ACTUALS_WINDOWS.update({
        "Current Period Actuals"      : (CURRENT_FY_QX_START, CURRENT_FY_QX_END),
        "Prior Period Actuals"         : (PRIOR_FY_QX_START,   PRIOR_FY_QX_END),
        "2 Years Ago Period Actuals"   : (PP_FY_START,         PP_FY_QX_END),
        "3 Years Ago Period Actuals"   : (PPP_FY_START,        PPP_FY_QX_END),
        "Prior Year Actuals"           : (pd.Timestamp(f"{fy_num - 2}-09-01"), pd.Timestamp(f"{fy_num - 1}-08-31")),
        "2 Years Ago Actuals"          : (PP_FY_START,         PP_FY_END),
        "3 Years Ago Actuals"          : (PPP_FY_START,        PPP_FY_END),
    })

    # ── Misc derived ─────────────────────────────────────────────
    RETENTION_START_CUTOFF = pd.Timestamp(f"{fy_num - 2}-09-01")
    OCC_CURRENT_DATE       = CURRENT_Q_END
    OCC_PRIOR_DATE         = PRIOR_FY_QX_END

    # ── Alumni-specific windows (Script 5) ────────────────────────
    import calendar as _cal
    def _shift_back(ts, months):
        m = ts.month - months
        y = ts.year + (m - 1) // 12
        m = ((m - 1) % 12) + 1
        d = min(ts.day, _cal.monthrange(y, m)[1])
        return pd.Timestamp(f"{y}-{m:02d}-{d:02d}")

    # Sobriety: 7 quarters back; if Q4, extend to include Q3
    SOBRIETY_END   = _shift_back(CURRENT_Q_END,   21)
    SOBRIETY_START = _shift_back(CURRENT_Q_START, 21)
    if get_fiscal_quarter(SOBRIETY_START) == "Q4":
        SOBRIETY_START = _shift_back(SOBRIETY_START, 3)

    # LW & Housing: one year ago same Q; if Q4, use full prior FY
    if Q_LABEL == "Q4":
        LW_HOUSING_START = pd.Timestamp(f"{fy_num - 2}-09-01")
        LW_HOUSING_END   = pd.Timestamp(f"{fy_num - 1}-08-31")
    else:
        LW_HOUSING_START = CURRENT_Q_START.replace(year=CURRENT_Q_START.year - 1)
        LW_HOUSING_END   = CURRENT_Q_END.replace(year=CURRENT_Q_END.year - 1)

    # ── FY Quarters helper ────────────────────────────────────────
    FY_QUARTERS = generate_fy_quarters(fy_num)

    # ── Constants sheet mappings ──────────────────────────────────
    PROGRAM_GOALS.update({
        "Chester Renew"      : get_num("B10"), "GV Renew"           : get_num("B14"),
        "Oakland Renew"      : get_num("B18"), "Portland Renew"     : get_num("B22"),
        "San Jose Men Renew" : get_num("B26"),
    })
    PROGRAM_PROJECTIONS.update({
        "Chester Renew"      : get_num("C10"), "GV Renew"           : get_num("C14"),
        "Oakland Renew"      : get_num("C18"), "Portland Renew"     : get_num("C22"),
        "San Jose Men Renew" : get_num("C26"),
    })
    PROGRAM_CAPACITY.update({
        "Chester Renew"      : get_num("D10"), "GV Renew"           : get_num("D14"),
        "Oakland Renew"      : get_num("D18"), "Portland Renew"     : get_num("D22"),
        "San Jose Men Renew" : get_num("D26"),
    })
    PROGRAM_THEORETICAL_MAX.update({
        "Chester Renew"      : get_num("E10"), "GV Renew"           : get_num("E14"),
        "Oakland Renew"      : get_num("E18"), "Portland Renew"     : get_num("E22"),
        "San Jose Men Renew" : get_num("E26"),
    })
    OCCUPANCY_CAPACITY.update({
        "Chester Renew"      : get_num("G10"), "GV Renew"           : get_num("G14"),
        "Oakland Renew"      : get_num("G18"), "Portland Renew"     : get_num("G22"),
        "San Jose Men Renew" : get_num("G26"),
    })
    OCCUPANCY_GOAL.update({
        "Chester Renew"      : get_num("H10"), "GV Renew"           : get_num("H14"),
        "Oakland Renew"      : get_num("H18"), "Portland Renew"     : get_num("H22"),
        "San Jose Men Renew" : get_num("H26"),
    })
    OCCUPANCY_PRIOR_FY.update({
        "Chester Renew"      : get_num("I10"), "GV Renew"           : get_num("I14"),
        "Oakland Renew"      : get_num("I18"), "Portland Renew"     : get_num("I22"),
        "San Jose Men Renew" : get_num("I26"),
    })
    LW_CRITERIA.update({
        "Chester Renew"      : get_num("K9"),  "GV Renew"           : get_num("K11"),
        "Oakland Renew"      : get_num("K13"), "Portland Renew"     : get_num("K15"),
        "San Jose Men Renew" : get_num("K17"),
    })
    TP_CAPACITY.update({
        "Chester Men Turning Point"   : get_num("M10"), "Chester Women Turning Point" : get_num("M14"),
        "Oakland Men Turning Point"   : get_num("M18"), "Oakland Women Turning Point" : get_num("M22"),
        "Portland Community of Hope"  : get_num("M26"), "GV Turning Point"            : get_num("M30"),
        "Heritage Home"               : get_num("M34"), "San Jose Men Turning Point"  : get_num("M38"),
        "San Jose Youth Collective"   : get_num("M42"),
    })
    TP_GOAL.update({
        "Chester Men Turning Point"   : get_num("N10"), "Chester Women Turning Point" : get_num("N14"),
        "Oakland Men Turning Point"   : get_num("N18"), "Oakland Women Turning Point" : get_num("N22"),
        "Portland Community of Hope"  : get_num("N26"), "GV Turning Point"            : get_num("N30"),
        "Heritage Home"               : get_num("N34"), "San Jose Men Turning Point"  : get_num("N38"),
        "San Jose Youth Collective"   : get_num("N42"),
    })
    TP_PRIOR_FY.update({
        "Chester Men Turning Point"   : get_num("O10"), "Chester Women Turning Point" : get_num("O14"),
        "Oakland Men Turning Point"   : get_num("O18"), "Oakland Women Turning Point" : get_num("O22"),
        "Portland Community of Hope"  : get_num("O26"), "GV Turning Point"            : get_num("O30"),
        "Heritage Home"               : get_num("O34"), "San Jose Men Turning Point"  : get_num("O38"),
        "San Jose Youth Collective"   : get_num("O42"),
    })

    print(f"✅ All dynamic constants loaded for {FY_LABEL} {Q_LABEL}.")


# ============================================================
#  HELPER FUNCTIONS
# ============================================================

def get_fiscal_year(date):
    if pd.isnull(date): return None
    for offset in range(-5, 6):
        fy = fy_num + offset
        if pd.Timestamp(f"{fy - 1}-09-01") <= date <= pd.Timestamp(f"{fy}-08-31"):
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

print("✅ Static configuration loaded.")