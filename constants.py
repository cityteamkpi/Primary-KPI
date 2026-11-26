# ============================================================
#  constants.py — Fiscal Year, Quarter & Time Window Module
#  Reads FY & Q from 'FY Q Update' in CityTeam Impact Data/CityTeam KPIs.
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

fy_num = 2027
FY_LABEL, Q_LABEL = "FY27", "Q1"
PRIOR_FY_LABEL = PP_FY_LABEL = PPP_FY_LABEL = ""

CURRENT_Q_START = CURRENT_Q_END = CURRENT_FY_QX_START = CURRENT_FY_QX_END = None
PRIOR_FY_QX_START = PRIOR_FY_QX_END = None
PP_FY_START = PP_FY_END = PP_FY_QX_END = None
PPP_FY_START = PPP_FY_END = PPP_FY_QX_END = None
FY_START = FY_END = RETENTION_START_CUTOFF = None
OCC_CURRENT_DATE = OCC_PRIOR_DATE = None
FY_QUARTERS = []

OCC_CURRENT_LABEL, OCC_PRIOR_LABEL = "Current Period Actuals", "Prior Period Actuals"
ACTUALS_WINDOWS = {}


# ============================================================
#  LOAD FY & Q FROM 'FY Q Update' GOOGLE SHEET
# ============================================================

def _resolve_kpis_folder_id(drive_service):
    """Traverses CityTeam Impact Data -> CityTeam KPIs in Shared Drives."""
    parent_id = None
    for folder_name in ["CityTeam Impact Data", "CityTeam KPIs"]:
        q = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if parent_id:
            q += f" and '{parent_id}' in parents"
        res = drive_service.files().list(
            q=q,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        if not files:
            return resolve_folder_id(drive_service, "CityTeam KPIs", "Constants")
        parent_id = files[0]["id"]
    return parent_id


def sync_constants():
    drive_service, _, creds = get_services()
    folder_id = _resolve_kpis_folder_id(drive_service)
    sheet_name = "FY Q Update"

    target_file = find_file_id(drive_service, sheet_name, folder_id)
    if not target_file:
        raise FileNotFoundError(f"❌ '{sheet_name}' not found in CityTeam Impact Data/CityTeam KPIs")

    sheets_service = build("sheets", "v4", credentials=creds)
    sheet_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=target_file["id"], range="A1:E10"
    ).execute().get("values", [])

    # Read Fiscal Year (B3) and Quarter (C3)
    row3 = sheet_data[2] if len(sheet_data) > 2 else []
    raw_fy = row3[1] if len(row3) > 1 else None
    raw_q  = row3[2] if len(row3) > 2 else None

    if raw_fy is None or str(raw_fy).strip() == "":
        raise ValueError("❌ Fiscal Year missing at B3 in 'FY Q Update'.")
    if raw_q is None or str(raw_q).strip() == "":
        raise ValueError("❌ Quarter missing at C3 in 'FY Q Update'.")

    fy = int(str(raw_fy).strip().upper().replace("FY", ""))
    fy = fy + 2000 if fy < 100 else fy
    fy_lbl = f"FY{str(fy)[-2:]}"

    q_clean = str(raw_q).strip().upper()
    q_lbl = f"Q{q_clean}" if q_clean.isdigit() else q_clean

    # Calculate Quarter & Fiscal Year Date Boundaries
    q_months = {"Q1": (9, 11), "Q2": (12, 2), "Q3": (3, 5), "Q4": (6, 8)}
    if q_lbl not in q_months:
        raise ValueError(f"❌ Invalid Quarter '{q_lbl}' in 'FY Q Update'. Expected Q1-Q4.")

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

    globals().update({
        "fy_num": fy,
        "FY_LABEL": fy_lbl,
        "Q_LABEL": q_lbl,
        "PRIOR_FY_LABEL": f"FY{str(prior_fy_num)[-2:]}",
        "PP_FY_LABEL": f"FY{str(fy - 2)[-2:]}",
        "PPP_FY_LABEL": f"FY{str(fy - 3)[-2:]}",
        "CURRENT_Q_START": cur_q_start,
        "CURRENT_Q_END": cur_q_end,
        "CURRENT_FY_QX_START": fy_start,
        "CURRENT_FY_QX_END": cur_q_end,
        "PRIOR_FY_QX_START": pd.Timestamp(f"{prior_fy_num - 1}-09-01"),
        "PRIOR_FY_QX_END": prior_fy_qx_end,
        "PP_FY_START": pd.Timestamp(f"{fy - 3}-09-01"),
        "PP_FY_END": pd.Timestamp(f"{fy - 2}-08-31"),
        "PP_FY_QX_END": prior_fy_qx_end - pd.DateOffset(years=1),
        "PPP_FY_START": pd.Timestamp(f"{fy - 4}-09-01"),
        "PPP_FY_END": pd.Timestamp(f"{fy - 3}-08-31"),
        "PPP_FY_QX_END": prior_fy_qx_end - pd.DateOffset(years=2),
        "FY_START": fy_start,
        "FY_END": pd.Timestamp(f"{fy}-08-31"),
        "RETENTION_START_CUTOFF": pd.Timestamp(f"{fy - 2}-09-01"),
        "OCC_CURRENT_DATE": cur_q_end,
        "OCC_PRIOR_DATE": prior_fy_qx_end,
        "FY_QUARTERS": generate_fy_quarters(fy),
    })

    ACTUALS_WINDOWS.clear()
    ACTUALS_WINDOWS.update({
        "Current Period Actuals"    : (fy_start, cur_q_end),
        "Prior Period Actuals"      : (pd.Timestamp(f"{prior_fy_num - 1}-09-01"), prior_fy_qx_end),
        "2 Years Ago Period Actuals": (pd.Timestamp(f"{fy - 3}-09-01"), prior_fy_qx_end - pd.DateOffset(years=1)),
        "3 Years Ago Period Actuals": (pd.Timestamp(f"{fy - 4}-09-01"), prior_fy_qx_end - pd.DateOffset(years=2)),
        "Prior Year Actuals"        : (pd.Timestamp(f"{fy - 2}-09-01"), pd.Timestamp(f"{fy - 1}-08-31")),
        "2 Years Ago Actuals"       : (pd.Timestamp(f"{fy - 3}-09-01"), pd.Timestamp(f"{fy - 2}-08-31")),
        "3 Years Ago Actuals"       : (pd.Timestamp(f"{fy - 4}-09-01"), pd.Timestamp(f"{fy - 3}-08-31")),
    })

    print(f"✅ Loaded {FY_LABEL} {Q_LABEL} time windows from '{sheet_name}'.")


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