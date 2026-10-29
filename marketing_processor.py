import io
import sys
import socket
import warnings
import gspread
import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file

socket.setdefaulttimeout(120)
warnings.filterwarnings("ignore", category=UserWarning)

# Normalized set of intern program titles (case-insensitive)
INTERN_PROGRAM_NAMES = {
    "program graduate intern",
    "leadership module",
    "cdac trainee",
}

def log(msg):
    print(msg, flush=True)

# ============================================================
#  CONFIGURATION & DESTINATION SHEETS MAP
# ============================================================
DESTINATION_SHEET_ID = "1bvcXczlUgCpdY5jPJopjJ6lT5J03g1qrGo9zQgc93Eg"

FY_CONFIGS = {
    "FY26": {
        "start": pd.Timestamp("2025-09-01"),
        "end": pd.Timestamp("2026-08-31"),
        "sheet_name": "FY2026 Shell"
    },
    "FY27": {
        "start": pd.Timestamp("2026-09-01"),
        "end": pd.Timestamp("2027-08-31"),
        "sheet_name": "Current Fiscal Year To-Date"
    }
}

FILE_OCCUPANCY_NAME   = "Clients in All Programs Average"
FILE_RENEW_NAME       = "Renew Programs Report - Processed"
FILE_TP_NAME          = "Turning Point Report - Processed"
FILE_MRB_NAME         = "Mentorship Reunifications Baptisms - Processed"
FILE_ALUMNI_NAME      = "Alumni KPI - Processed"
FILE_ATWORK_NAME      = "@Work KPIs and Barriers - Processed"

COLUMN_INDICES = {
    "San Jose Men": 2,      # Col B
    "San Jose Women": 3,    # Col C
    "Oakland Men": 4,       # Col D
    "Oakland Women": 5,     # Col E
    "Portland Men": 6,      # Col F
    "Portland Women": 7,    # Col G
    "Chester Men": 8,       # Col H
    "Chester Women": 9,     # Col I
    "San Francisco": 10     # Col J
}

PROGRAM_MAP = {
    "San Jose Men": {
        "renew": ["San Jose Men Renew"],
        "tp": ["San Jose Men Turning Point"],
        "fwd_intern": ["GV Men's Forward"]
    },
    "San Jose Women": {
        "renew": ["GV Renew", "House of Grace"],
        "tp": ["GV Turning Point", "Heritage Home"],
        "fwd_intern": ["GV Women's Forward"]
    },
    "Oakland Men": {
        "renew": ["Oakland Renew"],
        "tp": ["Oakland Men Turning Point"],
        "fwd_intern": ["Oakland Forward Men"]
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
    },
    "San Francisco": {
        "renew": [],
        "tp": [],
        "fwd_intern": []
    }
}

# ------------------------------------------------------------
# ROW INDEX MAP
# ------------------------------------------------------------
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


def count_fy_distinct_enrollments(
    df, 
    fy_start, 
    fy_end, 
    prog_list, 
    id_col_candidates=["record id_102", "record id_2086", "record id"], 
    date_col_candidates=["start date_2090", "start date"], 
    prog_col_candidates=["program enrolling_2091", "program enrolling", "program", "intern program"]
):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}

    id_col = next((col_map[cand] for cand in id_col_candidates if cand in col_map), None)
    date_col = next((col_map[cand] for cand in date_col_candidates if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in prog_col_candidates if cand in col_map), None)

    if not date_col or not prog_col:
        return 0

    starts = pd.to_datetime(df[date_col], errors="coerce")
    df_fy = df[(starts >= fy_start) & (starts <= fy_end)].copy()
    if df_fy.empty:
        return 0

    # Case-insensitive program matching with INTERN_PROGRAM_NAMES support
    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    df_prog = df_fy[df_fy[prog_col].astype(str).str.strip().str.lower().isin(prog_set_lower)].copy()
    if df_prog.empty:
        return 0

    if id_col:
        valid_records = df_prog.dropna(subset=[id_col, date_col, prog_col])
        return len(valid_records[[id_col, date_col, prog_col]].drop_duplicates())

    return len(df_prog)


def calculate_nights_of_shelter(
    df, 
    prog_list, 
    fy_start, 
    fy_end, 
    id_col_cand=["record id_102", "record id_2086", "record id"],
    start_col_cand=["start date_2090", "start date"], 
    exit_col_cand=["exit date_2100", "exit date"], 
    prog_col_cand=["program enrolling_2091", "program enrolling", "program"]
):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in id_col_cand if cand in col_map), None)
    start_col = next((col_map[cand] for cand in start_col_cand if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in exit_col_cand if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in prog_col_cand if cand in col_map), None)

    if not start_col or not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    df_sub = df[df[prog_col].astype(str).str.strip().str.lower().isin(prog_set_lower)].copy()
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


def calculate_case_management_hours(
    df, 
    prog_list, 
    fy_start, 
    fy_end, 
    is_chester=False,
    id_col_cand=["record id_102", "record id_2086", "record id"],
    start_col_cand=["start date_2090", "start date"], 
    exit_col_cand=["exit date_2100", "exit date"], 
    prog_col_cand=["program enrolling_2091", "program enrolling", "program"]
):
    if df.empty or not prog_list:
        return 0

    col_map = {str(c).lower().strip(): c for c in df.columns}
    id_col = next((col_map[cand] for cand in id_col_cand if cand in col_map), None)
    start_col = next((col_map[cand] for cand in start_col_cand if cand in col_map), None)
    exit_col = next((col_map[cand] for cand in exit_col_cand if cand in col_map), None)
    prog_col = next((col_map[cand] for cand in prog_col_cand if cand in col_map), None)

    if not start_col or not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    df_sub = df[df[prog_col].astype(str).str.strip().str.lower().isin(prog_set_lower)].copy()
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


def count_fy_dates(df, date_col_name, prog_list, fy_start, fy_end, prog_col_candidates=["program enrolling_2091", "program enrolling", "program", "intern program"]):
    if df.empty or not prog_list:
        return 0
    
    col_map = {str(c).lower().strip(): c for c in df.columns}
    date_col = col_map.get(str(date_col_name).lower().strip())
    prog_col = next((col_map[cand] for cand in prog_col_candidates if cand in col_map), None)

    if not date_col or not prog_col:
        return 0

    prog_set_lower = {str(p).strip().lower() for p in prog_list}
    df_sub = df[df[prog_col].astype(str).str.strip().str.lower().isin(prog_set_lower)].copy()
    if df_sub.empty:
        return 0

    dates = pd.to_datetime(df_sub[date_col], errors="coerce")
    return int(((dates >= fy_start) & (dates <= fy_end)).sum())


def thread_safe_download(task):
    key, fname, folder_id = task
    drive_service, _, _ = get_services()
    fh, file_id, _ = download_drive_file(drive_service, fname, folder_id)
    return key, fh, file_id


def process_fiscal_year(fy_key, fy_info, gc, dfs):
    fy_start = fy_info["start"]
    fy_end = fy_info["end"]
    sheet_name = fy_info["sheet_name"]

    log(f"\n📊 Processing Metrics for {fy_key} ({fy_start.strftime('%Y-%m-%d')} to {fy_end.strftime('%Y-%m-%d')})...")
    log(f"   Connecting to worksheet: '{sheet_name}'...")
    doc = gc.open_by_key(DESTINATION_SHEET_ID)
    ws = doc.worksheet(sheet_name)

    cell_updates = []

    for loc_name, col_idx in COLUMN_INDICES.items():
        renew_progs = PROGRAM_MAP[loc_name]["renew"]
        tp_progs = PROGRAM_MAP[loc_name]["tp"]
        fwd_intern_progs = PROGRAM_MAP[loc_name]["fwd_intern"]
        
        # Include all intern program variations in the fwd_intern_all program list
        intern_variations = list(INTERN_PROGRAM_NAMES)
        fwd_intern_all = fwd_intern_progs + intern_variations if fwd_intern_progs else []
        all_progs = renew_progs + tp_progs + fwd_intern_all
        res_progs = renew_progs + tp_progs

        df_res_raw_combined = pd.concat([dfs["renew_raw"], dfs["tp_raw"]], ignore_index=True)

        # ----------------------------------------------------
        # ROWS 5-6: ENROLLMENTS
        # ----------------------------------------------------
        # Row 5: Residential Program New Adult Enrollments
        val_r5 = count_fy_distinct_enrollments(df_res_raw_combined, fy_start, fy_end, res_progs)
        cell_updates.append({"row": ROW_RES_NEW_ADULT_ENROLLMENTS, "col": col_idx, "val": val_r5})

        # Row 6: Residential Program New Children Enrollments
        val_r6 = 0
        col_map_res = {str(c).lower().strip(): c for c in df_res_raw_combined.columns}
        if "number of children" in col_map_res:
            n_kids_col = col_map_res["number of children"]
            d_col = next((col_map_res[c] for c in ["start date_2090", "start date"] if c in col_map_res), None)
            p_col = next((col_map_res[c] for c in ["program enrolling_2091", "program enrolling"] if c in col_map_res), None)
            if d_col and p_col:
                starts = pd.to_datetime(df_res_raw_combined[d_col], errors="coerce")
                res_progs_lower = {str(p).strip().lower() for p in res_progs}
                in_fy = df_res_raw_combined[(df_res_raw_combined[p_col].astype(str).str.strip().str.lower().isin(res_progs_lower)) & (starts >= fy_start) & (starts <= fy_end)]
                val_r6 = int(in_fy[n_kids_col].sum()) if not in_fy.empty else 0
        cell_updates.append({"row": ROW_RES_NEW_CHILD_ENROLLMENTS, "col": col_idx, "val": val_r6})

        # ----------------------------------------------------
        # ROWS 7-12: AVERAGE OCCUPANCY (Only populate for FY27)
        # ----------------------------------------------------
        if fy_key == "FY27":
            renew_lower = {str(p).strip().lower() for p in renew_progs}
            tp_lower = {str(p).strip().lower() for p in tp_progs}
            fwd_intern_lower = {str(p).strip().lower() for p in fwd_intern_all}

            val_r7  = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(renew_lower)]["Count"].sum()) if renew_progs and not dfs["occ"].empty else 0
            val_r8  = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(renew_lower)]["Children Count"].sum()) if (renew_progs and not dfs["occ"].empty and "Children Count" in dfs["occ"].columns) else 0
            val_r9  = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(tp_lower)]["Count"].sum()) if tp_progs and not dfs["occ"].empty else 0
            val_r10 = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(tp_lower)]["Children Count"].sum()) if (tp_progs and not dfs["occ"].empty and "Children Count" in dfs["occ"].columns) else 0
            val_r11 = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(fwd_intern_lower)]["Count"].sum()) if fwd_intern_all and not dfs["occ"].empty else 0
            val_r12 = int(dfs["occ"][dfs["occ"]["Program"].astype(str).str.strip().str.lower().isin(fwd_intern_lower)]["Children Count"].sum()) if (fwd_intern_all and not dfs["occ"].empty and "Children Count" in dfs["occ"].columns) else 0
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
        renew_lower = {str(p).strip().lower() for p in renew_progs}
        tp_lower = {str(p).strip().lower() for p in tp_progs}
        all_lower = {str(p).strip().lower() for p in all_progs}

        val_r18 = int(dfs["renew_grad"][(dfs["renew_grad"]["Year"] == fy_key) & (dfs["renew_grad"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(renew_lower))]["Current Period Actuals"].sum()) if renew_progs and not dfs["renew_grad"].empty and "Year" in dfs["renew_grad"].columns else 0
        cell_updates.append({"row": ROW_RENEW_GRADUATIONS, "col": col_idx, "val": val_r18})

        val_r19 = int(dfs["tp_grad"][(dfs["tp_grad"]["Year"] == fy_key) & (dfs["tp_grad"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(tp_lower))]["Current Period Actuals"].sum()) if tp_progs and not dfs["tp_grad"].empty and "Year" in dfs["tp_grad"].columns else 0
        cell_updates.append({"row": ROW_TP_GRADUATIONS, "col": col_idx, "val": val_r19})

        val_r20 = int(dfs["tp_housed"][(dfs["tp_housed"]["Year"] == fy_key) & (dfs["tp_housed"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(tp_lower))]["Current Period Actuals"].sum()) if tp_progs and not dfs["tp_housed"].empty and "Year" in dfs["tp_housed"].columns else 0
        cell_updates.append({"row": ROW_TP_SUCCESSFULLY_HOUSED, "col": col_idx, "val": val_r20})

        # ----------------------------------------------------
        # ROWS 22-26: SPIRITUAL & PROGRAM METRICS
        # ----------------------------------------------------
        val_r22 = int(dfs["reun"][(dfs["reun"]["Year"] == fy_key) & (dfs["reun"]["Program Enrolling"].astype(str).str.strip().str.lower().isin(all_lower))]["Current Period Actuals"].sum()) if all_progs and not dfs["reun"].empty and "Year" in dfs["reun"].columns else 0
        cell_updates.append({"row": ROW_FAMILIES_REUNIFIED, "col": col_idx, "val": val_r22})

        val_r23 = int(dfs["ment"][(dfs["ment"]["Year"] == fy_key) & (dfs["ment"]["Program Enrolling"].astype(str).str.strip().str.lower().isin(all_lower))]["Current Period Actuals"].sum()) if all_progs and not dfs["ment"].empty and "Year" in dfs["ment"].columns else 0
        cell_updates.append({"row": ROW_SPIRITUAL_MENTORSHIPS, "col": col_idx, "val": val_r23})

        val_r24 = int(dfs["bap"][(dfs["bap"]["Year"] == fy_key) & (dfs["bap"]["Program Enrolling"].astype(str).str.strip().str.lower().isin(all_lower))]["Current Period Actuals"].sum()) if all_progs and not dfs["bap"].empty and "Year" in dfs["bap"].columns else 0
        cell_updates.append({"row": ROW_BAPTISMS, "col": col_idx, "val": val_r24})

        # Row 25: Sobriety Rate
        if not dfs["sobriety"].empty and "Sobriety 1 Year" in dfs["sobriety"].columns:
            df_sob_sub = dfs["sobriety"][dfs["sobriety"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(all_lower)] if "Program Enrolling_2091" in dfs["sobriety"].columns else dfs["sobriety"][dfs["sobriety"]["City"] == loc_name]
            val_r25 = float(df_sob_sub["Sobriety 1 Year"].mean()) if not df_sob_sub.empty else ""
        else:
            val_r25 = ""
        cell_updates.append({"row": ROW_SOBRIETY_RATE, "col": col_idx, "val": round(val_r25, 4) if isinstance(val_r25, float) and not np.isnan(val_r25) else ""})

        # Row 26: Imputed Case Management Hours
        is_chester_loc = "Chester" in loc_name
        val_r26 = calculate_case_management_hours(df_res_raw_combined, res_progs, fy_start, fy_end, is_chester=is_chester_loc)
        cell_updates.append({"row": ROW_CASE_MANAGEMENT_HOURS, "col": col_idx, "val": val_r26})

        # ----------------------------------------------------
        # ROWS 28-35: @WORK & AFTERCARE IMPACT METRICS
        # ----------------------------------------------------
        val_r28 = count_fy_dates(dfs["atwork"], "Date Job Acquired_7009", all_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_JOBS_ATTAINED, "col": col_idx, "val": val_r28})

        val_r29 = count_fy_dates(dfs["atwork"], "Certificate/License Date", all_progs, fy_start, fy_end) if "Certificate/License Date" in dfs["atwork"].columns else 0
        cell_updates.append({"row": ROW_CREDENTIALS_ATTAINED, "col": col_idx, "val": val_r29})

        cnt_bc = count_fy_dates(dfs["atwork"], "Birth Certificate Obtained Date_6638", all_progs, fy_start, fy_end)
        cnt_id = count_fy_dates(dfs["atwork"], "State ID Obtained Date_6637", all_progs, fy_start, fy_end)
        cnt_ss = count_fy_dates(dfs["atwork"], "Social Security Card Obtained Date_6639", all_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_IDS_CARDS_DOCS, "col": col_idx, "val": (cnt_bc + cnt_id + cnt_ss)})

        val_r31 = count_fy_dates(dfs["atwork"], "Driver's License Date_5071", all_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_DRIVERS_LICENSES, "col": col_idx, "val": val_r31})

        val_r32 = count_fy_dates(dfs["atwork"], "High School Equivalency (HSE) Date_5110", all_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_HSE_ATTAINED, "col": col_idx, "val": val_r32})

        val_r33 = count_fy_dates(dfs["atwork"], "Bank Account Obtained Date_6640", all_progs, fy_start, fy_end)
        cell_updates.append({"row": ROW_BANK_ACCOUNTS, "col": col_idx, "val": val_r33})

        # Row 34: Housing Burden (1-Yr Post-Graduation)
        if not dfs["housing"].empty and "Housing Under 30% (Recent)" in dfs["housing"].columns:
            df_h_sub = dfs["housing"][dfs["housing"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(all_lower)] if "Program Enrolling_2091" in dfs["housing"].columns else dfs["housing"][dfs["housing"]["City"] == loc_name]
            val_r34 = float(df_h_sub["Housing Under 30% (Recent)"].mean()) if not df_h_sub.empty else ""
        else:
            val_r34 = ""
        cell_updates.append({"row": ROW_HOUSING_BURDEN_1YR, "col": col_idx, "val": round(val_r34, 4) if isinstance(val_r34, float) and not np.isnan(val_r34) else ""})

        # Row 35: Wage Increase (Graduation to 1-Yr Post-Graduation)
        if not dfs["lw"].empty and "Wage (Grad)" in dfs["lw"].columns and "Wage (Recent)" in dfs["lw"].columns:
            df_w_sub = dfs["lw"][dfs["lw"]["Program Enrolling_2091"].astype(str).str.strip().str.lower().isin(all_lower)] if "Program Enrolling_2091" in dfs["lw"].columns else dfs["lw"][dfs["lw"]["City"] == loc_name]
            if not df_w_sub.empty:
                w_grad = pd.to_numeric(df_w_sub["Wage (Grad)"], errors="coerce").fillna(0)
                w_rec = pd.to_numeric(df_w_sub["Wage (Recent)"], errors="coerce").fillna(0)
                wage_diff = (w_rec - w_grad)
                val_r35 = float(wage_diff.mean()) if not wage_diff.empty else ""
            else:
                val_r35 = ""
        else:
            val_r35 = ""
        cell_updates.append({"row": ROW_WAGE_INCREASE, "col": col_idx, "val": round(val_r35, 2) if isinstance(val_r35, float) and not np.isnan(val_r35) else ""})

    # Batch Update
    log(f"   Writing cells directly to Google Sheet tab '{sheet_name}'...")
    gspread_data = [{"range": gspread.utils.rowcol_to_a1(item["row"], item["col"]), "values": [[item["val"]]]} for item in cell_updates]
    ws.batch_update(gspread_data, value_input_option='USER_ENTERED')
    log(f"✅ Successfully written values directly to '{sheet_name}'!")


def run_marketing_processor(input_folder_name="Quarterly KPI Processed Data"):
    log("🚀 Starting Marketing Impact Data Processing...")

    main_drive_service, gc, _ = get_services()

    # 1. Resolve Drive Input Folder ID
    input_folder_id = resolve_folder_id(main_drive_service, input_folder_name, "Input")

    # 2. Download Input Files in Parallel
    log("⚡ Downloading source files in parallel...")
    download_tasks = [
        ("occ", FILE_OCCUPANCY_NAME, input_folder_id),
        ("renew", FILE_RENEW_NAME, input_folder_id),
        ("tp", FILE_TP_NAME, input_folder_id),
        ("mrb", FILE_MRB_NAME, input_folder_id),
        ("alum", FILE_ALUMNI_NAME, input_folder_id),
        ("atwork", FILE_ATWORK_NAME, input_folder_id)
    ]

    downloads = {}
    with ThreadPoolExecutor(max_workers=len(download_tasks)) as executor:
        for key, fh, file_id in executor.map(thread_safe_download, download_tasks):
            downloads[key] = (fh, file_id)
            log(f"   ✓ Downloaded: {key}")

    fh_occ, _ = downloads["occ"]
    fh_renew, _ = downloads["renew"]
    fh_tp, _ = downloads["tp"]
    fh_mrb, _ = downloads["mrb"]
    fh_alum, _ = downloads["alum"]
    fh_atwork, _ = downloads["atwork"]

    # 3. Load Source DataFrames
    log("   Parsing Excel data into DataFrames...")
    df_occ = pd.read_excel(fh_occ, sheet_name="Latest Occupancy")

    xls_renew = pd.ExcelFile(fh_renew)
    renew_sheets = xls_renew.sheet_names
    renew_raw_sheet = "Renew Raw Data" if "Renew Raw Data" in renew_sheets else ("Raw Data" if "Raw Data" in renew_sheets else renew_sheets[0])
    df_renew_raw  = pd.read_excel(xls_renew, sheet_name=renew_raw_sheet)
    df_renew_occ  = pd.read_excel(xls_renew, sheet_name="Occupancy") if "Occupancy" in renew_sheets else df_renew_raw
    df_renew_grad = pd.read_excel(xls_renew, sheet_name="Renew Graduates") if "Renew Graduates" in renew_sheets else pd.DataFrame()

    xls_tp = pd.ExcelFile(fh_tp)
    tp_sheets = xls_tp.sheet_names
    tp_raw_sheet = "Turning Point Raw Data" if "Turning Point Raw Data" in tp_sheets else ("Raw Data" if "Raw Data" in tp_sheets else tp_sheets[0])
    df_tp_raw    = pd.read_excel(xls_tp, sheet_name=tp_raw_sheet)
    df_tp_occ    = pd.read_excel(xls_tp, sheet_name="Occupancy") if "Occupancy" in tp_sheets else df_tp_raw
    df_tp_grad   = pd.read_excel(xls_tp, sheet_name="Turning Point Graduates") if "Turning Point Graduates" in tp_sheets else pd.DataFrame()
    df_tp_housed = pd.read_excel(xls_tp, sheet_name="Successfully Housed") if "Successfully Housed" in tp_sheets else pd.DataFrame()

    xls_mrb = pd.ExcelFile(fh_mrb)
    mrb_sheets = xls_mrb.sheet_names
    df_ment = pd.read_excel(xls_mrb, sheet_name="Mentorship Processed") if "Mentorship Processed" in mrb_sheets else pd.DataFrame()
    df_reun = pd.read_excel(xls_mrb, sheet_name="Reunifications Processed") if "Reunifications Processed" in mrb_sheets else pd.DataFrame()
    df_bap  = pd.read_excel(xls_mrb, sheet_name="Baptisms Processed") if "Baptisms Processed" in mrb_sheets else pd.DataFrame()

    xls_alum = pd.ExcelFile(fh_alum)
    alum_sheets = xls_alum.sheet_names
    df_sobriety = pd.read_excel(xls_alum, sheet_name="Sobriety") if "Sobriety" in alum_sheets else pd.DataFrame()
    df_housing  = pd.read_excel(xls_alum, sheet_name="Housing") if "Housing" in alum_sheets else pd.DataFrame()
    df_lw       = pd.read_excel(xls_alum, sheet_name="Living Wage") if "Living Wage" in alum_sheets else pd.DataFrame()

    xls_atwork = pd.ExcelFile(fh_atwork)
    atwork_sheets = xls_atwork.sheet_names
    df_atwork  = pd.read_excel(xls_atwork, sheet_name="@Work Raw Data") if "@Work Raw Data" in atwork_sheets else pd.DataFrame()

    dfs = {
        "occ": df_occ,
        "renew_raw": df_renew_raw,
        "renew_occ": df_renew_occ,
        "renew_grad": df_renew_grad,
        "tp_raw": df_tp_raw,
        "tp_occ": df_tp_occ,
        "tp_grad": df_tp_grad,
        "tp_housed": df_tp_housed,
        "ment": df_ment,
        "reun": df_reun,
        "bap": df_bap,
        "sobriety": df_sobriety,
        "housing": df_housing,
        "lw": df_lw,
        "atwork": df_atwork
    }

    # 4. Iterate over Fiscal Year configurations (FY26 & FY27)
    for fy_key, fy_info in FY_CONFIGS.items():
        process_fiscal_year(fy_key, fy_info, gc, dfs)

    log("\n🎉 Marketing Processor Execution Complete for FY26 and FY27!")


if __name__ == "__main__":
    run_marketing_processor()