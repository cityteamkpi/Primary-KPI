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

def log(msg):
    print(msg, flush=True)

# ============================================================
#  CONFIGURATION & EXACT DESTINATION SHEET ID
# ============================================================
TARGET_FY = "FY27"
FY27_START = pd.Timestamp("2026-09-01")
FY27_END   = pd.Timestamp("2027-08-31")

FILE_OCCUPANCY_NAME   = "Clients in All Programs Average"
FILE_RENEW_NAME       = "Renew Programs Report - Processed"
FILE_TP_NAME          = "Turning Point Report - Processed"
FILE_MRB_NAME         = "Mentorship Reunifications Baptisms - Processed"
FILE_ALUMNI_NAME      = "Alumni KPI - Processed"
FILE_ATWORK_NAME      = "@Work KPIs and Barriers - Processed"

# Exact Google Sheet ID:
DESTINATION_SHEET_ID  = "1bvcXczlUgCpdY5jPJopjJ6lT5J03g1qrGo9zQgc93Eg"
TARGET_SHEET_NAME     = "Current Fiscal Year To-Date"

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
        "fwd_intern": ["San Jose Men Forward", "Program Graduate Intern"]
    },
    "San Jose Women": {
        "renew": ["GV Renew"],
        "tp": ["GV Turning Point", "Heritage Home", "San Jose Youth Collective"],
        "fwd_intern": ["San Jose Women Forward", "Program Graduate Intern"]
    },
    "Oakland Men": {
        "renew": ["Oakland Renew"],
        "tp": ["Oakland Men Turning Point"],
        "fwd_intern": ["Oakland Men Forward", "Program Graduate Intern"]
    },
    "Oakland Women": {
        "renew": [],
        "tp": ["Oakland Women Turning Point"],
        "fwd_intern": ["Oakland Women Forward", "Program Graduate Intern"]
    },
    "Portland Men": {
        "renew": ["Portland Renew"],
        "tp": ["Portland Men Turning Point"],
        "fwd_intern": ["Portland Men Forward", "Program Graduate Intern"]
    },
    "Portland Women": {
        "renew": [],
        "tp": ["Portland Community of Hope"],
        "fwd_intern": ["Portland Women Forward", "Program Graduate Intern"]
    },
    "Chester Men": {
        "renew": ["Chester Renew"],
        "tp": ["Chester Men Turning Point"],
        "fwd_intern": ["Chester Men Forward", "Program Graduate Intern"]
    },
    "Chester Women": {
        "renew": [],
        "tp": ["Chester Women Turning Point", "Chester Women's Turning Point"],
        "fwd_intern": ["Chester Women Forward", "Program Graduate Intern"]
    },
    "San Francisco": {
        "renew": [],
        "tp": [],
        "fwd_intern": []
    }
}

# ------------------------------------------------------------
# UPDATED ROW INDEX MAP
# ------------------------------------------------------------
# Rows 5 - 12: Enrollments & Occupancy
ROW_RES_NEW_ADULT_ENROLLMENTS      = 5
ROW_RES_NEW_CHILD_ENROLLMENTS      = 6
ROW_RENEW_ADULT_AVG_OCCUPANCY      = 7
ROW_RENEW_CHILD_AVG_OCCUPANCY      = 8
ROW_TP_ADULT_AVG_OCCUPANCY         = 9
ROW_TP_CHILD_AVG_OCCUPANCY         = 10
ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY = 11
ROW_FWD_INTERN_CHILD_AVG_OCCUPANCY = 12

# Rows 14 - 16: Nights of Shelter
ROW_RENEW_ADULT_NIGHTS             = 14
ROW_TP_ADULT_NIGHTS                = 15
ROW_FWD_INTERN_NIGHTS              = 16

# Rows 18 - 20: Graduations & Housing Placement
ROW_RENEW_GRADUATIONS              = 18
ROW_TP_GRADUATIONS                 = 19
ROW_TP_SUCCESSFULLY_HOUSED         = 20

# Rows 22 - 26: Spiritual & Program Metrics
ROW_FAMILIES_REUNIFIED             = 22
ROW_SPIRITUAL_MENTORSHIPS          = 23
ROW_BAPTISMS                       = 24
ROW_SOBRIETY_RATE                  = 25
ROW_CASE_MANAGEMENT_HOURS          = 26

# Rows 28 - 35: @Work & Aftercare Impact Metrics
ROW_JOBS_ATTAINED                  = 28
ROW_CREDENTIALS_ATTAINED           = 29
ROW_IDS_CARDS_DOCS                 = 30
ROW_DRIVERS_LICENSES               = 31
ROW_HSE_ATTAINED                   = 32
ROW_BANK_ACCOUNTS                  = 33
ROW_HOUSING_BURDEN_1YR             = 34
ROW_WAGE_INCREASE                  = 35


def calculate_nights_of_shelter(df, prog_list, start_col="Start Date_2090", exit_col="Exit Date_2100", prog_col="Program Enrolling_2091"):
    if df.empty or not prog_list or prog_col not in df.columns:
        return 0

    df_sub = df[df[prog_col].isin(prog_list)].copy()
    if df_sub.empty:
        return 0

    today = pd.Timestamp.now().normalize()
    effective_end_limit = min(today, FY27_END)

    starts = pd.to_datetime(df_sub[start_col], errors="coerce")
    exits = pd.to_datetime(df_sub[exit_col], errors="coerce").fillna(today)

    overlap_start = starts.clip(lower=FY27_START)
    overlap_end = exits.clip(upper=effective_end_limit)

    nights = (overlap_end - overlap_start).dt.days
    nights = nights.where((nights > 0) & (starts <= effective_end_limit) & (exits >= FY27_START), 0)

    return int(nights.sum())


def count_fy27_dates(df, date_col, prog_list, prog_col="Program Enrolling_2091"):
    if df.empty or date_col not in df.columns or not prog_list or prog_col not in df.columns:
        return 0
    
    df_sub = df[df[prog_col].isin(prog_list)].copy()
    if df_sub.empty:
        return 0

    dates = pd.to_datetime(df_sub[date_col], errors="coerce")
    return int(((dates >= FY27_START) & (dates <= FY27_END)).sum())


def thread_safe_download(task):
    key, fname, folder_id = task
    drive_service, _, _ = get_services()
    fh, file_id, _ = download_drive_file(drive_service, fname, folder_id)
    return key, fh, file_id


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
    df_renew_occ  = pd.read_excel(xls_renew, sheet_name="Occupancy") if "Occupancy" in xls_renew.sheet_names else pd.DataFrame()
    df_renew_grad = pd.read_excel(xls_renew, sheet_name="Renew Graduates") if "Renew Graduates" in xls_renew.sheet_names else pd.DataFrame()

    xls_tp = pd.ExcelFile(fh_tp)
    df_tp_occ    = pd.read_excel(xls_tp, sheet_name="Occupancy") if "Occupancy" in xls_tp.sheet_names else pd.DataFrame()
    df_tp_grad   = pd.read_excel(xls_tp, sheet_name="Turning Point Graduates") if "Turning Point Graduates" in xls_tp.sheet_names else pd.DataFrame()
    df_tp_housed = pd.read_excel(xls_tp, sheet_name="Successfully Housed") if "Successfully Housed" in xls_tp.sheet_names else pd.DataFrame()

    xls_mrb = pd.ExcelFile(fh_mrb)
    df_ment = pd.read_excel(xls_mrb, sheet_name="Mentorship Processed") if "Mentorship Processed" in xls_mrb.sheet_names else pd.DataFrame()
    df_reun = pd.read_excel(xls_mrb, sheet_name="Reunifications Processed") if "Reunifications Processed" in xls_mrb.sheet_names else pd.DataFrame()
    df_bap  = pd.read_excel(xls_mrb, sheet_name="Baptisms Processed") if "Baptisms Processed" in xls_mrb.sheet_names else pd.DataFrame()

    xls_alum = pd.ExcelFile(fh_alum)
    df_sobriety = pd.read_excel(xls_alum, sheet_name="Sobriety") if "Sobriety" in xls_alum.sheet_names else pd.DataFrame()
    df_housing  = pd.read_excel(xls_alum, sheet_name="Housing") if "Housing" in xls_alum.sheet_names else pd.DataFrame()
    df_lw       = pd.read_excel(xls_alum, sheet_name="Living Wage") if "Living Wage" in xls_alum.sheet_names else pd.DataFrame()

    xls_atwork = pd.ExcelFile(fh_atwork)
    df_atwork  = pd.read_excel(xls_atwork, sheet_name="@Work Raw Data") if "@Work Raw Data" in xls_atwork.sheet_names else pd.DataFrame()

    # 4. Connect directly to target Google Sheet via gspread
    log(f"   Connecting directly to Google Sheet ID {DESTINATION_SHEET_ID} via gspread...")
    doc = gc.open_by_key(DESTINATION_SHEET_ID)
    ws = doc.worksheet(TARGET_SHEET_NAME)

    # 5. Build Cell Batch Updates
    log(f"   Calculating metrics for Target Fiscal Year: {TARGET_FY}...")
    cell_updates = []

    for loc_name, col_idx in COLUMN_INDICES.items():
        renew_progs = PROGRAM_MAP[loc_name]["renew"]
        tp_progs = PROGRAM_MAP[loc_name]["tp"]
        fwd_intern_progs = PROGRAM_MAP[loc_name]["fwd_intern"]
        all_progs = renew_progs + tp_progs + fwd_intern_progs

        # ----------------------------------------------------
        # ROWS 5-12: ENROLLMENTS & OCCUPANCY
        # ----------------------------------------------------
        # Row 5: Residential Program New Adult Enrollments
        if not df_renew_occ.empty and "Start Date_2090" in df_renew_occ.columns and renew_progs:
            starts = pd.to_datetime(df_renew_occ["Start Date_2090"], errors="coerce")
            val_r5 = int(((df_renew_occ["Program Enrolling_2091"].isin(renew_progs)) & (starts >= FY27_START) & (starts <= FY27_END)).sum())
        else:
            val_r5 = 0
        cell_updates.append({"row": ROW_RES_NEW_ADULT_ENROLLMENTS, "col": col_idx, "val": val_r5})

        # Row 6: Residential Program New Children Enrollments
        val_r6 = int(df_renew_occ["Number of Children"].sum()) if (not df_renew_occ.empty and "Number of Children" in df_renew_occ.columns and loc_name == "San Jose Women") else 0
        cell_updates.append({"row": ROW_RES_NEW_CHILD_ENROLLMENTS, "col": col_idx, "val": val_r6})

        # Row 7 & 8: Renew Adult & Child Average Occupancy
        val_r7 = int(df_occ[df_occ["Program"].isin(renew_progs)]["Count"].sum()) if renew_progs else 0
        cell_updates.append({"row": ROW_RENEW_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r7})

        val_r8 = int(df_occ[df_occ["Program"].isin(renew_progs)]["Children Count"].sum()) if (renew_progs and "Children Count" in df_occ.columns) else 0
        cell_updates.append({"row": ROW_RENEW_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r8})

        # Row 9 & 10: Turning Point Adult & Child Average Occupancy
        val_r9 = int(df_occ[df_occ["Program"].isin(tp_progs)]["Count"].sum()) if tp_progs else 0
        cell_updates.append({"row": ROW_TP_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r9})

        val_r10 = int(df_occ[df_occ["Program"].isin(tp_progs)]["Children Count"].sum()) if (tp_progs and "Children Count" in df_occ.columns) else 0
        cell_updates.append({"row": ROW_TP_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r10})

        # Row 11 & 12: Forward + Intern Program Adult & Child Average Occupancy
        val_r11 = int(df_occ[df_occ["Program"].isin(fwd_intern_progs)]["Count"].sum()) if fwd_intern_progs else 0
        cell_updates.append({"row": ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r11})

        val_r12 = int(df_occ[df_occ["Program"].isin(fwd_intern_progs)]["Children Count"].sum()) if (fwd_intern_progs and "Children Count" in df_occ.columns) else 0
        cell_updates.append({"row": ROW_FWD_INTERN_CHILD_AVG_OCCUPANCY, "col": col_idx, "val": val_r12})

        # ----------------------------------------------------
        # ROWS 14-16: NIGHTS OF SHELTER
        # ----------------------------------------------------
        val_r14 = calculate_nights_of_shelter(df_renew_occ, renew_progs)
        cell_updates.append({"row": ROW_RENEW_ADULT_NIGHTS, "col": col_idx, "val": val_r14})

        val_r15 = calculate_nights_of_shelter(df_tp_occ, tp_progs)
        cell_updates.append({"row": ROW_TP_ADULT_NIGHTS, "col": col_idx, "val": val_r15})

        val_r16 = calculate_nights_of_shelter(df_renew_occ, fwd_intern_progs) if not df_renew_occ.empty else 0
        cell_updates.append({"row": ROW_FWD_INTERN_NIGHTS, "col": col_idx, "val": val_r16})

        # ----------------------------------------------------
        # ROWS 18-20: GRADUATIONS & HOUSING
        # ----------------------------------------------------
        val_r18 = int(df_renew_grad[(df_renew_grad["Year"] == TARGET_FY) & (df_renew_grad["Program Enrolling_2091"].isin(renew_progs))]["Current Period Actuals"].sum()) if renew_progs and "Year" in df_renew_grad.columns else 0
        cell_updates.append({"row": ROW_RENEW_GRADUATIONS, "col": col_idx, "val": val_r18})

        val_r19 = int(df_tp_grad[(df_tp_grad["Year"] == TARGET_FY) & (df_tp_grad["Program Enrolling_2091"].isin(tp_progs))]["Current Period Actuals"].sum()) if tp_progs and "Year" in df_tp_grad.columns else 0
        cell_updates.append({"row": ROW_TP_GRADUATIONS, "col": col_idx, "val": val_r19})

        val_r20 = int(df_tp_housed[(df_tp_housed["Year"] == TARGET_FY) & (df_tp_housed["Program Enrolling_2091"].isin(tp_progs))]["Current Period Actuals"].sum()) if tp_progs and "Year" in df_tp_housed.columns else 0
        cell_updates.append({"row": ROW_TP_SUCCESSFULLY_HOUSED, "col": col_idx, "val": val_r20})

        # ----------------------------------------------------
        # ROWS 22-26: SPIRITUAL & PROGRAM METRICS
        # ----------------------------------------------------
        val_r22 = int(df_reun[(df_reun["Year"] == TARGET_FY) & (df_reun["Program Enrolling"].isin(all_progs))]["Current Period Actuals"].sum()) if all_progs and "Year" in df_reun.columns and not df_reun.empty else 0
        cell_updates.append({"row": ROW_FAMILIES_REUNIFIED, "col": col_idx, "val": val_r22})

        val_r23 = int(df_ment[(df_ment["Year"] == TARGET_FY) & (df_ment["Program Enrolling"].isin(all_progs))]["Current Period Actuals"].sum()) if all_progs and "Year" in df_ment.columns and not df_ment.empty else 0
        cell_updates.append({"row": ROW_SPIRITUAL_MENTORSHIPS, "col": col_idx, "val": val_r23})

        val_r24 = int(df_bap[(df_bap["Year"] == TARGET_FY) & (df_bap["Program Enrolling"].isin(all_progs))]["Current Period Actuals"].sum()) if all_progs and "Year" in df_bap.columns and not df_bap.empty else 0
        cell_updates.append({"row": ROW_BAPTISMS, "col": col_idx, "val": val_r24})

        # Row 25: Sobriety Rate
        if not df_sobriety.empty and "Sobriety 1 Year" in df_sobriety.columns:
            df_sob_sub = df_sobriety[df_sobriety["Program Enrolling_2091"].isin(all_progs)] if "Program Enrolling_2091" in df_sobriety.columns else df_sobriety[df_sobriety["City"] == loc_name]
            val_r25 = float(df_sob_sub["Sobriety 1 Year"].mean()) if not df_sob_sub.empty else ""
        else:
            val_r25 = ""
        cell_updates.append({"row": ROW_SOBRIETY_RATE, "col": col_idx, "val": round(val_r25, 4) if isinstance(val_r25, float) and not np.isnan(val_r25) else ""})

        # Row 26: Case Management Hours
        val_r26 = count_fy27_dates(df_atwork, "Case Management Date", all_progs) if "Case Management Date" in df_atwork.columns else 0
        cell_updates.append({"row": ROW_CASE_MANAGEMENT_HOURS, "col": col_idx, "val": val_r26})

        # ----------------------------------------------------
        # ROWS 28-35: @WORK & AFTERCARE IMPACT METRICS
        # ----------------------------------------------------
        val_r28 = count_fy27_dates(df_atwork, "Date Job Acquired_7009", all_progs)
        cell_updates.append({"row": ROW_JOBS_ATTAINED, "col": col_idx, "val": val_r28})

        val_r29 = count_fy27_dates(df_atwork, "Certificate/License Date", all_progs) if "Certificate/License Date" in df_atwork.columns else 0
        cell_updates.append({"row": ROW_CREDENTIALS_ATTAINED, "col": col_idx, "val": val_r29})

        cnt_bc = count_fy27_dates(df_atwork, "Birth Certificate Obtained Date_6638", all_progs)
        cnt_id = count_fy27_dates(df_atwork, "State ID Obtained Date_6637", all_progs)
        cnt_ss = count_fy27_dates(df_atwork, "Social Security Card Obtained Date_6639", all_progs)
        cell_updates.append({"row": ROW_IDS_CARDS_DOCS, "col": col_idx, "val": (cnt_bc + cnt_id + cnt_ss)})

        val_r31 = count_fy27_dates(df_atwork, "Driver's License Date_5071", all_progs)
        cell_updates.append({"row": ROW_DRIVERS_LICENSES, "col": col_idx, "val": val_r31})

        val_r32 = count_fy27_dates(df_atwork, "High School Equivalency (HSE) Date_5110", all_progs)
        cell_updates.append({"row": ROW_HSE_ATTAINED, "col": col_idx, "val": val_r32})

        val_r33 = count_fy27_dates(df_atwork, "Bank Account Obtained Date_6640", all_progs)
        cell_updates.append({"row": ROW_BANK_ACCOUNTS, "col": col_idx, "val": val_r33})

        # Row 34: Housing Burden (1-Yr Post-Graduation)
        if not df_housing.empty and "Housing Under 30% (Recent)" in df_housing.columns:
            df_h_sub = df_housing[df_housing["Program Enrolling_2091"].isin(all_progs)] if "Program Enrolling_2091" in df_housing.columns else df_housing[df_housing["City"] == loc_name]
            val_r34 = float(df_h_sub["Housing Under 30% (Recent)"].mean()) if not df_h_sub.empty else ""
        else:
            val_r34 = ""
        cell_updates.append({"row": ROW_HOUSING_BURDEN_1YR, "col": col_idx, "val": round(val_r34, 4) if isinstance(val_r34, float) and not np.isnan(val_r34) else ""})

        # Row 35: Wage Increase (Graduation to 1-Yr Post-Graduation)
        if not df_lw.empty and "Wage (Grad)" in df_lw.columns and "Wage (Recent)" in df_lw.columns:
            df_w_sub = df_lw[df_lw["Program Enrolling_2091"].isin(all_progs)] if "Program Enrolling_2091" in df_lw.columns else df_lw[df_lw["City"] == loc_name]
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

    # 6. Direct Batch Update via gspread with USER_ENTERED option
    log("   Writing cells directly to Google Sheet...")
    gspread_data = [{"range": gspread.utils.rowcol_to_a1(item["row"], item["col"]), "values": [[item["val"]]]} for item in cell_updates]
    ws.batch_update(gspread_data, value_input_option='USER_ENTERED')

    log(f"✅ Successfully written values directly to Google Sheet tab '{TARGET_SHEET_NAME}'!")
    log("🎉 Marketing Processor Execution Complete!")


if __name__ == "__main__":
    run_marketing_processor()