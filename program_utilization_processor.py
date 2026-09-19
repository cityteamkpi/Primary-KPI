# =========================================================================
# This script processes program utilization.
#
# 1. Reads "Clients in All Programs.xlsx.xlsx"
# 2. Creates "Clients in All Programs - Processed" Google Sheet 
#       - with Renew, Men's TP and Women's TP tabs
# =========================================================================

import time
import warnings
import pandas as pd
import gspread
from zoneinfo import ZoneInfo
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime
from googleapiclient.discovery import build as _build
from auth_utils import get_services
from drive_utils import download_drive_file, resolve_folder_id, find_file_id, load_raw, write_tab

# Suppress openpyxl warnings regarding styles and validation metadata
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

MENTOR_PROGRAMS = {
    "Chester Renew", "GV Renew", "Oakland Renew", "Portland Renew", "San Jose Men Renew",
    "Chester Men Turning Point", "Chester Women Turning Point",
    "Oakland Men Turning Point", "Oakland Women Turning Point",
    "Portland Men Turning Point",
    "GV Turning Point", "Heritage Home",
    "San Jose Men Turning Point", "San Jose Youth Collective",
}

PROGRAM_CATEGORY = {
    # Renew
    "Chester Renew": "Renew",
    "GV Renew": "Renew",
    "GV Renew Number of Children": "Renew",
    "Oakland Renew": "Renew",
    "Portland Renew": "Renew",
    "San Jose Men Renew": "Renew",
    # Men Turning Point
    "Chester Men Turning Point": "Men Turning Point",
    "Oakland Men Turning Point": "Men Turning Point",
    "San Jose Men Turning Point": "Men Turning Point",
    "Portland Men Turning Point": "Men Turning Point",
    # Women Turning Point
    "Chester Women Turning Point": "Women Turning Point",
    "Chester Women Turning Point Number of Children": "Women Turning Point",
    "GV Turning Point": "Women Turning Point",
    "GV Turning Point Number of Children": "Women Turning Point",
    "Heritage Home": "Women Turning Point",
    "Heritage Home Number of Children": "Women Turning Point",
    "Oakland Women Turning Point": "Women Turning Point",
    "Oakland Women Turning Point Number of Children": "Women Turning Point",
    "Portland Community of Hope": "Women Turning Point",
    "Portland Community of Hope Number of Children": "Women Turning Point",
    "San Jose Youth Collective": "Women Turning Point",
    "San Jose Youth Collective Number of Children": "Women Turning Point",
    # Forward
    "Chester Forward": "Forward",
    "GV Men's Forward": "Forward",
    "GV Women's Forward": "Forward",
    "GV Women's Forward Number of Children": "Forward",
    "San Jose Women's House of Light": "Forward",
    "San Jose Women's House of Light Number of Children": "Forward",
    "Oakland Forwad Men": "Forward",
    "Oakland Forward Women": "Forward",
    "Oakland Forward Women Number of Children": "Forward",
    "Portland Forward": "Forward",
    # Other
    "SJ WP Transition Phase": "Other",
    "GV Immediate Temporary Housing": "Other",
    # APH
    "GV APH": "APH",
    # Elevate
    "Elevate Mayfair Level 1": "Elevate",
    "Elevate Mt View Level 1": "Elevate",
    "Elevate Redwood City Level 1": "Elevate",
}

def run_utilization_processing(
    input_report_file       = 'Clients in All Programs.xlsx',
    input_mentorship_file   = 'Mentorship.xlsx',
    input_utilization_file  = 'Clients in All Programs - Processed',
    input_folder_name       = None,
    output_folder_name      = None
):

    # =========================================================================
    # Configuration CONSTANTS
    # =========================================================================
    RAW_DATA_HEADER_ROW = 1
    RAW_DATA_START_COL = 1
    COL_PROGRAM_HEADER = "Program Enrolling_2091"

    # Mentorship file constants
    MENTOR_HEADER_ROW  = 3  # row 4 = index 3
    MENTOR_START_COL   = 1  # starts at column B

    # =========================================================================
    # 1. Initialize Google Drive API Service
    # =========================================================================
    drive_service, gc, _ = get_services()

    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Download main data file
    # =========================================================================
    print(f"Downloading {input_report_file}...")
    report_stream, _, _ = download_drive_file(drive_service, input_report_file, input_folder_id)
    df_raw = load_raw(report_stream, header_row=RAW_DATA_HEADER_ROW, start_col=RAW_DATA_START_COL)

    # =========================================================================
    # 2b. Download Mentorship file
    # =========================================================================
    df_mentorship = None
    try:
        print(f"Downloading {input_mentorship_file}...")
        m_stream, _, _ = download_drive_file(drive_service, input_mentorship_file, input_folder_id)
        df_mentorship = load_raw(m_stream, header_row=MENTOR_HEADER_ROW, start_col=MENTOR_START_COL)
        print(f"✅ Mentorship loaded: {len(df_mentorship)} rows")
    except Exception as e:
        print(f"⚠️  Could not load Mentorship.xlsx: {e}")

    # =========================================================================
    # 3. Find Processed Google Sheet
    # =========================================================================
    util_file = find_file_id(drive_service, input_utilization_file, output_folder_id, "application/vnd.google-apps.spreadsheet")
    if not util_file:
        raise FileNotFoundError(f"Google Sheet not found: {input_utilization_file}")
    utilization_file_id = util_file['id']

    # =========================================================================
    # 4. Calculate client enrollment counts
    # =========================================================================
    if COL_PROGRAM_HEADER in df_raw.columns:
        program_col = COL_PROGRAM_HEADER
    elif any(col.startswith('Program Enrolling') for col in df_raw.columns):
        program_col = [col for col in df_raw.columns if col.startswith('Program Enrolling')][0]
    elif 'Program Enrolling' in df_raw.columns:
        program_col = 'Program Enrolling'
    elif len(df_raw.columns) >= 4:
        program_col = df_raw.columns[3]
    else:
        program_col = df_raw.columns[0]

    spreadsheet = gc.open_by_key(utilization_file_id)
    tabs_to_update = ["Renew", "Men Turning Point", "Women Turning Point", "Forward", "Program Interns", "Other", "APH", "Elevate"]

    program_counts = df_raw[program_col].astype(str).str.strip().value_counts()

    # =========================================================================
    # 4b. Calculate children counts for women's programs
    # =========================================================================
    WOMEN_PROGRAMS = [
        "Chester Women Turning Point", "Oakland Women Turning Point",
        "GV Turning Point", "Heritage Home", "San Jose Youth Collective",
        "Portland Community of Hope",
        "GV Women's Forward", "San Jose Women's House of Light", "Oakland Forward Women",
        "GV Renew"
    ]
    COL_CHILD = "Name of Child_6313"

    def count_children_in_cell(val):
        import re as _re
        if pd.isna(val) or str(val).strip() == "": return 0
        parts = _re.split(r'[,&]|\band\b', str(val).strip(), flags=_re.IGNORECASE)
        return len([p for p in parts if p.strip()])

    children_counts = {}
    if COL_CHILD in df_raw.columns:
        for prog in WOMEN_PROGRAMS:
            prog_df = df_raw[df_raw[program_col].astype(str).str.strip() == prog]
            children_counts[prog] = int(prog_df[COL_CHILD].apply(count_children_in_cell).sum())

    # =========================================================================
    # 5. Process each tab in the utilization tracking sheet
    # =========================================================================
    today_date = datetime.now(ZoneInfo("America/Los_Angeles")).strftime('%-m/%-d/%y')

    for tab_name in tabs_to_update:
        try:
            sheet = spreadsheet.worksheet(tab_name)
            print(f"Updating tab: {tab_name}")
        except gspread.exceptions.WorksheetNotFound:
            print(f"Warning: Tab '{tab_name}' not found. Skipping.")
            continue

        df_sheet = get_as_dataframe(sheet)
        df_sheet = df_sheet.dropna(how='all')

        if not df_sheet.columns.empty:
            df_sheet = df_sheet.set_index(df_sheet.columns[0])
            df_sheet = df_sheet.loc[:, ~df_sheet.columns.astype(str).str.match(r'^Unnamed')]

        df_sheet.index.name = 'Date'
        df_sheet = df_sheet.copy()

        valid_dates = pd.to_datetime(df_sheet.index, errors='coerce', format='mixed')
        mask = valid_dates.notna()
        df_sheet = df_sheet[mask]
        if not df_sheet.empty:
            df_sheet.index = valid_dates[mask].strftime('%-m/%-d/%y')

        # Ensure children count columns exist right next to women's program columns
        cols_list = list(df_sheet.columns)
        for prog in WOMEN_PROGRAMS:
            children_col = prog + " Number of Children"
            if prog in cols_list and children_col not in cols_list:
                insert_idx = cols_list.index(prog) + 1
                df_sheet.insert(insert_idx, children_col, 0)
                cols_list = list(df_sheet.columns)  # update after insert

        # Update today's row
        for program in df_sheet.columns:
            clean_prog = str(program).strip()
            if clean_prog.endswith(" Number of Children"): continue
            count = int(program_counts.get(clean_prog, 0))
            df_sheet.loc[today_date, program] = count
            children_col = clean_prog + " Number of Children"
            if clean_prog in WOMEN_PROGRAMS and children_col in df_sheet.columns:
                df_sheet.loc[today_date, children_col] = children_counts.get(clean_prog, 0)

        df_sheet.index.name = 'Date'
        write_tab(spreadsheet, tab_name, df_sheet.reset_index())

    # =========================================================================
    # 6. Process Mentorship tab
    # =========================================================================
    if df_mentorship is not None:
        try:
            # Column names in Mentorship.xlsx
            COL_START     = "Start Date_2090"
            COL_MENTOR_DT = "Start Date for Mentoring_4365"
            COL_RECORD_ID = "Record Id_102"

            df_m = df_mentorship.copy()

            # Parse dates
            if COL_START in df_m.columns:
                df_m[COL_START] = pd.to_datetime(df_m[COL_START], errors="coerce")
            if COL_MENTOR_DT in df_m.columns:
                df_m[COL_MENTOR_DT] = pd.to_datetime(df_m[COL_MENTOR_DT], errors="coerce")

            # Filter: Acquired Spiritual Mentor = Yes
            COL_MENTOR = "Acquired Spiritual Mentor_5136"
            if COL_MENTOR in df_m.columns:
                df_m = df_m[df_m[COL_MENTOR].astype(str).str.strip() == "Yes"]

            # Filter: Exit Date is null (still in program)
            COL_EXIT = "Exit Date_2100"
            if COL_EXIT in df_m.columns:
                df_m[COL_EXIT] = pd.to_datetime(df_m[COL_EXIT], errors="coerce")
                df_m = df_m[df_m[COL_EXIT].isna()]

            # Filter: in program > 60 days
            if COL_START in df_m.columns:
                today = pd.Timestamp.now()
                df_m = df_m[df_m[COL_START].notna() & ((today - df_m[COL_START]).dt.days > 60)]

            # Dedup: one row per Record Id — keep latest Start Date_2090
            # If still duplicates, use Start Date for Mentoring_4365
            if COL_RECORD_ID in df_m.columns and COL_START in df_m.columns:
                df_m = df_m.sort_values(COL_START, ascending=False)
                if COL_MENTOR_DT in df_m.columns:
                    df_m = df_m.sort_values([COL_START, COL_MENTOR_DT], ascending=[False, False])
                df_m = df_m.drop_duplicates(subset=[COL_RECORD_ID], keep="first").reset_index(drop=True)

            # Detect program column
            prog_col_candidates = [c for c in df_m.columns if 'Program Enrolling' in str(c)]
            COL_PROG = prog_col_candidates[0] if prog_col_candidates else None

            if COL_PROG is None:
                raise ValueError("Program Enrolling column not found in Mentorship.xlsx")

            # Get or create Mentorship tab
            try:
                mentor_sheet = spreadsheet.worksheet("Mentorship")
                mentor_df = pd.DataFrame(mentor_sheet.get_all_records())
                if "Date" in mentor_df.columns:
                    mentor_df["Date"] = pd.to_datetime(mentor_df["Date"], errors="coerce", format="mixed")
                    mentor_df = mentor_df.set_index("Date")
                else:
                    mentor_df = pd.DataFrame()
            except gspread.exceptions.WorksheetNotFound:
                mentor_df = pd.DataFrame()

            # Get program columns from tab header; fallback to unique programs in data
            # Only include Renew + TP programs
            if not mentor_df.empty:
                mentor_programs = [c for c in mentor_df.columns if c != "Total" and c in MENTOR_PROGRAMS]
            else:
                mentor_programs = sorted(MENTOR_PROGRAMS)

            # Filter to Renew + TP programs only
            df_m = df_m[df_m[COL_PROG].astype(str).str.strip().isin(MENTOR_PROGRAMS)]

            # Count clients per program
            mentor_counts = df_m[COL_PROG].astype(str).str.strip().value_counts()

            # Build today's row
            new_row = {prog: int(mentor_counts.get(prog, 0)) for prog in mentor_programs}
            new_row["Total"] = int(sum(new_row.values()))

            # Overwrite if same date exists, otherwise append
            today_str = today_date
            if mentor_df.empty:
                mentor_df = pd.DataFrame([new_row], index=pd.Index([today_str], name="Date"))
            else:
                # Convert index to consistent string format
                mentor_df.index = pd.to_datetime(mentor_df.index, errors="coerce", format="mixed")
                mentor_df = mentor_df[mentor_df.index.notna()]
                mentor_df.index = mentor_df.index.strftime('%-m/%-d/%y')
                mentor_df.index.name = "Date"
                mentor_df.loc[today_str] = new_row

            write_tab(spreadsheet, "Mentorship", mentor_df.reset_index())
            print(f"✅ Mentorship tab updated: {new_row['Total']} clients")

        except Exception as e:
            print(f"⚠️  Error processing Mentorship tab: {e}")

    print(f"Successfully updated Google Sheet: {input_utilization_file}")
    print(f"Link: https://docs.google.com/spreadsheets/d/{utilization_file_id}")

    # Run averages in same folder
    run_averages_processing(folder_id=output_folder_id)


def run_averages_processing(
    source_file = 'Clients in All Programs - Processed',
    output_file = 'Clients in All Programs Average',
    source_folder_name = None,
    folder_id = None
):
    drive_service, gc, _ = get_services()
    if folder_id is None:
        folder_id = resolve_folder_id(drive_service, source_folder_name, "Source")
    source_folder_id = folder_id
    output_folder_id = folder_id  # Save in same folder as source

    # Open source sheet
    src_file = find_file_id(drive_service, source_file, source_folder_id, "application/vnd.google-apps.spreadsheet")
    if not src_file:
        raise FileNotFoundError(f"Source sheet not found: {source_file}")
    src_ss = gc.open_by_key(src_file['id'])

    # Open or create output sheet
    out_file = find_file_id(drive_service, output_file, output_folder_id, "application/vnd.google-apps.spreadsheet")
    if out_file:
        out_ss = gc.open_by_key(out_file['id'])
    else:
        file_metadata = {'name': output_file, 'mimeType': 'application/vnd.google-apps.spreadsheet'}
        if output_folder_id:
            file_metadata['parents'] = [output_folder_id]
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        out_ss = gc.open_by_key(new_sheet['id'])

    # Helper functions
    def get_fy(date):
        return date.year + 1 if date.month >= 9 else date.year

    def get_fy_label(date):
        fy = get_fy(date)
        return f"FY{str(fy)[-2:]}"

    def get_week_num(date):
        fy = get_fy(date)
        fy_start = pd.Timestamp(f"{fy-1}-09-01")  # Always Sep 1
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

    def read_tab(tab_name):
        for attempt in range(3):
            try:
                time.sleep(5)
                ws = src_ss.worksheet(tab_name)
                df = pd.DataFrame(ws.get_all_records())
                if df.empty or "Date" not in df.columns:
                    return None
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce", format="mixed")
                df = df.dropna(subset=["Date"])
                prog_cols = [c for c in df.columns if c != "Date" and not str(c).endswith("Number of Children") and str(c) != "Total"]
                # For Mentorship tab, filter to Renew + TP programs only
                if tab_name == "Mentorship":
                    prog_cols = [c for c in prog_cols if c in MENTOR_PROGRAMS]
                return df, prog_cols
            except Exception as e:
                if "429" in str(e) and attempt < 2:
                    print(f"   Rate limited on '{tab_name}', retrying in 30s...")
                    time.sleep(30)
                else:
                    print(f"⚠️  Could not read tab '{tab_name}': {e}")
                    return None
        return None

    def compute_averages(df, prog_cols, period_fn, fy_period_fn=None):
        rows = []
        df = df.copy()
        df["_period"] = df["Date"].apply(period_fn)
        df["_fy_period"] = df["Date"].apply(fy_period_fn) if fy_period_fn else df["Date"].apply(period_fn)
        df["_fy"] = df["Date"].apply(get_fy_label)
        for period, grp in df.groupby("_period", sort=False):
            fy = grp["_fy"].iloc[0]
            fy_period = grp["_fy_period"].iloc[0]
            for prog in prog_cols:
                if prog not in grp.columns: continue
                vals = pd.to_numeric(grp[prog], errors="coerce").dropna()
                avg = round(vals.mean(), 2) if not vals.empty else 0
                rows.append({"FY": fy, "Period": period, "FY Period": fy_period, "Category": PROGRAM_CATEGORY.get(prog, "Other"), "Program": prog, "Average": avg})
        return pd.DataFrame(rows)

    def write_averages(tab_name, df_avg, period_col_name):
        df_out = df_avg.rename(columns={
            "Period": period_col_name,
            "FY Period": f"FY {period_col_name}"
        })
        # Ensure column order
        cols = ["FY", period_col_name, f"FY {period_col_name}", "Category", "Program", "Average"]
        df_out = df_out[[c for c in cols if c in df_out.columns]]
        try:
            ws = out_ss.worksheet(tab_name)
            existing = pd.DataFrame(ws.get_all_records())
            if not existing.empty and period_col_name in existing.columns:
                merge_key = ["FY", period_col_name, "Program"]
                existing = existing[~existing.set_index(merge_key).index.isin(df_out.set_index(merge_key).index)]
                df_out = pd.concat([existing, df_out], ignore_index=True)
        except gspread.exceptions.WorksheetNotFound:
            pass
        write_tab(out_ss, tab_name, df_out)
        print(f"✅ {tab_name} written")

    # Process occupancy tabs
    occ_tabs = ["Renew", "Men Turning Point", "Women Turning Point", "Forward", "Program Interns", "Other", "APH", "Elevate"]
    all_occ_dfs = []
    occ_prog_cols = []
    for tab_name in occ_tabs:
        result = read_tab(tab_name)
        if result:
            df_tab, prog_cols = result
            all_occ_dfs.append(df_tab)
            for p in prog_cols:
                if p not in occ_prog_cols:
                    occ_prog_cols.append(p)

    if all_occ_dfs:
        df_occ = pd.concat(all_occ_dfs, ignore_index=True).groupby("Date", as_index=False).sum(numeric_only=True)
        write_averages("Occupancy Weekly Avg",    compute_averages(df_occ, occ_prog_cols, get_week_label, get_fy_week_label), "Week")
        write_averages("Occupancy Monthly Avg",   compute_averages(df_occ, occ_prog_cols, get_month, get_fy_month), "Month")
        write_averages("Occupancy Quarterly Avg", compute_averages(df_occ, occ_prog_cols, get_quarter, get_fy_quarter), "Quarter")

    # Process Mentorship tab
    result = read_tab("Mentorship")
    if result:
        df_m, mentor_progs = result
        write_averages("Mentorship Weekly Avg",    compute_averages(df_m, mentor_progs, get_week_label, get_fy_week_label), "Week")
        write_averages("Mentorship Monthly Avg",   compute_averages(df_m, mentor_progs, get_month, get_fy_month), "Month")
        write_averages("Mentorship Quarterly Avg", compute_averages(df_m, mentor_progs, get_quarter, get_fy_quarter), "Quarter")

    # =========================================================================
    # Fetch FY27 Goals & Extra Cells from Constants Update Sheet
    # =========================================================================
    PROGRAM_GOAL_CELLS = {
        "Chester Renew": "D7",
        "GV Renew": "D11",
        "Oakland Renew": "D15",
        "Portland Renew": "D19",
        "San Jose Men Renew": "D23",
        "Chester Men Turning Point": "I11",
        "Oakland Men Turning Point": "I15",
        "San Jose Men Turning Point": "I39",
        "Portland Men Turning Point": "I27",
        "Chester Women Turning Point": "I11",
        "GV Turning Point": "I31",
        "Heritage Home": "I35",
        "Oakland Women Turning Point": "I19",
        "Portland Community of Hope": "I23",
        "San Jose Youth Collective": "I43",
    }

    program_goals = {}
    gv_tp_extra_count = 0  # To store cell P30 value from Sheet1

    try:
        _, _, _creds = get_services()
        _sheets = _build("sheets", "v4", credentials=_creds)
        _folder = resolve_folder_id(drive_service, "CityTeam KPIs", "Constants")
        _cf = find_file_id(drive_service, "Constants Update", _folder)

        if _cf:
            # Batch fetch goal ranges from FY27 Goals + cell P30 from Sheet1
            ranges = [f"'FY27 Goals'!{cell}" for cell in PROGRAM_GOAL_CELLS.values()]
            ranges.append("'Sheet1'!P30")

            batch_resp = _sheets.spreadsheets().values().batchGet(
                spreadsheetId=_cf["id"], ranges=ranges
            ).execute()

            # Map cell values to programs and extract P30
            value_ranges = batch_resp.get("valueRanges", [])
            cell_val_map = {}
            for vr in value_ranges:
                full_range = vr.get("range", "")
                cell_name = full_range.split("!")[-1].replace("$", "")
                vals = vr.get("values", [[""]])
                val_str = str(vals[0][0]).replace(",", "").strip() if vals and vals[0] else None
                
                try:
                    num_val = int(float(val_str)) if val_str not in [None, ""] else 0
                except ValueError:
                    num_val = 0

                if "Sheet1" in full_range and cell_name == "P30":
                    gv_tp_extra_count = num_val
                else:
                    cell_val_map[cell_name] = num_val

            # Assign retrieved values back to program key
            for prog, cell_ref in PROGRAM_GOAL_CELLS.items():
                program_goals[prog] = cell_val_map.get(cell_ref, 0)

            print(f"✅ Loaded FY27 goals and Sheet1!P30 extra count ({gv_tp_extra_count}) from Constants Update")

    except Exception as e:
        print(f"⚠️  Could not load Constants Update data: {e}")

    # =========================================================================
    # Latest Occupancy tab — most recent row from each occupancy tab
    # =========================================================================
    latest_occ_rows = []
    for tab_name in occ_tabs:
        result = read_tab(tab_name)
        if result:
            df_tab, prog_cols = result
            if df_tab.empty: continue
            latest_date = df_tab["Date"].max()
            latest_row = df_tab[df_tab["Date"] == latest_date].iloc[0]
            for prog in prog_cols:
                if prog in latest_row:
                    latest_occ_rows.append({
                        "Date": latest_date.strftime("%-m/%-d/%y"), 
                        "Program": prog, 
                        "Count": pd.to_numeric(latest_row[prog], errors="coerce")
                    })

    if latest_occ_rows:
        df_latest_occ = pd.DataFrame(latest_occ_rows)
        df_latest_occ["Count"] = df_latest_occ["Count"].fillna(0)

        # Add Sheet1!P30 extra count to GV Turning Point
        gv_mask = df_latest_occ["Program"] == "GV Turning Point"
        if gv_mask.any():
            df_latest_occ.loc[gv_mask, "Count"] += gv_tp_extra_count

        df_latest_occ["Category"] = df_latest_occ["Program"].map(PROGRAM_CATEGORY).fillna("Other")
        
        # Map Next FY Goal using program name lookup
        df_latest_occ["Next FY Goal"] = df_latest_occ["Program"].map(program_goals).fillna(0).astype(int)

        # Include Next FY Goal in column ordering
        df_latest_occ = df_latest_occ[["Date", "Category", "Program", "Count", "Next FY Goal"]]

        write_tab(out_ss, "Latest Occupancy", df_latest_occ)
        print("✅ Latest Occupancy tab written with updated GV Turning Point Count and Next FY Goal column")

    # =========================================================================
    # Latest Mentorship tab — most recent row from Mentorship tab
    # =========================================================================
    result = read_tab("Mentorship")
    if result:
        df_m2, mentor_progs2 = result
        if not df_m2.empty:
            latest_date = df_m2["Date"].max()
            latest_row = df_m2[df_m2["Date"] == latest_date].iloc[0]
            latest_m_rows = []
            for prog in mentor_progs2:
                if prog in latest_row:
                    latest_m_rows.append({"Date": latest_date.strftime("%-m/%-d/%y"), "Program": prog, "Count": pd.to_numeric(latest_row[prog], errors="coerce")})
            if latest_m_rows:
                df_latest_m = pd.DataFrame(latest_m_rows)
                df_latest_m["Category"] = df_latest_m["Program"].map(PROGRAM_CATEGORY).fillna("Other")
                df_latest_m = df_latest_m[["Date", "Category", "Program", "Count"]]
                write_tab(out_ss, "Latest Mentorship", df_latest_m)
                print("✅ Latest Mentorship tab written")

    # Cleanup Sheet1
    try:
        out_ss.del_worksheet(out_ss.worksheet("Sheet1"))
    except Exception:
        pass

    print(f"✅ Averages done: {out_ss.url}")


if __name__ == "__main__":
    print("🚀 Starting Program Utilization Processing"),
    run_utilization_processing()