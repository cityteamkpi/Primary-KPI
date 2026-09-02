# =========================================================================
# This script processes program utilization.
#
# 1. Reads "Clients in All Programs.xlsx.xlsx"
# 2. Creates "Clients in All Programs - Processed" Google Sheet 
#       - with Renew, Men's TP and Women's TP tabs
# =========================================================================


import warnings
import pandas as pd
import gspread
from zoneinfo import ZoneInfo
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime
from auth_utils import get_services
from drive_utils import download_drive_file, resolve_folder_id, find_file_id, load_raw, write_tab


# Suppress openpyxl warnings regarding styles and validation metadata
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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
                cols_list = list(df_sheet.columns)

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
                    mentor_df["Date"] = pd.to_datetime(mentor_df["Date"], errors="coerce")
                    mentor_df = mentor_df.set_index("Date")
                else:
                    mentor_df = pd.DataFrame()
            except gspread.exceptions.WorksheetNotFound:
                mentor_df = pd.DataFrame()

            # Get program columns from tab header; fallback to unique programs in data
            if not mentor_df.empty:
                mentor_programs = [c for c in mentor_df.columns if c != "Total"]
            else:
                mentor_programs = sorted(df_m[COL_PROG].dropna().astype(str).str.strip().unique().tolist())

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
                mentor_df.index = mentor_df.index.strftime('%-m/%-d/%y') if hasattr(mentor_df.index, 'strftime') else mentor_df.index
                mentor_df.loc[today_str] = new_row

            write_tab(spreadsheet, "Mentorship", mentor_df.reset_index())
            print(f"✅ Mentorship tab updated: {new_row['Total']} clients")

        except Exception as e:
            print(f"⚠️  Error processing Mentorship tab: {e}")

    print(f"Successfully updated Google Sheet: {input_utilization_file}")
    print(f"Link: https://docs.google.com/spreadsheets/d/{utilization_file_id}")

    # Run averages in same folder
    run_averages_processing(folder_id=output_folder_id)

if __name__ == "__main__":
    run_utilization_processing()


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

    def get_week_label(date):
        fy = get_fy(date)
        fy_label = get_fy_label(date)
        fy_start = pd.Timestamp(f"{fy-1}-09-01")
        sep1_week_monday = fy_start - pd.Timedelta(days=fy_start.weekday())
        delta = (date - sep1_week_monday).days
        w = max(1, delta // 7 + 1)
        return f"{fy_label} Week {w}"

    def get_month(date):
        return date.strftime("%B %Y")

    def get_quarter(date):
        fy_label = get_fy_label(date)
        m = date.month
        if m in [9, 10, 11]: q = "Q1"
        elif m in [12, 1, 2]: q = "Q2"
        elif m in [3, 4, 5]: q = "Q3"
        else: q = "Q4"
        return f"{fy_label} {q}"

    def read_tab(tab_name):
        try:
            ws = src_ss.worksheet(tab_name)
            df = pd.DataFrame(ws.get_all_records())
            if df.empty or "Date" not in df.columns:
                return None
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"])
            prog_cols = [c for c in df.columns if c != "Date" and not str(c).endswith("Number of Children") and str(c) != "Total"]
            return df, prog_cols
        except Exception as e:
            print(f"⚠️  Could not read tab '{tab_name}': {e}")
            return None

    def compute_averages(df, prog_cols, period_fn):
        rows = []
        df = df.copy()
        df["_period"] = df["Date"].apply(period_fn)
        df["_fy"] = df["Date"].apply(get_fy_label)
        for period, grp in df.groupby("_period", sort=False):
            fy = grp["_fy"].iloc[0]
            for prog in prog_cols:
                if prog not in grp.columns: continue
                vals = pd.to_numeric(grp[prog], errors="coerce").dropna()
                avg = round(vals.mean(), 2) if not vals.empty else 0
                rows.append({"FY": fy, "Period": period, "Program": prog, "Average": avg})
        return pd.DataFrame(rows)

    def write_averages(tab_name, df_avg, period_col_name):
        df_out = df_avg.rename(columns={"Period": period_col_name})
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
        write_averages("Occupancy Weekly Avg",    compute_averages(df_occ, occ_prog_cols, get_week_label), "Week")
        write_averages("Occupancy Monthly Avg",   compute_averages(df_occ, occ_prog_cols, get_month), "Month")
        write_averages("Occupancy Quarterly Avg", compute_averages(df_occ, occ_prog_cols, get_quarter), "Quarter")

    # Process Mentorship tab
    result = read_tab("Mentorship")
    if result:
        df_m, mentor_progs = result
        write_averages("Mentorship Weekly Avg",    compute_averages(df_m, mentor_progs, get_week_label), "Week")
        write_averages("Mentorship Monthly Avg",   compute_averages(df_m, mentor_progs, get_month), "Month")
        write_averages("Mentorship Quarterly Avg", compute_averages(df_m, mentor_progs, get_quarter), "Quarter")

    # Cleanup Sheet1
    try:
        out_ss.del_worksheet(out_ss.worksheet("Sheet1"))
    except:
        pass

    print(f"✅ Averages done: {out_ss.url}")