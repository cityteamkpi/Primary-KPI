# ==========================================
#
# This script processes the Apricot NR
# report output - "All Programs All Alumni KPIs.xlsx"
# 
# ==========================================


# --- Third-Party Imports ---
import warnings
import gspread
from gspread_dataframe import set_with_dataframe
import pandas as pd
from auth_utils import get_services
from drive_utils import resolve_folder_id, download_drive_file, write_tab

# Suppress openpyxl warnings regarding styles and validation metadata
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ==========================================
# STEP 0: DEFINE  RULES
# ==========================================

# We define the Living Wage 'Target' for each program as a Dictionary
# Need to get from Google Sheet eventually
living_wage_targets = {
    'Chester Renew': 17.50,
    'Portland Renew': 20.00,
    'Oakland Renew': 22.00,
    'GV Renew': 24.00,
    'San Jose Men Renew': 24.00
}


def run_alum_processing(start_date, end_date, input_folder_name, output_folder_name):

   # Hardcoded configurations
    input_filename = "All Programs All Alumni KPIs.xlsx"
    output_file_name = None  # Will default to dynamic name below
    
    
    # ==========================================
    # STEP 1:
    # Initialize Google Drive API Service using auth_utils 
    # (Google Drive access authorization) and find specified folders
    # ==========================================
    drive_service, client, creds = get_services()

    # Resolve folder IDs
    input_folder_id = resolve_folder_id(drive_service, input_folder_name, "Input")
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # Default output title if none provided
    output_file_name = output_file_name or f"Alumni Analysis {start_date} to {end_date}"

    # ==========================================
    # STEP 2: GET FILE FROM GOOGLE DRIVE (Using utility)
    # ==========================================
    fh, target_folder_id, _ = download_drive_file(drive_service, input_filename, input_folder_id)
    print(f"Successfully downloaded {input_filename} from folder: {target_folder_id}")

    # Finalize output folder: use explicit output folder, or fallback to the file's current home
    final_output_folder_id = output_folder_id or target_folder_id

    # ==========================================
    # STEP 3: LOAD AND CLEAN THE DATA
    # ==========================================
    df = pd.read_excel(fh)
    df['Graduation Date'] = pd.to_datetime(df['Graduation Date'])
    df['Checkin Date'] = pd.to_datetime(df['Checkin Date'])

    def clean_wage(wage_str):
        if pd.isna(wage_str): return 0.0
        if isinstance(wage_str, (int, float)): return float(wage_str)
        clean_str = str(wage_str).replace('$', '').replace(',', '').strip()
        try: return float(clean_str)
        except ValueError: return 0.0

    df['Hourly Wage Cleaned'] = df['Hourly Wage'].apply(clean_wage)
    df['Housing Expense'] = df['Housing Expense'].apply(clean_wage)

    MONTHLY_HOURS = 173.33
    df['Monthly Gross Income'] = df['Hourly Wage Cleaned'] * MONTHLY_HOURS
    # Ensure column is string to avoid errors with .str accessor on float/NaN values
    df['Sober?'] = df['Sober?'].astype(str).str.strip()

    print(f"Total Alumni processed: {len(df)}")

    # Convert string dates to datetime objects to avoid invalid comparison errors
    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)

    cohort_filter = (df['Graduation Date'] >= start_dt) & (df['Graduation Date'] <= end_dt)
    cohort_list = df[cohort_filter][['Name', 'Program', 'Graduation Date']].drop_duplicates()
  

    # ==========================================
    # STEP 4: ANALYZE HOUSING AND SOBRIETY
    # ==========================================
    results = []
    for _, row in cohort_list.iterrows():
        name = row['Name']
        program = row['Program']
        grad_date = row['Graduation Date']
        target = living_wage_targets.get(program, 0.0)
        
        person_records = df[(df['Name'] == name) & (df['Program'] == program)].sort_values(by='Checkin Date')
        
        sober_records = person_records[person_records['Sober?'].isin(['Yes', 'No'])]
        has_sustained_relapse = False
        last_state = None
        first_no_date = None
        last_checkin_sober = False

        for _, s_row in sober_records.iterrows():
            current_state = s_row['Sober?']
            current_date = s_row['Checkin Date']
            if last_state == 'Yes' and current_state == 'No':
                if first_no_date is None: first_no_date = current_date
            elif last_state == 'No' and current_state == 'Yes':
                if first_no_date is not None:
                    days_elapsed = (current_date - first_no_date).days
                    if days_elapsed > 30: has_sustained_relapse = True
                first_no_date = None
            last_state = current_state
            last_checkin_sober = (current_state == 'Yes')
        is_sober = last_checkin_sober and not has_sustained_relapse

        # Ensure column is string to avoid "Can only use .str accessor with string values" error
        grad_record = person_records[person_records['How Checked'].astype(str).str.contains('Graduation', na=False)]
        grad_wage = grad_record['Hourly Wage Cleaned'].iloc[0] if not grad_record.empty else 0.0
        recent_wage = person_records.iloc[-1]['Hourly Wage Cleaned']
        
        if not grad_record.empty:
            grad_income = grad_record['Monthly Gross Income'].iloc[0]
            grad_housing = grad_record['Housing Expense'].iloc[0]
            if grad_housing == 0: grad_under_30 = 100.0
            elif grad_income == 0: grad_under_30 = 0.0
            else: grad_under_30 = 100.0 if (grad_housing / grad_income) < 0.30 else 0.0
        else: grad_under_30 = None

        recent_record = person_records.iloc[-1]
        recent_checkin = recent_record['Checkin Date']
        recent_income = recent_record['Monthly Gross Income']
        recent_housing = recent_record['Housing Expense']
        
        if recent_housing == 0: recent_under_30 = 100.0
        elif recent_income == 0: recent_under_30 = 0.0
        else: recent_under_30 = 100.0 if (recent_housing / recent_income) < 0.30 else 0.0
            
        results.append({
            'Name': name, 'Program': program, 'Graduation Date': grad_date,
            'Most Recent Checkin': recent_checkin, 'is_sober': is_sober,
            'Wage (Grad)': grad_wage, 'At Living Wage (Grad)': 100.0 if grad_wage >= target else 0.0,
            'Wage (Recent)': recent_wage, 'At Living Wage (Recent)': 100.0 if recent_wage >= target else 0.0,
            'Housing Under 30% (Grad)': grad_under_30, 'Housing Under 30% (Recent)': recent_under_30,
            'Final Status': "Sober" if is_sober else "Not Sober", 'Sustained Relapse?': "Yes" if has_sustained_relapse else "No"
        })

    results_df = pd.DataFrame(results)
    all_stats = results_df.groupby('Program').agg(
        Total_Grads=('Name', 'count'), LW_at_Grad=('At Living Wage (Grad)', 'mean'),
        Housing_Under30_at_Grad=('Housing Under 30% (Grad)', 'mean')
    ).reset_index()

    sober_results_df = results_df[results_df['is_sober'] == True]
    sober_stats = sober_results_df.groupby('Program').agg(
        Sober_Grads=('Name', 'count'), LW_at_Recent_Check=('At Living Wage (Recent)', 'mean'),
        Housing_Under30_at_Recent_Check=('Housing Under 30% (Recent)', 'mean')
    ).reset_index()

    analysis = pd.merge(all_stats, sober_stats, on='Program', how='left').fillna(0)

    # Ensure all columns representing percentages are rounded to 2 digits of precision
    pct_cols = ['LW_at_Grad', 'LW_at_Recent_Check', 'Housing_Under30_at_Grad', 'Housing_Under30_at_Recent_Check']
    analysis[pct_cols] = analysis[pct_cols].round(2)

    # Update display columns to use the rounded source values
    analysis['% LW (Graduation)'] = analysis['LW_at_Grad']
    analysis['% LW (Recent)'] = analysis['LW_at_Recent_Check']
    analysis['% Housing (Grad)'] = analysis['Housing_Under30_at_Grad']
    analysis['% Housing (Recent)'] = analysis['Housing_Under30_at_Recent_Check']
    analysis['% Sober'] = (analysis['Sober_Grads'] / analysis['Total_Grads'] * 100).round(2)

    detailed_results_df = results_df.sort_values(by=['Program','Name'])

    # ==========================================
    # STEP 5: EXPORT TO GOOGLE SHEETS
    # ==========================================
    
    search_q = f"name = '{output_file_name}' and '{final_output_folder_id}' in parents and trashed = false"
    search_res = drive_service.files().list(q=search_q, supportsAllDrives=True, includeItemsFromAllDrives=True, fields="files(id)").execute()
    existing_summary = search_res.get('files', [])

    if existing_summary:
        spreadsheet = client.open_by_key(existing_summary[0]['id'])
    else:
        file_metadata = {'name': output_file_name, 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [final_output_folder_id]}
        new_file = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        spreadsheet = client.open_by_key(new_file['id'])
    
    for tab_name, data in [("Quick Stats", analysis), ("Detailed Analysis", detailed_results_df)]:
        write_tab(spreadsheet, tab_name, data)

    for program in detailed_results_df['Program'].unique():
        prog_sheet_title = f"{program} Analysis ({start_date} to {end_date})"
        prog_search_q = f"name = '{prog_sheet_title}' and '{final_output_folder_id}' in parents and trashed = false"
        prog_search_res = drive_service.files().list(q=prog_search_q, supportsAllDrives=True, includeItemsFromAllDrives=True, fields="files(id)").execute()
        existing_prog = prog_search_res.get('files', [])

        if existing_prog: prog_spreadsheet = client.open_by_key(existing_prog[0]['id'])
        else:
            file_metadata = {'name': prog_sheet_title, 'mimeType': 'application/vnd.google-apps.spreadsheet', 'parents': [final_output_folder_id]}
            new_file = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
            prog_spreadsheet = client.open_by_key(new_file['id'])

        for tab, data in [("Stats", analysis[analysis['Program'] == program]), ("Detail", detailed_results_df[detailed_results_df['Program'] == program])]:
            write_tab(prog_spreadsheet, tab, data)

    return f"Successfully processed {input_filename}."
