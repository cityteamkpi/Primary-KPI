import io
import socket
import sys
import time
import warnings
from auth_utils import get_services
from concurrent.futures import ThreadPoolExecutor
from drive_utils import download_drive_file, find_file_id, resolve_folder_id
import gspread
import numpy as np
import pandas as pd

socket.setdefaulttimeout(120)
warnings.filterwarnings("ignore", category=UserWarning)

# ============================================================
#  GLOBAL CONSTANTS & MAPPINGS
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
    "Electrical Trainee Card Date_5727",
    "NCCER Date_5728",
    "Blueprint Reading Date_6785",
    "Forklift Operator Date_5106",
    "OSHA10 Date_5118",
    "MA Date_5115",
    "Google IT Support Certificate Date_5107",
    "HAZWOPER Date_5109",
    "HAZMAT Date_5108",
    "CPR/FA/AED Date_5105",
    "CADC Date_5789",
    "ServSafe Food Manager Date_5126",
    "ServSafe Food Handler Date_5125",
    "Material Handling Date_5116",
    "Multi-Core Craft Curriculum Completion Date_5117",
    "RADT Date_5121",
    "HVAC Date_5113",
    "CDL Date_6104",
    "CompTIA and/or A+ Date_5104",
    "OTHER Date_5119",
    "Guard Card Date_6561",
    "TWIC Date_6631",
    "Google Automate Cybersecurity w/Python Date_6273",
    "CNC Machining Date_5103",
    "Confined Space Date_6788",
    "DVAT Date_6791",
    "Fundamentals of Construction Date_6790",
    "HR Associate Date_6789",
    "Intuit Bookkeeping Date_6105",
    "Massage Therapy Cert Date_7032",
    "Peer Support Date_7031",
    "Record Expungement Date_6787",
    "Record Expungement Attorney Consultation Date_6879",
    "Traffic Control Date_6786",
    "Certificate/License Date",
]

DESTINATION_SHEET_ID = "1bvcXczlUgCpdY5jPJopjJ6lT5J03g1qrGo9zQgc93Eg"

FY_CONFIGS = {
    "FY26": {
        "start": pd.Timestamp("2025-09-01"),
        "end": pd.Timestamp("2026-08-31"),
        "sheet_name": "Prior Fiscal Year",
    },
    "FY27": {
        "start": pd.Timestamp("2026-09-01"),
        "end": pd.Timestamp.now().normalize(),
        "sheet_name": "Current Fiscal Year To-Date",
    },
}

FILE_OCCUPANCY_NAME = "Clients in All Programs Average"
FILE_RENEW_NAME = "Renew Programs Report.xlsx"
FILE_TP_NAME = "Turning Point Report.xlsx"
FILE_BAPTISMS_NAME = "CityTeam Baptisms.xlsx"
FILE_MENTOR_NAME = "Mentorship.xlsx"
FILE_REUNIF_NAME = "Family Reunifications.xlsx"
FILE_ALUMNI_NAME = "All Programs All Alumni KPIs.xlsx"
FILE_ATWORK_NAME = "All Programs @Work Fiscal Year Impacts.xlsx"
FILE_CLIENTS_ALL_NAME = "Clients in All Programs.xlsx"

COLUMN_INDICES = {
    "San Jose Men": 2,
    "San Jose Women": 3,
    "Oakland Men": 4,
    "Oakland Women": 5,
    "Portland Men": 6,
    "Portland Women": 7,
    "Chester Men": 8,
    "Chester Women": 9,
}

PROGRAM_MAP = {
    "San Jose Men": {
        "renew": ["San Jose Men Renew"],
        "tp": ["San Jose Men Turning Point"],
        "fwd_intern": ["GV Men's Forward"],
        "connect": ["San Jose Men Connect"],
    },
    "San Jose Women": {
        "renew": ["GV Renew", "House of Grace"],
        "tp": [
            "GV Turning Point",
            "Heritage Home",
            "San Jose Youth Collective",
        ],
        "fwd_intern": [
            "GV Women's Forward",
            "San Jose Women's House of Light",
        ],
        "connect": ["San Jose Connect"],
    },
    "Oakland Men": {
        "renew": ["Oakland Renew"],
        "tp": [
            "Oakland Men Turning Point",
            "Oakland Youth Collective",
        ],
        "fwd_intern": ["Oakland Forward Men", "Oakland Forwad Men"],
        "connect": [],
    },
    "Oakland Women": {
        "renew": [],
        "tp": ["Oakland Women Turning Point"],
        "fwd_intern": ["Oakland Forward Women"],
        "connect": [],
    },
    "Portland Men": {
        "renew": ["Portland Renew"],
        "tp": [
            "Portland Men Turning Point",
            "Portland Youth Collective",
        ],
        "fwd_intern": ["Portland Forward"],
        "connect": [],
    },
    "Portland Women": {
        "renew": [],
        "tp": ["Portland Community of Hope"],
        "fwd_intern": [],
        "connect": [],
    },
    "Chester Men": {
        "renew": ["Chester Renew"],
        "tp": [
            "Chester Men Turning Point",
            "Chester Turning Point",
        ],
        "fwd_intern": ["Chester Forward"],
        "connect": [],
    },
    "Chester Women": {
        "renew": [],
        "tp": [
            "Chester Women Turning Point",
            "Chester Women's Turning Point",
        ],
        "fwd_intern": [],
        "connect": [],
    },
}

# Complete all-program mapping by City & Gender for Spiritual (Rows 25-27) & @Work (Rows 31-36)
ALL_CITY_PROGRAMS_MAP = {
    "San Jose Men": {
        "san jose men renew",
        "san jose men turning point",
        "gv men's forward",
        "san jose men connect",
    },
    "San Jose Women": {
        "gv renew",
        "house of grace",
        "gv turning point",
        "heritage home",
        "san jose youth collective",
        "gv women's forward",
        "san jose women's house of light",
        "san jose connect",
        "grace village connect",
        "gv aph",
        "gv immediate temporary housing",
        "gv transition phase",
        "sj wp transition phase",
    },
    "Oakland Men": {
        "oakland renew",
        "oakland men turning point",
        "oakland youth collective",
        "oakland forward men",
        "oakland forwad men",
        "oakland connect",
    },
    "Oakland Women": {
        "oakland women turning point",
        "oakland forward women",
    },
    "Portland Men": {
        "portland renew",
        "portland men turning point",
        "portland forward",
        "portland youth collective",
        "portland connect",
    },
    "Portland Women": {
        "portland community of hope",
    },
    "Chester Men": {
        "chester renew",
        "chester men turning point",
        "chester turning point",
        "chester forward",
    },
    "Chester Women": {
        "chester women turning point",
        "chester women's turning point",
    },
}

ROW_RES_NEW_ADULT_ENROLLMENTS = 5
ROW_RES_NEW_CHILD_ENROLLMENTS = 6

ROW_RENEW_ADULT_AVG_OCCUPANCY = 7
ROW_TP_ADULT_AVG_OCCUPANCY = 9
ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY = 11

ROW_RENEW_ADULT_NIGHTS = 14
ROW_TP_ADULT_NIGHTS = 15
ROW_FWD_INTERN_NIGHTS = 16

ROW_RENEW_GRADUATIONS = 21
ROW_TP_GRADUATIONS = 22
ROW_TP_SUCCESSFULLY_HOUSED = 23

ROW_FAMILIES_REUNIFIED = 25
ROW_SPIRITUAL_MENTORSHIPS = 26
ROW_BAPTISMS = 27
ROW_SOBRIETY_RATE = 28
ROW_CASE_MANAGEMENT_HOURS = 29

ROW_JOBS_ATTAINED = 31
ROW_CREDENTIALS_ATTAINED = 32
ROW_IDS_CARDS_DOCS = 33
ROW_DRIVERS_LICENSES = 34
ROW_HSE_ATTAINED = 35
ROW_BANK_ACCOUNTS = 36
ROW_HOUSING_BURDEN_1YR = 37
ROW_WAGE_INCREASE = 38

ROW_JUSTICE_INVOLVED = 43
ROW_HOMELESSNESS = 44
ROW_SUBSTANCE_ABUSE = 45
ROW_DOMESTIC_VIOLENCE = 46

ROWS_KPI_SPEC = [
    (
        "Row 5",
        "Res. New Adult Enrollments",
        "Renew Programs Report / Turning Point Report",
        (
            "Count of unique [Record ID, Program Enrolling] starting Renew/TP"
            " in FY"
        ),
    ),
    (
        "Row 6",
        "Res. New Child Enrollments",
        "Renew Programs Report / Turning Point Report",
        "Sum of 'Number of Children' for Renew/TP enrollments in FY",
    ),
    (
        "Row 7",
        "Renew Adult Avg Occupancy",
        "Clients in All Programs Average ('Latest Occupancy')",
        "Sum of 'Count' for Renew category (FY27)",
    ),
    (
        "Row 9",
        "Turning Point Adult Avg Occupancy",
        "Clients in All Programs Average ('Latest Occupancy')",
        "Sum of 'Count' for TP Category (FY27)",
    ),
    (
        "Row 11",
        "Forward - Intern Adult Avg Occ",
        "Clients in All Programs Average ('Latest Occupancy')",
        "Sum of 'Count' for Forward & Intern Categories (FY27)",
    ),
    (
        "Row 14",
        "Renew Adult Nights of Shelter",
        "Formula Calculation",
        "Row 7 * 365 (FY26) | Row 7 * 7 * weeks_to_date (FY27)",
    ),
    (
        "Row 15",
        "Turning Point Adult Nights",
        "Formula Calculation",
        "Row 9 * 365 (FY26) | Row 9 * 7 * weeks_to_date (FY27)",
    ),
    (
        "Row 16",
        "Forward - Intern Nights",
        "Formula Calculation",
        "Row 11 * 365 (FY26) | Row 11 * 7 * weeks_to_date (FY27)",
    ),
    (
        "Row 21",
        "Renew Graduations",
        "Renew Programs Report.xlsx",
        "Count of unique graduates with primary reason in Graduation list",
    ),
    (
        "Row 22",
        "Turning Point Graduations",
        "Turning Point Report.xlsx",
        "Count of unique graduates with reason 'Graduation'",
    ),
    (
        "Row 23",
        "Turning Point Housed",
        "Turning Point Report.xlsx",
        (
            "Count of unique clients (incl. Interns) with Date of Housing Check"
            " (fallback Exit Date) in FY & (Housed=='Yes' or Date of Housing"
            " Check not null)"
        ),
    ),
    (
        "Row 25",
        "Families Reunified",
        "Family Reunifications.xlsx",
        (
            "Count of unique client reunifications occurring in FY (all city"
            " programs)"
        ),
    ),
    (
        "Row 26",
        "Spiritual Mentorships",
        "Mentorship.xlsx / Latest Mentorship",
        (
            "FY26: Raw start dates in FY; FY27: Count from 'Latest Mentorship'"
            " tab (all city programs)"
        ),
    ),
    (
        "Row 27",
        "Baptisms",
        "CityTeam Baptisms.xlsx",
        (
            "Count of unique client baptism dates occurring in FY (all city"
            " programs)"
        ),
    ),
    (
        "Row 28",
        "Sobriety Rate",
        "All Programs All Alumni KPIs.xlsx",
        (
            "Proportion of prior FY Renew graduates sustaining post-grad"
            " sobriety"
        ),
    ),
    (
        "Row 29",
        "Case Management Hours",
        "Formula Calculation",
        (
            "(Row 7 + Row 9) * Multiplier (FY26: 52/104; FY27: 1.0/2.0 *"
            " weeks_to_date)"
        ),
    ),
    (
        "Row 31",
        "Jobs Attained",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        "Count of unique clients acquiring job in FY (all city programs)",
    ),
    (
        "Row 32",
        "Credentials Attained",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        (
            "Total count of 35 industry cert dates earned in FY (all city"
            " programs)"
        ),
    ),
    (
        "Row 33",
        "Essential Documents",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        (
            "Combined count of State ID, Birth Cert, and SS Card dates in FY"
            " (all city programs)"
        ),
    ),
    (
        "Row 34",
        "Driver's Licenses",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        (
            "Count of unique driver's licenses obtained in FY (all city"
            " programs)"
        ),
    ),
    (
        "Row 35",
        "High School Equivalency (HSE)",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        "Count of unique HSE dates obtained in FY (all city programs)",
    ),
    (
        "Row 36",
        "Bank Accounts Obtained",
        "All Programs @Work Fiscal Year Impacts.xlsx",
        (
            "Count of unique bank account dates obtained in FY (all city"
            " programs)"
        ),
    ),
    (
        "Row 37",
        "Housing Burden Under 30%",
        "All Programs All Alumni KPIs.xlsx",
        (
            "Proportion of prior FY Renew graduates with housing cost <= 30% of"
            " income"
        ),
    ),
    (
        "Row 38",
        "Hourly Wage Increase",
        "All Programs All Alumni KPIs.xlsx",
        (
            "Average wage difference ($/hr) from graduation to recent check-in"
            " for Renew grads"
        ),
    ),
    (
        "Row 43",
        "Justice Involved",
        "Clients in All Programs.xlsx",
        (
            "% of Row 5 enrollments (Start Date_2090) with 'justice' in Reason"
            " for reaching out"
        ),
    ),
    (
        "Row 44",
        "Homelessness",
        "Clients in All Programs.xlsx",
        (
            "% of Row 5 enrollments (Start Date_2090) with 'homeless' in Reason"
            " for reaching out"
        ),
    ),
    (
        "Row 45",
        "Substance Abuse",
        "Clients in All Programs.xlsx",
        (
            "% of Row 5 enrollments (Start Date_2090) with 'substance abuse' in"
            " Reason for reaching out"
        ),
    ),
    (
        "Row 46",
        "Domestic Violence",
        "Clients in All Programs.xlsx",
        (
            "% of Row 5 Women enrollments (Start Date_2090) with 'domestic"
            " violence' in Reason for reaching out"
        ),
    ),
]


def log(msg):
  print(msg, flush=True)


def sanitize_sheet_title(r_num, name):
  clean_name = (
      name.replace("/", "-")
      .replace("\\", "-")
      .replace("?", "")
      .replace("*", "")
      .replace("[", "")
      .replace("]", "")
      .replace(":", "")
  )
  return f"{r_num} - {clean_name}"[:31]


def clean_df_columns(df):
  if not df.empty:
    df.columns = df.columns.astype(str).str.strip()
  return df


def parse_clients_all_file(fh_or_path):
  try:
    df_temp = pd.read_excel(fh_or_path, header=None)
  except Exception:
    df_temp = pd.read_csv(fh_or_path, header=None)
  headers = []
  for col_idx in range(df_temp.shape[1]):
    h0, h1 = (
        str(df_temp.iloc[0, col_idx]).strip(),
        str(df_temp.iloc[1, col_idx]).strip(),
    )
    headers.append(
        h0
        if h0 and h0.lower() != "nan"
        else (h1 if h1 and h1.lower() != "nan" else f"Unnamed_{col_idx}")
    )
  df_clean = df_temp.iloc[2:].copy()
  df_clean.columns = headers
  return clean_df_columns(df_clean)


def get_adult_occupancy_value(df_occ, prog_list, target_categories=None):
  if df_occ.empty or not prog_list:
    return 0
  df = clean_df_columns(df_occ.copy())
  col_map = {str(c).lower().strip(): c for c in df.columns}
  cat_col, prog_col, cnt_col = (
      col_map.get("category"),
      col_map.get("program"),
      col_map.get("count"),
  )
  if not prog_col or not cnt_col:
    return 0
  if target_categories and cat_col:
    cats_lower = [str(c).strip().lower() for c in target_categories]
    if any("turning point" in c for c in cats_lower):
      df = df[
          df[cat_col]
          .astype(str)
          .str.strip()
          .str.lower()
          .str.contains("turning point")
      ]
    else:
      df = df[
          df[cat_col].astype(str).str.strip().str.lower().isin(cats_lower)
      ]
  prog_lower = {str(p).strip().lower() for p in prog_list}
  df_sub = df[
      df[prog_col].astype(str).str.strip().str.lower().isin(prog_lower)
  ].copy()
  if df_sub.empty:
    return 0
  clean_counts = (
      df_sub[cnt_col].astype(str).str.replace(",", "").str.strip()
  )
  return int(pd.to_numeric(clean_counts, errors="coerce").fillna(0).sum())


def calculate_case_management_hours(
    df_occ, renew_progs, tp_progs, fy_key, weeks_to_date=52.0, is_chester=False
):
  renew_occ = get_adult_occupancy_value(
      df_occ, renew_progs, target_categories=["Renew"]
  )
  tp_occ = get_adult_occupancy_value(
      df_occ,
      tp_progs,
      target_categories=["Men Turning Point", "Women Turning Point"],
  )
  total_occ = renew_occ + tp_occ
  weekly_rate = 2.0 if is_chester else 1.0
  multiplier = (
      (104.0 if is_chester else 52.0)
      if fy_key == "FY26"
      else (weekly_rate * weeks_to_date)
  )
  return int(round(total_occ * multiplier))


def get_intern_program_value(row, col_map=None):
  if col_map:
    col_pg = next(
        (
            col_map[c]
            for c in col_map
            if "post graduate program association" in c
        ),
        None,
    )
    col_intern = next(
        (col_map[c] for c in col_map if "intern program" in c), None
    )
    val_pg = (
        str(row[col_pg]).strip()
        if col_pg and col_pg in row and pd.notna(row[col_pg])
        else ""
    )
    val_intern = (
        str(row[col_intern]).strip()
        if col_intern and col_intern in row and pd.notna(row[col_intern])
        else ""
    )
  else:
    post_grad = row.get(
        "Post Graduate Program Association_6619"
    ) or row.get("Post Graduate Program Association")
    intern_prog = row.get("Intern Program_6619") or row.get("Intern Program")
    val_pg = str(post_grad).strip() if pd.notna(post_grad) else ""
    val_intern = str(intern_prog).strip() if pd.notna(intern_prog) else ""

  invalid_tokens = {"", "nan", "none", "nat", "<na>"}
  pg_valid = val_pg if val_pg.lower() not in invalid_tokens else None
  intern_valid = val_intern if val_intern.lower() not in invalid_tokens else None
  return pg_valid if pg_valid else (intern_valid if intern_valid else "")


def is_matching_program(
    row, prog_col, target_progs_lower, col_map, allow_intern_fallback=False
):
  p_val = str(row[prog_col]).strip().lower() if pd.notna(row[prog_col]) else ""
  if p_val in target_progs_lower:
    return True
  if allow_intern_fallback and (p_val in INTERN_PROGRAM_NAMES):
    secondary_val = get_intern_program_value(row, col_map).lower()
    if secondary_val in target_progs_lower:
      return True
  return False


def matches_all_city_programs(row, loc_label, col_map):
  """
    Resolves intern programs via Post Graduate Program Association / Intern Program,
    and assigns every record to its exact City/Gender column across all programs.
    """
  prog_col = next(
      (
          col_map[c]
          for c in ["program enrolling_2091", "program enrolling", "program"]
          if c in col_map
      ),
      None,
  )
  if not prog_col or pd.isna(row[prog_col]):
    return False

  p_val = str(row[prog_col]).strip().lower()
  if p_val in INTERN_PROGRAM_NAMES:
    resolved_prog = get_intern_program_value(row, col_map).strip().lower()
    if resolved_prog:
      p_val = resolved_prog

  return p_val in ALL_CITY_PROGRAMS_MAP.get(loc_label, set())


def clean_wage(val):
  if pd.isna(val):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  try:
    return float(str(val).replace("$", "").replace(",", "").strip())
  except:
    return 0.0


def combine_columns_with_priority(
    df, col_new, col_old, target_col, win_start, win_end
):
  invalid_tokens = ["", "nan", "none", "nat", "<na>"]
  s_new = (
      df[col_new].astype(str).str.strip().str.lower()
      if col_new in df.columns
      else pd.Series("", index=df.index)
  )
  s_new_clean = (
      df[col_new].where(~s_new.isin(invalid_tokens), np.nan)
      if col_new in df.columns
      else pd.Series(np.nan, index=df.index)
  )
  dt_new = pd.to_datetime(s_new_clean, errors="coerce")
  dt_new_valid = dt_new.where((dt_new >= win_start) & (dt_new <= win_end), pd.NaT)

  s_old = (
      df[col_old].astype(str).str.strip().str.lower()
      if col_old in df.columns
      else pd.Series("", index=df.index)
  )
  s_old_clean = (
      df[col_old].where(~s_old.isin(invalid_tokens), np.nan)
      if col_old in df.columns
      else pd.Series(np.nan, index=df.index)
  )
  dt_old = pd.to_datetime(s_old_clean, errors="coerce")
  dt_old_valid = dt_old.where((dt_old >= win_start) & (dt_old <= win_end), pd.NaT)

  df[target_col] = dt_new_valid.fillna(dt_old_valid)
  return df


# ============================================================
#  PROCESSED DEDUPLICATED METRIC LOGIC
# ============================================================


def count_residential_enrollments(
    df, fy_start, fy_end, prog_list, loc_label, fy_label
):
  if df.empty or not prog_list:
    return 0, pd.DataFrame()
  col_map = {str(c).lower().strip(): c for c in df.columns}
  id_col = next(
      (
          col_map[cand]
          for cand in [
              "record id_102",
              "record id_2086",
              "record id_152",
              "record id",
          ]
          if cand in col_map
      ),
      None,
  )
  date_col = next(
      (
          col_map[cand]
          for cand in ["start date_2090", "start date"]
          if cand in col_map
      ),
      None,
  )
  prog_col = next(
      (
          col_map[cand]
          for cand in ["program enrolling_2091", "program enrolling", "program"]
          if cand in col_map
      ),
      None,
  )
  if not date_col or not prog_col:
    return 0, pd.DataFrame()

  starts = pd.to_datetime(df[date_col], errors="coerce")
  df_fy = df[(starts >= fy_start) & (starts <= fy_end)].copy()
  if df_fy.empty:
    return 0, pd.DataFrame()

  prog_set_lower = {str(p).strip().lower() for p in prog_list}
  mask = df_fy.apply(
      lambda row: is_matching_program(
          row, prog_col, prog_set_lower, col_map, allow_intern_fallback=False
      ),
      axis=1,
  )
  df_res = df_fy[mask].copy()
  if df_res.empty:
    return 0, pd.DataFrame()

  # Deduplicate strictly by [Record ID, Program Enrolling]
  dedup_cols = [c for c in [id_col, prog_col] if c and c in df_res.columns]
  if dedup_cols:
    df_res = (
        df_res.dropna(subset=dedup_cols)
        .drop_duplicates(subset=dedup_cols, keep="first")
        .copy()
    )

  df_res["Counted_Location"], df_res["Counted_Fiscal_Year"] = (
      loc_label,
      fy_label,
  )
  return len(df_res), df_res


def compute_reason_percentages(
    df_row5_cohort, df_clients_all, is_women_col, loc_name, fy_key
):
  empty_res = {
      "justice": "0.00%",
      "homeless": "0.00%",
      "substance": "0.00%",
      "dv": "0.00%" if is_women_col else "",
  }, {
      43: pd.DataFrame(),
      44: pd.DataFrame(),
      45: pd.DataFrame(),
      46: pd.DataFrame(),
  }

  if df_row5_cohort.empty or df_clients_all.empty:
    return empty_res

  col_map_c = {str(c).lower().strip(): c for c in df_clients_all.columns}
  col_map_r5 = {str(c).lower().strip(): c for c in df_row5_cohort.columns}

  id_col_c = next(
      (
          col_map_c[cand]
          for cand in [
              "record id_102",
              "record id_2086",
              "record id_152",
              "record id",
          ]
          if cand in col_map_c
      ),
      None,
  )
  id_col_r5 = next(
      (
          col_map_r5[cand]
          for cand in [
              "record id_102",
              "record id_2086",
              "record id_152",
              "record id",
          ]
          if cand in col_map_r5
      ),
      None,
  )
  reason_col = next(
      (
          col_map_c[cand]
          for cand in [
              "reason(s) for reaching out to cityteam_7314",
              "reason(s) for reaching out to cityteam",
          ]
          if cand in col_map_c
      ),
      None,
  )

  if not id_col_c or not id_col_r5 or not reason_col:
    return empty_res

  df_merged = df_row5_cohort.merge(
      df_clients_all[[id_col_c, reason_col]].drop_duplicates(
          subset=[id_col_c]
      ),
      left_on=id_col_r5,
      right_on=id_col_c,
      how="left",
  )
  denominator = len(df_merged)
  if denominator == 0:
    return empty_res

  reasons_series = df_merged[reason_col].astype(str).str.lower()
  m_just = reasons_series.str.contains("justice", na=False)
  m_home = reasons_series.str.contains("homeless", na=False)
  m_sub = reasons_series.str.contains("substance abuse", na=False)
  m_dv = (
      reasons_series.str.contains("domestic violence", na=False)
      if is_women_col
      else pd.Series(False, index=df_merged.index)
  )

  metrics_out = {
      "justice": f"{(m_just.sum() / denominator) * 100:.2f}%",
      "homeless": f"{(m_home.sum() / denominator) * 100:.2f}%",
      "substance": f"{(m_sub.sum() / denominator) * 100:.2f}%",
      "dv": f"{(m_dv.sum() / denominator) * 100:.2f}%" if is_women_col else "",
  }

  detailed_dfs = {}
  for r_num, mask in [(43, m_just), (44, m_home), (45, m_sub), (46, m_dv)]:
    df_sub = df_merged[mask].copy()
    if not df_sub.empty:
      df_sub["Counted_Location"], df_sub["Counted_Fiscal_Year"] = (
          loc_name,
          fy_key,
      )
    detailed_dfs[r_num] = df_sub

  return metrics_out, detailed_dfs


def count_native_graduations(
    df,
    fy_start,
    fy_end,
    prog_list,
    valid_reasons,
    loc_label,
    fy_label,
    allow_intern_fallback=False,
):
  if df.empty or not prog_list:
    return 0, pd.DataFrame()
  col_map = {str(c).lower().strip(): c for c in df.columns}
  id_col = next(
      (
          col_map[cand]
          for cand in ["record id_102", "record id_2086", "record id"]
          if cand in col_map
      ),
      None,
  )
  exit_col = next(
      (
          col_map[cand]
          for cand in ["exit date_2100", "exit date"]
          if cand in col_map
      ),
      None,
  )
  reason_col = next(
      (
          col_map[cand]
          for cand in [
              "primary reason for exit_2102",
              "primary reason for exit",
              "exit reason",
          ]
          if cand in col_map
      ),
      None,
  )
  prog_col = next(
      (
          col_map[cand]
          for cand in ["program enrolling_2091", "program enrolling", "program"]
          if cand in col_map
      ),
      None,
  )
  if not exit_col or not prog_col or not reason_col:
    return 0, pd.DataFrame()

  exits = pd.to_datetime(df[exit_col], errors="coerce")
  prog_set_lower = {str(p).strip().lower() for p in prog_list}
  reasons_lower = {str(r).strip().lower() for r in valid_reasons}

  mask = (
      (exits >= fy_start)
      & (exits <= fy_end)
      & df.apply(
          lambda row: is_matching_program(
              row,
              prog_col,
              prog_set_lower,
              col_map,
              allow_intern_fallback=allow_intern_fallback,
          ),
          axis=1,
      )
      & (
          df[reason_col]
          .astype(str)
          .str.strip()
          .str.lower()
          .isin(reasons_lower)
      )
  )
  df_sub = df[mask].copy()
  if df_sub.empty:
    return 0, pd.DataFrame()

  if id_col and id_col in df_sub.columns:
    df_sub = (
        df_sub.dropna(subset=[id_col])
        .drop_duplicates(subset=[id_col], keep="first")
        .copy()
    )

  df_sub["Counted_Location"], df_sub["Counted_Fiscal_Year"] = loc_label, fy_label
  return len(df_sub), df_sub


def count_native_tp_housed(
    df_tp_raw, fy_start, fy_end, tp_progs, loc_label, fy_label
):
  if df_tp_raw.empty or not tp_progs:
    return 0, pd.DataFrame()
  df = df_tp_raw.copy()
  col_map = {str(c).lower().strip(): c for c in df.columns}
  id_col = next(
      (
          col_map[cand]
          for cand in ["record id_102", "record id"]
          if cand in col_map
      ),
      None,
  )
  prog_col = next(
      (
          col_map[cand]
          for cand in ["program enrolling_2091", "program enrolling", "program"]
          if cand in col_map
      ),
      None,
  )
  housing_check_col = next(
      (
          col_map[cand]
          for cand in ["date of housing check_5141", "date of housing check"]
          if cand in col_map
      ),
      None,
  )
  exit_col = next(
      (
          col_map[cand]
          for cand in ["exit date_2100", "exit date"]
          if cand in col_map
      ),
      None,
  )
  h_new_col = next(
      (col_map[cand] for cand in ["successfully housed_7011"] if cand in col_map),
      None,
  )
  h_old_col = next(
      (
          col_map[cand]
          for cand in ["successfully housed (is housing healthy?)_4368"]
          if cand in col_map
      ),
      None,
  )

  if not prog_col or (not housing_check_col and not exit_col):
    return 0, pd.DataFrame()

  dt_housing_all = (
      pd.to_datetime(df[housing_check_col], errors="coerce")
      if housing_check_col
      else pd.Series(pd.NaT, index=df.index)
  )

  def eval_housed(idx, row):
    val_n = str(row.get(h_new_col, "")).strip().lower() if h_new_col else ""
    if val_n in ["yes", "1", "true"]:
      return 1
    val_o = str(row.get(h_old_col, "")).strip().lower() if h_old_col else ""
    if val_o in ["yes", "1", "true"]:
      return 1
    # Count as housed if Date of Housing Check_5141 is populated (not null)
    if pd.notna(dt_housing_all.loc[idx]):
      return 1
    return 0

  df["_is_housed"] = [eval_housed(idx, row) for idx, row in df.iterrows()]
  df_h = df[df["_is_housed"] == 1].copy()
  if df_h.empty:
    return 0, pd.DataFrame()

  tp_lower = {str(p).strip().lower() for p in tp_progs}
  mask = df_h.apply(
      lambda row: is_matching_program(
          row, prog_col, tp_lower, col_map, allow_intern_fallback=True
      ),
      axis=1,
  )
  df_sub = df_h[mask].copy()
  if df_sub.empty:
    return 0, pd.DataFrame()

  # Prioritize Date of Housing Check_5141, falling back to Exit Date_2100
  dt_housing = (
      pd.to_datetime(df_sub[housing_check_col], errors="coerce")
      if housing_check_col
      else pd.Series(pd.NaT, index=df_sub.index)
  )
  dt_exit = (
      pd.to_datetime(df_sub[exit_col], errors="coerce")
      if exit_col
      else pd.Series(pd.NaT, index=df_sub.index)
  )
  effective_dates = dt_housing.fillna(dt_exit)

  df_valid = df_sub[
      (effective_dates >= fy_start) & (effective_dates <= fy_end)
  ].copy()
  if df_valid.empty:
    return 0, pd.DataFrame()

  if id_col and id_col in df_valid.columns:
    df_valid = (
        df_valid.dropna(subset=[id_col])
        .drop_duplicates(subset=[id_col], keep="first")
        .copy()
    )

  df_valid["Counted_Location"], df_valid["Counted_Fiscal_Year"] = (
      loc_label,
      fy_label,
  )
  return len(df_valid), df_valid


def count_latest_mentorship_occ(df_ment_occ, loc_label, fy_label):
  if df_ment_occ.empty:
    return 0, pd.DataFrame()
  df = df_ment_occ.copy()
  if "Program" not in df.columns or "Count" not in df.columns:
    return 0, pd.DataFrame()
  progs_lower = ALL_CITY_PROGRAMS_MAP.get(loc_label, set())
  if not progs_lower:
    return 0, pd.DataFrame()
  df_sub = df[
      df["Program"].astype(str).str.strip().str.lower().isin(progs_lower)
  ].copy()
  if df_sub.empty:
    return 0, pd.DataFrame()
  df_sub["Counted_Location"], df_sub["Counted_Fiscal_Year"] = loc_label, fy_label
  cnt = int(pd.to_numeric(df_sub["Count"], errors="coerce").fillna(0).sum())
  return cnt, df_sub


def count_native_spiritual_metric(
    df_raw, metric_label, fy_start, fy_end, loc_label, fy_label
):
  if df_raw.empty:
    return 0, pd.DataFrame()
  df = df_raw.copy()
  col_map = {str(c).lower().strip(): c for c in df.columns}
  id_col = next(
      (
          col_map[cand]
          for cand in ["record id_102", "record id"]
          if cand in col_map
      ),
      None,
  )
  prog_col = next(
      (
          col_map[cand]
          for cand in ["program enrolling_2091", "program enrolling", "program"]
          if cand in col_map
      ),
      None,
  )
  if not prog_col:
    return 0, pd.DataFrame()

  if metric_label == "Mentorship":
    date_new, date_old = col_map.get(
        "start date for mentoring_7199"
    ), col_map.get("start date for mentoring_4365")
    dt_new = (
        pd.to_datetime(df[date_new], errors="coerce")
        if date_new
        else pd.Series(pd.NaT, index=df.index)
    )
    dt_old = (
        pd.to_datetime(df[date_old], errors="coerce")
        if date_old
        else pd.Series(pd.NaT, index=df.index)
    )
    df["_event_date"] = dt_new.fillna(dt_old)
    flag_new, flag_old = col_map.get(
        "acquired spiritual mentor_7196"
    ), col_map.get("acquired spiritual mentor_5136")
    v_new = (
        df[flag_new].astype(str).str.strip()
        if flag_new
        else pd.Series("", index=df.index)
    )
    v_old = (
        df[flag_old].astype(str).str.strip()
        if flag_old
        else pd.Series("", index=df.index)
    )
    df = df[(v_new == "Yes") | (v_old == "Yes")].copy()
  elif metric_label == "Reunifications":
    date_new, date_old = col_map.get("date of reunification_7018"), col_map.get(
        "date of reunification_5140"
    )
    dt_new = (
        pd.to_datetime(df[date_new], errors="coerce")
        if date_new
        else pd.Series(pd.NaT, index=df.index)
    )
    dt_old = (
        pd.to_datetime(df[date_old], errors="coerce")
        if date_old
        else pd.Series(pd.NaT, index=df.index)
    )
    df["_event_date"] = dt_new.fillna(dt_old)
  elif metric_label == "Baptisms":
    date_new, date_old = col_map.get("date of baptism_7016"), col_map.get(
        "baptism date_4378"
    )
    dt_new = (
        pd.to_datetime(df[date_new], errors="coerce")
        if date_new
        else pd.Series(pd.NaT, index=df.index)
    )
    dt_old = (
        pd.to_datetime(df[date_old], errors="coerce")
        if date_old
        else pd.Series(pd.NaT, index=df.index)
    )
    df["_event_date"] = dt_new.fillna(dt_old)

  if df.empty:
    return 0, pd.DataFrame()

  # Assign across all city programs (including Interns & Forward)
  mask = df.apply(
      lambda row: matches_all_city_programs(row, loc_label, col_map),
      axis=1,
  )
  df_sub = df[mask].copy()
  if df_sub.empty:
    return 0, pd.DataFrame()

  dates = pd.to_datetime(df_sub["_event_date"], errors="coerce")
  df_valid = df_sub[(dates >= fy_start) & (dates <= fy_end)].copy()
  if df_valid.empty:
    return 0, pd.DataFrame()

  if id_col and id_col in df_valid.columns:
    df_valid = (
        df_valid.dropna(subset=[id_col])
        .drop_duplicates(subset=[id_col, "_event_date"], keep="first")
        .copy()
    )

  df_valid["Counted_Location"], df_valid["Counted_Fiscal_Year"] = (
      loc_label,
      fy_label,
  )
  return len(df_valid), df_valid


def process_atwork_raw_to_kpis(df_raw, fy_start, fy_end):
  if df_raw.empty:
    return pd.DataFrame()
  df = df_raw.copy()
  col_map = {str(c).lower().strip(): c for c in df.columns}
  actual_id_col = next(
      (
          col_map[cand]
          for cand in [
              "record id_152",
              "record id_102",
              "record id_2086",
              "record id",
          ]
          if cand in col_map
      ),
      None,
  )
  actual_prog_col = next(
      (
          col_map[cand]
          for cand in ["program enrolling_2091", "program enrolling", "program"]
          if cand in col_map
      ),
      None,
  )
  if not actual_id_col or not actual_prog_col:
    return df_raw

  col_job_new, col_job_old = next(
      (col_map[c] for c in col_map if "7009" in c), None
  ), next((col_map[c] for c in col_map if "5590" in c), None)
  if col_job_new or col_job_old:
    df = combine_columns_with_priority(
        df,
        col_new=col_job_new if col_job_new else "_dummy_new",
        col_old=col_job_old if col_job_old else "_dummy_old",
        target_col="Date Job Acquired",
        win_start=fy_start,
        win_end=fy_end,
    )
  date_cols = [c for c in df.columns if "date" in str(c).lower()]
  non_date_cols = [
      c for c in df.columns if c not in date_cols and c != actual_id_col
  ]
  for c in date_cols:
    df[c] = pd.to_datetime(df[c], errors="coerce")
  agg_dict = {c: "first" for c in non_date_cols}
  agg_dict.update({c: "max" for c in date_cols})
  return df.groupby(actual_id_col, as_index=False, sort=False).agg(agg_dict)


def count_atwork_dates_from_raw(
    df_raw,
    date_col_names,
    fy_start,
    fy_end,
    loc_label,
    fy_label,
    event_based=False,
):
  if df_raw.empty:
    return 0, pd.DataFrame()
  df_sub_consolidated = process_atwork_raw_to_kpis(df_raw, fy_start, fy_end)
  if df_sub_consolidated.empty:
    return 0, pd.DataFrame()
  col_map = {str(c).lower().strip(): c for c in df_sub_consolidated.columns}

  # Assign across all city programs (including Interns & Forward)
  mask = df_sub_consolidated.apply(
      lambda row: matches_all_city_programs(row, loc_label, col_map),
      axis=1,
  )
  df_sub = df_sub_consolidated[mask].copy()
  if df_sub.empty:
    return 0, pd.DataFrame()

  if isinstance(date_col_names, str):
    date_col_names = [date_col_names]
  matched_cols = [
      col_map[c.lower().strip()]
      for c in date_col_names
      if c.lower().strip() in col_map
  ]
  if not matched_cols:
    return 0, pd.DataFrame()

  in_window_mask = pd.Series(False, index=df_sub.index)
  total_events = 0
  for col in matched_cols:
    dates = pd.to_datetime(df_sub[col], errors="coerce")
    col_in_win = (dates >= fy_start) & (dates <= fy_end)
    in_window_mask |= col_in_win
    total_events += int(col_in_win.sum())

  df_valid = df_sub[in_window_mask].copy()
  if df_valid.empty:
    return 0, pd.DataFrame()

  actual_id_col = next(
      (
          col_map[cand]
          for cand in [
              "record id_152",
              "record id_102",
              "record id_2086",
              "record id",
          ]
          if cand in col_map
      ),
      None,
  )
  if actual_id_col and actual_id_col in df_valid.columns:
    df_valid = (
        df_valid.dropna(subset=[actual_id_col])
        .drop_duplicates(subset=[actual_id_col], keep="first")
        .copy()
    )

  df_valid["Counted_Location"], df_valid["Counted_Fiscal_Year"] = (
      loc_label,
      fy_label,
  )
  return (total_events if event_based else len(df_valid)), df_valid


# --- ALUMNI METRICS (STRICT RENEW ONLY) ---
def compute_native_alumni_metrics(
    df_alumni_raw, target_fy_str, renew_progs, loc_label, fy_label
):
  if df_alumni_raw.empty or not renew_progs:
    return {"sobriety": "", "housing": "", "wage": ""}, pd.DataFrame()
  df = clean_df_columns(df_alumni_raw.copy())
  col_map = {str(c).lower().strip(): c for c in df.columns}
  name_col, exit_col = next(
      (col_map[c] for c in col_map if "name" in c), None
  ), next((col_map[c] for c in col_map if "exit date" in c), None)
  prog_col, reason_col = next(
      (
          col_map[c]
          for c in col_map
          if "program enrolling" in c or "program" in c
      ),
      None,
  ), next(
      (
          col_map[c]
          for c in col_map
          if "primary reason for exit" in c or "exit reason" in c
      ),
      None,
  )

  if not name_col or not exit_col or not prog_col or not reason_col:
    return {"sobriety": "", "housing": "", "wage": ""}, pd.DataFrame()

  df = df.dropna(subset=[name_col]).copy()
  checkin_col, wage_col = next(
      (
          col_map[c]
          for c in col_map
          if "checkin date" in c or "check-in date" in c
      ),
      None,
  ), next(
      (col_map[c] for c in col_map if "hourly wage" in c or "wage" in c), None
  )
  housing_col, sober_col = next(
      (col_map[c] for c in col_map if "housing expense" in c or "housing" in c),
      None,
  ), next((col_map[c] for c in col_map if "sober?" in c or "sober" in c), None)

  df[exit_col] = pd.to_datetime(df[exit_col], errors="coerce")
  if checkin_col and checkin_col in df.columns:
    df[checkin_col] = pd.to_datetime(df[checkin_col], errors="coerce")
  if wage_col and wage_col in df.columns:
    df[wage_col] = df[wage_col].apply(clean_wage)
  if housing_col and housing_col in df.columns:
    df[housing_col] = df[housing_col].apply(clean_wage)

  renew_lower = {str(p).strip().lower() for p in renew_progs}
  mask_renew_grads = (
      df.apply(
          lambda r: is_matching_program(
              r, prog_col, renew_lower, col_map, allow_intern_fallback=False
          ),
          axis=1,
      )
      & (df[reason_col].astype(str).str.strip().isin(GRAD_REASONS))
      & df[exit_col].notna()
  )
  df_grads = df[mask_renew_grads].copy()
  df_grads["FY_Label"] = df_grads[exit_col].apply(
      lambda dt: ""
      if pd.isna(dt)
      else f"FY{str(dt.year + 1 if dt.month >= 9 else dt.year)[-2:]}"
  )
  df_grads_target = df_grads[df_grads["FY_Label"] == target_fy_str].copy()
  if df_grads_target.empty:
    return {"sobriety": "", "housing": "", "wage": ""}, pd.DataFrame()

  df_grads_dedup = (
      df_grads_target.sort_values(by=exit_col)
      .drop_duplicates(subset=[name_col, prog_col], keep="last")
      .copy()
  )
  sobriety_vals, housing_under_30_vals, wage_diff_vals = [], [], []
  is_sober_list, h_under_30_list, wage_diff_list = [], [], []

  for _, row in df_grads_dedup.iterrows():
    c_name, c_prog, g_date = row[name_col], row[prog_col], row[exit_col]
    person_records = df[
        (df[name_col] == c_name) & (df[prog_col] == c_prog)
    ].copy()
    if checkin_col and checkin_col in person_records.columns:
      person_records = person_records.sort_values(by=checkin_col)

    sober_recs = (
        person_records[
            person_records[sober_col].isin(["Yes", "No"])
            & (person_records[checkin_col] > g_date)
        ]
        if (
            sober_col
            and checkin_col
            and sober_col in person_records.columns
            and checkin_col in person_records.columns
        )
        else pd.DataFrame()
    )

    if not sober_recs.empty:
      latest_sober_val = sober_recs.iloc[-1][sober_col]
      first_no_date, has_sustained_relapse = None, False
      for _, s_row in sober_recs.iterrows():
        c_state, c_date = s_row[sober_col], s_row[checkin_col]
        if c_state == "No":
          if first_no_date is None:
            first_no_date = c_date
        elif c_state == "Yes" and first_no_date is not None:
          if (c_date - first_no_date).days > 30:
            has_sustained_relapse = True
          first_no_date = None
      is_sober = (latest_sober_val == "Yes") and (not has_sustained_relapse)
    else:
      is_sober = True

    sobriety_vals.append(1 if is_sober else 0)
    is_sober_list.append("Yes" if is_sober else "No")
    grad_wage = (
        float(row[wage_col])
        if (wage_col and wage_col in row and pd.notna(row[wage_col]))
        else 0.0
    )
    grad_housing = (
        float(row[housing_col])
        if (housing_col and housing_col in row and pd.notna(row[housing_col]))
        else 0.0
    )

    post_recs = (
        person_records[
            person_records[checkin_col].notna()
            & (person_records[checkin_col] > g_date)
        ]
        if (checkin_col and checkin_col in person_records.columns)
        else pd.DataFrame()
    )
    if not post_recs.empty:
      latest_row = post_recs[
          post_recs[checkin_col] == post_recs[checkin_col].max()
      ].iloc[0]
      recent_wage = (
          float(latest_row[wage_col])
          if (
              wage_col
              and wage_col in latest_row
              and pd.notna(latest_row[wage_col])
          )
          else grad_wage
      )
      recent_housing = (
          float(latest_row[housing_col])
          if (
              housing_col
              and housing_col in latest_row
              and pd.notna(latest_row[housing_col])
          )
          else grad_housing
      )
    else:
      recent_wage, recent_housing = grad_wage, grad_housing

    w_diff = recent_wage - grad_wage
    wage_diff_list.append(round(w_diff, 2))

    if is_sober:
      recent_income = recent_wage * MONTHLY_HOURS
      h_under_30 = (
          1
          if (recent_housing == 0)
          else (
              1
              if (recent_income > 0 and (recent_housing / recent_income) <= 0.30)
              else 0
          )
      )
      housing_under_30_vals.append(h_under_30)
      h_under_30_list.append("Yes" if h_under_30 == 1 else "No")
      wage_diff_vals.append(w_diff)
    else:
      h_under_30_list.append("Excluded (Relapse)")

  df_grads_dedup["Calculated_Is_Sober_1Yr"] = is_sober_list
  df_grads_dedup["Calculated_Housing_Under_30"] = h_under_30_list
  df_grads_dedup["Calculated_Wage_Increase"] = wage_diff_list
  df_grads_dedup["Counted_Location"], df_grads_dedup["Counted_Fiscal_Year"] = (
      loc_label,
      fy_label,
  )

  val_r28 = float(np.mean(sobriety_vals)) if sobriety_vals else ""
  val_r37 = float(np.mean(housing_under_30_vals)) if housing_under_30_vals else ""
  val_r38 = float(np.mean(wage_diff_vals)) if wage_diff_vals else ""

  return {
      "sobriety": f"{val_r28 * 100:.2f}%" if isinstance(val_r28, float) else "",
      "housing": f"{val_r37 * 100:.2f}%" if isinstance(val_r37, float) else "",
      "wage": f"${val_r38:.2f}" if isinstance(val_r38, float) else "",
  }, df_grads_dedup


# ============================================================
#  NATIVE GOOGLE SPREADSHEET WRITER
# ============================================================


def create_separate_processed_spreadsheet(
    drive_service, gc, target_folder_id, processed_kpi_records
):
  sheet_title = "Processed Marketing Impact Data - Detailed"
  log(f"\n📊 Writing native Google Sheet: '{sheet_title}'...")

  existing_file = find_file_id(
      drive_service,
      sheet_title,
      target_folder_id,
      "application/vnd.google-apps.spreadsheet",
  )
  if existing_file:
    doc = gc.open_by_key(existing_file["id"])
    log(f"   ✓ Located existing Google Sheet (ID: {existing_file['id']}).")
  else:
    file = (
        drive_service.files()
        .create(
            body={
                "name": sheet_title,
                "mimeType": "application/vnd.google-apps.spreadsheet",
                "parents": [target_folder_id],
            },
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    doc = gc.open_by_key(file["id"])
    log(f"   ✅ Created new native Google Sheet (ID: {file['id']}).")

  for r_num, name, src, logic in ROWS_KPI_SPEC:
    tab_title = sanitize_sheet_title(r_num, name)
    df_processed = processed_kpi_records.get(r_num, pd.DataFrame())
    if df_processed.empty:
      log(f"   ⏭️ Skipping tab '{tab_title}' (No records found).")
      continue

    col_map_tab = {str(c).lower().strip(): c for c in df_processed.columns}
    id_col_tab = next(
        (
            col_map_tab[cand]
            for cand in [
                "record id_102",
                "record id_2086",
                "record id_152",
                "record id",
            ]
            if cand in col_map_tab
        ),
        None,
    )
    prog_col_tab = next(
        (
            col_map_tab[cand]
            for cand in [
                "program enrolling_2091",
                "program enrolling",
                "program",
            ]
            if cand in col_map_tab
        ),
        None,
    )
    loc_col_tab, fy_col_tab = col_map_tab.get(
        "counted_location"
    ), col_map_tab.get("counted_fiscal_year")

    # Preserve multi-program enrollment rows for Row 5 and Rows 43-46 detail tabs
    if r_num in ["Row 5", "Row 43", "Row 44", "Row 45", "Row 46"]:
      dedup_tab_cols = [
          c
          for c in [id_col_tab, prog_col_tab, loc_col_tab, fy_col_tab]
          if c and c in df_processed.columns
      ]
    else:
      dedup_tab_cols = [
          c
          for c in [id_col_tab, loc_col_tab, fy_col_tab]
          if c and c in df_processed.columns
      ]

    if dedup_tab_cols:
      df_processed = df_processed.drop_duplicates(
          subset=dedup_tab_cols, keep="first"
      )
    elif id_col_tab:
      df_processed = df_processed.drop_duplicates(
          subset=[id_col_tab], keep="first"
      )

    if df_processed.empty:
      log(f"   ⏭️ Skipping tab '{tab_title}' (Empty after deduplication).")
      continue

    try:
      ws = doc.worksheet(tab_title)
      ws.clear()
    except gspread.WorksheetNotFound:
      ws = doc.add_worksheet(title=tab_title, rows="500", cols="50")

    df_clean = df_processed.astype(object).fillna("")
    headers = [str(c) for c in df_clean.columns]
    records = [
        [
            ""
            if (
                pd.isna(v)
                or v is None
                or str(v).strip().lower() in ["nan", "<na>", "nat", "none"]
            )
            else str(v)
            for v in row
        ]
        for row in df_clean.values.tolist()
    ]
    ws.update(
        range_name="A1",
        values=[headers] + records,
        value_input_option="USER_ENTERED",
    )
    log(f"   ✅ Written {len(records)} records on tab '{tab_title}'.")

  log(
      "✅ Successfully written all non-blank detail tabs to 'Processed"
      " Marketing Impact Data - Detailed'!"
  )


# ============================================================
#  EXECUTION ENGINE
# ============================================================


def thread_safe_download(task):
  key, fname, folder_id = task
  for attempt in range(1, 4):
    try:
      drive_service, _, _ = get_services()
      fh, file_id, _ = download_drive_file(drive_service, fname, folder_id)
      return key, fh, file_id
    except Exception as e:
      if attempt == 3:
        raise e
      time.sleep(2)


def process_fiscal_year(fy_key, fy_info, gc, dfs, processed_kpi_records):
  fy_start, fy_end, sheet_name = (
      fy_info["start"],
      fy_info["end"],
      fy_info["sheet_name"],
  )
  log(
      f"\n📊 Processing Metrics for {fy_key} ({fy_start.strftime('%Y-%m-%d')}"
      f" to {fy_end.strftime('%Y-%m-%d')})..."
  )
  doc = gc.open_by_key(DESTINATION_SHEET_ID)
  ws = doc.worksheet(sheet_name)
  existing_sheet_data = ws.get_all_values() if fy_key == "FY26" else []

  cell_updates = []
  target_grad_fy = f"FY{int(fy_key.replace('FY', '')) - 1}"
  weeks_to_date_fy27 = max((fy_end - fy_start).days / 7.0, 0.0)

  for loc_name, col_idx in COLUMN_INDICES.items():
    renew_progs, tp_progs, fwd_intern_progs, connect_progs = (
        PROGRAM_MAP[loc_name]["renew"],
        PROGRAM_MAP[loc_name]["tp"],
        PROGRAM_MAP[loc_name]["fwd_intern"],
        PROGRAM_MAP[loc_name].get("connect", []),
    )
    all_loc_progs, res_progs = (
        renew_progs + tp_progs + fwd_intern_progs + connect_progs,
        renew_progs + tp_progs,
    )
    df_res_raw_combined = pd.concat(
        [dfs["renew_raw"], dfs["tp_raw"]], ignore_index=True
    )

    # ROWS 5-6: ENROLLMENTS (Filtered by Start Date_2090)
    val_r5, df_r5 = count_residential_enrollments(
        df_res_raw_combined, fy_start, fy_end, res_progs, loc_name, fy_key
    )
    cell_updates.append(
        {"row": ROW_RES_NEW_ADULT_ENROLLMENTS, "col": col_idx, "val": val_r5}
    )
    if not df_r5.empty:
      processed_kpi_records["Row 5"] = pd.concat(
          [processed_kpi_records.get("Row 5", pd.DataFrame()), df_r5],
          ignore_index=True,
      )

    val_r6, col_map_res = 0, {
        str(c).lower().strip(): c for c in df_res_raw_combined.columns
    }
    if "number of children" in col_map_res:
      n_kids_col = col_map_res["number of children"]
      d_col, p_col = next(
          (
              col_map_res[c]
              for c in ["start date_2090", "start date"]
              if c in col_map_res
          ),
          None,
      ), next(
          (
              col_map_res[c]
              for c in ["program enrolling_2091", "program enrolling"]
              if c in col_map_res
          ),
          None,
      )
      if d_col and p_col:
        starts = pd.to_datetime(df_res_raw_combined[d_col], errors="coerce")
        res_progs_lower = {str(p).strip().lower() for p in res_progs}
        mask = df_res_raw_combined.apply(
            lambda r: is_matching_program(
                r, p_col, res_progs_lower, col_map_res
            ),
            axis=1,
        )
        in_fy = df_res_raw_combined[
            mask & (starts >= fy_start) & (starts <= fy_end)
        ].copy()
        val_r6 = int(in_fy[n_kids_col].sum()) if not in_fy.empty else 0
        if not in_fy.empty:
          in_fy["Counted_Location"], in_fy["Counted_Fiscal_Year"] = (
              loc_name,
              fy_key,
          )
          processed_kpi_records["Row 6"] = pd.concat(
              [processed_kpi_records.get("Row 6", pd.DataFrame()), in_fy],
              ignore_index=True,
          )
    cell_updates.append(
        {"row": ROW_RES_NEW_CHILD_ENROLLMENTS, "col": col_idx, "val": val_r6}
    )

    # ROWS 7, 9, 11 (OCCUPANCY) & 14, 15, 16 (NIGHTS OF SHELTER)
    if fy_key == "FY26":

      def parse_existing_cell(r_num, c_num):
        try:
          val_str = (
              str(existing_sheet_data[r_num - 1][c_num - 1])
              .strip()
              .replace(",", "")
          )
          return float(val_str) if val_str else 0.0
        except Exception:
          return 0.0

      r7_manual = parse_existing_cell(ROW_RENEW_ADULT_AVG_OCCUPANCY, col_idx)
      r9_manual = parse_existing_cell(ROW_TP_ADULT_AVG_OCCUPANCY, col_idx)
      r11_manual = parse_existing_cell(
          ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY, col_idx
      )
      cell_updates.extend([
          {
              "row": ROW_RENEW_ADULT_NIGHTS,
              "col": col_idx,
              "val": int(round(r7_manual * 365)),
          },
          {
              "row": ROW_TP_ADULT_NIGHTS,
              "col": col_idx,
              "val": int(round(r9_manual * 365)),
          },
          {
              "row": ROW_FWD_INTERN_NIGHTS,
              "col": col_idx,
              "val": int(round(r11_manual * 365)),
          },
      ])
    elif fy_key == "FY27":
      val_r7 = get_adult_occupancy_value(
          dfs["occ"], renew_progs, target_categories=["Renew"]
      )
      val_r9 = get_adult_occupancy_value(
          dfs["occ"],
          tp_progs,
          target_categories=["Men Turning Point", "Women Turning Point"],
      )
      val_r11 = get_adult_occupancy_value(
          dfs["occ"], all_loc_progs, target_categories=["Forward", "Intern"]
      )
      cell_updates.extend([
          {"row": ROW_RENEW_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r7},
          {"row": ROW_TP_ADULT_AVG_OCCUPANCY, "col": col_idx, "val": val_r9},
          {
              "row": ROW_FWD_INTERN_ADULT_AVG_OCCUPANCY,
              "col": col_idx,
              "val": val_r11,
          },
          {
              "row": ROW_RENEW_ADULT_NIGHTS,
              "col": col_idx,
              "val": int(round(val_r7 * 7 * weeks_to_date_fy27)),
          },
          {
              "row": ROW_TP_ADULT_NIGHTS,
              "col": col_idx,
              "val": int(round(val_r9 * 7 * weeks_to_date_fy27)),
          },
          {
              "row": ROW_FWD_INTERN_NIGHTS,
              "col": col_idx,
              "val": int(round(val_r11 * 7 * weeks_to_date_fy27)),
          },
      ])

    # ROWS 21-23: GRADUATIONS & HOUSING
    val_r21, df_r21 = count_native_graduations(
        dfs["renew_raw"],
        fy_start,
        fy_end,
        renew_progs,
        GRAD_REASONS,
        loc_name,
        fy_key,
        allow_intern_fallback=False,
    )
    cell_updates.append(
        {"row": ROW_RENEW_GRADUATIONS, "col": col_idx, "val": val_r21}
    )
    if not df_r21.empty:
      processed_kpi_records["Row 21"] = pd.concat(
          [processed_kpi_records.get("Row 21", pd.DataFrame()), df_r21],
          ignore_index=True,
      )

    val_r22, df_r22 = count_native_graduations(
        dfs["tp_raw"],
        fy_start,
        fy_end,
        tp_progs,
        ["Graduation"],
        loc_name,
        fy_key,
        allow_intern_fallback=False,
    )
    cell_updates.append(
        {"row": ROW_TP_GRADUATIONS, "col": col_idx, "val": val_r22}
    )
    if not df_r22.empty:
      processed_kpi_records["Row 22"] = pd.concat(
          [processed_kpi_records.get("Row 22", pd.DataFrame()), df_r22],
          ignore_index=True,
      )

    val_r23, df_r23 = count_native_tp_housed(
        dfs["tp_raw"], fy_start, fy_end, tp_progs, loc_name, fy_key
    )
    cell_updates.append(
        {"row": ROW_TP_SUCCESSFULLY_HOUSED, "col": col_idx, "val": val_r23}
    )
    if not df_r23.empty:
      processed_kpi_records["Row 23"] = pd.concat(
          [processed_kpi_records.get("Row 23", pd.DataFrame()), df_r23],
          ignore_index=True,
      )

    # ROWS 25-27: SPIRITUAL METRICS (All programs by City & Gender, including Interns)
    val_r25, df_r25 = count_native_spiritual_metric(
        dfs["reun_raw"],
        "Reunifications",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
    )
    cell_updates.append(
        {"row": ROW_FAMILIES_REUNIFIED, "col": col_idx, "val": val_r25}
    )
    if not df_r25.empty:
      processed_kpi_records["Row 25"] = pd.concat(
          [processed_kpi_records.get("Row 25", pd.DataFrame()), df_r25],
          ignore_index=True,
      )

    val_r26, df_r26 = (
        count_latest_mentorship_occ(dfs["ment_occ"], loc_name, fy_key)
        if fy_key == "FY27"
        else count_native_spiritual_metric(
            dfs["ment_raw"],
            "Mentorship",
            fy_start,
            fy_end,
            loc_name,
            fy_key,
        )
    )
    cell_updates.append(
        {"row": ROW_SPIRITUAL_MENTORSHIPS, "col": col_idx, "val": val_r26}
    )
    if not df_r26.empty:
      processed_kpi_records["Row 26"] = pd.concat(
          [processed_kpi_records.get("Row 26", pd.DataFrame()), df_r26],
          ignore_index=True,
      )

    val_r27, df_r27 = count_native_spiritual_metric(
        dfs["bap_raw"],
        "Baptisms",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
    )
    cell_updates.append({"row": ROW_BAPTISMS, "col": col_idx, "val": val_r27})
    if not df_r27.empty:
      processed_kpi_records["Row 27"] = pd.concat(
          [processed_kpi_records.get("Row 27", pd.DataFrame()), df_r27],
          ignore_index=True,
      )

    # ROWS 28-29: ALUMNI SOBRIETY & CASE MANAGEMENT HOURS
    alum_metrics_renew, df_alum = compute_native_alumni_metrics(
        dfs["alumni_raw"], target_grad_fy, renew_progs, loc_name, fy_key
    )
    cell_updates.extend([
        {
            "row": ROW_SOBRIETY_RATE,
            "col": col_idx,
            "val": alum_metrics_renew["sobriety"],
        },
        {
            "row": ROW_HOUSING_BURDEN_1YR,
            "col": col_idx,
            "val": alum_metrics_renew["housing"],
        },
        {
            "row": ROW_WAGE_INCREASE,
            "col": col_idx,
            "val": alum_metrics_renew["wage"],
        },
    ])
    if not df_alum.empty:
      for r_key in ["Row 28", "Row 37", "Row 38"]:
        processed_kpi_records[r_key] = pd.concat(
            [processed_kpi_records.get(r_key, pd.DataFrame()), df_alum],
            ignore_index=True,
        )

    val_r29 = calculate_case_management_hours(
        dfs["occ"],
        renew_progs,
        tp_progs,
        fy_key=fy_key,
        weeks_to_date=weeks_to_date_fy27,
        is_chester=("Chester" in loc_name),
    )
    cell_updates.append(
        {"row": ROW_CASE_MANAGEMENT_HOURS, "col": col_idx, "val": val_r29}
    )

    # ROWS 31-36: @WORK METRICS (All programs by City & Gender, including Interns)
    val_r31, df_r31 = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "Date Job Acquired",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cell_updates.append(
        {"row": ROW_JOBS_ATTAINED, "col": col_idx, "val": val_r31}
    )
    if not df_r31.empty:
      processed_kpi_records["Row 31"] = pd.concat(
          [processed_kpi_records.get("Row 31", pd.DataFrame()), df_r31],
          ignore_index=True,
      )

    val_r32, df_r32 = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        IC_DATE_COLS,
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=True,
    )
    cell_updates.append(
        {"row": ROW_CREDENTIALS_ATTAINED, "col": col_idx, "val": val_r32}
    )
    if not df_r32.empty:
      processed_kpi_records["Row 32"] = pd.concat(
          [processed_kpi_records.get("Row 32", pd.DataFrame()), df_r32],
          ignore_index=True,
      )

    cnt_bc, df_bc = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "Birth Certificate Obtained Date_6638",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cnt_id, df_id = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "State ID Obtained Date_6637",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cnt_ss, df_ss = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "Social Security Card Obtained Date_6639",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cell_updates.append({
        "row": ROW_IDS_CARDS_DOCS,
        "col": col_idx,
        "val": (cnt_bc + cnt_id + cnt_ss),
    })
    df_docs = pd.concat([df_bc, df_id, df_ss], ignore_index=True)
    if not df_docs.empty:
      processed_kpi_records["Row 33"] = pd.concat(
          [processed_kpi_records.get("Row 33", pd.DataFrame()), df_docs],
          ignore_index=True,
      )

    val_r34, df_r34 = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "Driver's License Date_5071",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cell_updates.append(
        {"row": ROW_DRIVERS_LICENSES, "col": col_idx, "val": val_r34}
    )
    if not df_r34.empty:
      processed_kpi_records["Row 34"] = pd.concat(
          [processed_kpi_records.get("Row 34", pd.DataFrame()), df_r34],
          ignore_index=True,
      )

    val_r35, df_r35 = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "High School Equivalency (HSE) Date_5110",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cell_updates.append({"row": ROW_HSE_ATTAINED, "col": col_idx, "val": val_r35})
    if not df_r35.empty:
      processed_kpi_records["Row 35"] = pd.concat(
          [processed_kpi_records.get("Row 35", pd.DataFrame()), df_r35],
          ignore_index=True,
      )

    val_r36, df_r36 = count_atwork_dates_from_raw(
        dfs["atwork_raw"],
        "Bank Account Obtained Date_6640",
        fy_start,
        fy_end,
        loc_name,
        fy_key,
        event_based=False,
    )
    cell_updates.append(
        {"row": ROW_BANK_ACCOUNTS, "col": col_idx, "val": val_r36}
    )
    if not df_r36.empty:
      processed_kpi_records["Row 36"] = pd.concat(
          [processed_kpi_records.get("Row 36", pd.DataFrame()), df_r36],
          ignore_index=True,
      )

    # ROWS 43-46: REASON FOR REACHING OUT PERCENTAGES (Based on Row 5 Start Date_2090 Cohort)
    reason_pcts, reason_detail_dfs = compute_reason_percentages(
        df_row5_cohort=df_r5,
        df_clients_all=dfs.get("clients_all_raw", pd.DataFrame()),
        is_women_col=("Women" in loc_name),
        loc_name=loc_name,
        fy_key=fy_key,
    )

    cell_updates.extend([
        {
            "row": ROW_JUSTICE_INVOLVED,
            "col": col_idx,
            "val": reason_pcts["justice"],
        },
        {
            "row": ROW_HOMELESSNESS,
            "col": col_idx,
            "val": reason_pcts["homeless"],
        },
        {
            "row": ROW_SUBSTANCE_ABUSE,
            "col": col_idx,
            "val": reason_pcts["substance"],
        },
        {
            "row": ROW_DOMESTIC_VIOLENCE,
            "col": col_idx,
            "val": reason_pcts["dv"],
        },
    ])

    for r_num, df_det in reason_detail_dfs.items():
      if not df_det.empty:
        r_key = f"Row {r_num}"
        processed_kpi_records[r_key] = pd.concat(
            [processed_kpi_records.get(r_key, pd.DataFrame()), df_det],
            ignore_index=True,
        )

  log(f"   Writing cell updates to Google Sheet tab '{sheet_name}'...")
  gspread_data = [
      {
          "range": gspread.utils.rowcol_to_a1(item["row"], item["col"]),
          "values": [[item["val"]]],
      }
      for item in cell_updates
  ]
  ws.batch_update(gspread_data, value_input_option="USER_ENTERED")
  log(f"✅ Successfully written values directly to '{sheet_name}'!")


def run_marketing_processor():
  log("🚀 Starting Marketing Impact Data Processing across Shared Drives...")
  main_drive_service, gc, _ = get_services()

  log("📂 Resolving root shared folder: 'CityTeam Impact Data'...")
  try:
    impact_folder_id = resolve_folder_id(
        main_drive_service, "CityTeam Impact Data", "Input"
    )
  except Exception:
    impact_folder_id = None

  incoming_res = find_file_id(
      main_drive_service,
      "Apricot Report Incoming",
      impact_folder_id,
      "application/vnd.google-apps.folder",
  )
  input_folder_id = (
      incoming_res["id"]
      if incoming_res
      else resolve_folder_id(main_drive_service, "Apricot Report Incoming")
  )

  try:
    mkt_res = find_file_id(
        main_drive_service,
        "Marketing Impacts",
        impact_folder_id,
        "application/vnd.google-apps.folder",
    )
    if not mkt_res:
      mkt_res = find_file_id(
          main_drive_service,
          "Marketing Impacts",
          None,
          "application/vnd.google-apps.folder",
      )
    target_output_folder_id = (
        mkt_res["id"] if mkt_res else (impact_folder_id or input_folder_id)
    )
  except Exception:
    target_output_folder_id = input_folder_id

  df_occ, df_ment_occ = pd.DataFrame(), pd.DataFrame()
  try:
    kpi_folder_res = find_file_id(
        main_drive_service,
        "CityTeam KPIs",
        impact_folder_id,
        "application/vnd.google-apps.folder",
    )
    kpi_folder_id = (
        kpi_folder_res["id"] if kpi_folder_res else impact_folder_id
    )
    target_kpi_res = find_file_id(
        main_drive_service,
        "Quarterly KPI Processed Data",
        kpi_folder_id,
        "application/vnd.google-apps.folder",
    )
    quarterly_kpi_folder_id = (
        target_kpi_res["id"] if target_kpi_res else kpi_folder_id
    )
    occ_file_res = find_file_id(
        main_drive_service, FILE_OCCUPANCY_NAME, quarterly_kpi_folder_id
    ) or find_file_id(
        main_drive_service, FILE_OCCUPANCY_NAME, target_output_folder_id
    )

    if occ_file_res:
      if (
          occ_file_res.get("mimeType", "")
          == "application/vnd.google-apps.spreadsheet"
      ):
        occ_ss = gc.open_by_key(occ_file_res["id"])
        try:
          df_occ = clean_df_columns(
              pd.DataFrame(occ_ss.worksheet("Latest Occupancy").get_all_records())
          )
        except Exception:
          pass
        try:
          df_ment_occ = clean_df_columns(
              pd.DataFrame(
                  occ_ss.worksheet("Latest Mentorship").get_all_records()
              )
          )
        except Exception:
          pass
      else:
        fh_occ, _, _ = download_drive_file(
            main_drive_service,
            occ_file_res["name"],
            occ_file_res.get("parents", [input_folder_id])[0],
        )
        xls = pd.ExcelFile(fh_occ)
        if "Latest Occupancy" in xls.sheet_names:
          df_occ = clean_df_columns(
              pd.read_excel(xls, sheet_name="Latest Occupancy")
          )
        if "Latest Mentorship" in xls.sheet_names:
          df_ment_occ = clean_df_columns(
              pd.read_excel(xls, sheet_name="Latest Mentorship")
          )
  except Exception as e:
    log(f"⚠️ Warning loading occupancy data: {e}")

  download_tasks = [
      ("renew", FILE_RENEW_NAME, input_folder_id),
      ("tp", FILE_TP_NAME, input_folder_id),
      ("bap", FILE_BAPTISMS_NAME, input_folder_id),
      ("ment", FILE_MENTOR_NAME, input_folder_id),
      ("reun", FILE_REUNIF_NAME, input_folder_id),
      ("alum", FILE_ALUMNI_NAME, input_folder_id),
      ("atwork", FILE_ATWORK_NAME, input_folder_id),
      ("clients_all", FILE_CLIENTS_ALL_NAME, input_folder_id),
  ]

  downloads = {}
  with ThreadPoolExecutor(max_workers=len(download_tasks)) as executor:
    for key, fh, file_id in executor.map(thread_safe_download, download_tasks):
      downloads[key] = (fh, file_id)

  dfs = {
      "occ": df_occ,
      "ment_occ": df_ment_occ,
      "renew_raw": clean_df_columns(
          pd.read_excel(downloads["renew"][0], header=3).iloc[:, 1:]
      ),
      "tp_raw": clean_df_columns(
          pd.read_excel(downloads["tp"][0], header=3).iloc[:, 1:]
      ),
      "bap_raw": clean_df_columns(
          pd.read_excel(downloads["bap"][0], header=3).iloc[:, 1:]
      ),
      "ment_raw": clean_df_columns(
          pd.read_excel(downloads["ment"][0], header=3).iloc[:, 1:]
      ),
      "reun_raw": clean_df_columns(
          pd.read_excel(downloads["reun"][0], header=3).iloc[:, 1:]
      ),
      "alumni_raw": clean_df_columns(
          pd.read_excel(downloads["alum"][0], header=3).iloc[:, 1:]
      ),
      "atwork_raw": clean_df_columns(
          pd.read_excel(downloads["atwork"][0], header=3).iloc[:, 1:]
      ),
      "clients_all_raw": parse_clients_all_file(downloads["clients_all"][0]),
  }

  processed_kpi_records = {}
  for fy_key, fy_info in FY_CONFIGS.items():
    process_fiscal_year(fy_key, fy_info, gc, dfs, processed_kpi_records)

  create_separate_processed_spreadsheet(
      main_drive_service, gc, target_output_folder_id, processed_kpi_records
  )
  log(
      "\n🎉 Marketing Processor Execution Complete! Both Google Spreadsheets"
      " Updated in 'Marketing Impacts'!"
  )


if __name__ == "__main__":
  run_marketing_processor()