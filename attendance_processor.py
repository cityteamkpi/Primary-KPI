#!/usr/bin/env python3
"""
Student Attendance Processing Pipeline
Reads monthly attendance tabs (current and previous months only) from the
Student Attendance Google Sheet and exports aggregated records to a separate
Google Sheet named 'wp-attendance-processed'.
"""

import re
import warnings
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import gspread
import pandas as pd

from auth_utils import get_services
from drive_utils import find_file_id, resolve_folder_id, write_tab

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =============================================================================
# Configuration & Constants
# =============================================================================

TARGET_SPREADSHEET_ID = "19nPVbQHGXWwd8ATYKyxY_LF307aHfZEVAUKJf9R2muw"
DEFAULT_OUTPUT_FILE = "wp-attendance-processed"
OUTPUT_TAB_NAME = "Dashboard_Data"

# Tabs within the input spreadsheet to ignore
IGNORE_TABS = [
    "Dashboard_Data",
    # Add any specific tab names you want to skip:
    # "Clients",
    # "Template",
]

# Month name lookup table (1-indexed to match Python datetime months)
MONTH_MAP: Dict[str, int] = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

HEADERS = [
    "Full Date",
    "Month",
    "Tab Name",
    "Date Label",
    "Day of Week",
    "Student Name",
    "Original Full Name",
    "Cohort / Tag",
    "Status",
    "Raw Mark",
    "Is Present",
    "Is Late",
    "Is Absent",
    "Attendance Score",
]

_TEXT_MONTH_RE = re.compile(
    r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b",
    re.IGNORECASE,
)
_YEAR_4_RE = re.compile(r"\b(20\d\d)\b")
_YEAR_2_RE = re.compile(r"(?:'|\s|[-_/])(\d{2})\b")
_Y_FIRST_RE = re.compile(r"\b(20\d\d)[-_/.](0?[1-9]|1[0-2])\b")
_M_FIRST_4_RE = re.compile(r"\b(0?[1-9]|1[0-2])[-_/.](20\d\d)\b")
_M_FIRST_2_RE = re.compile(r"\b(0?[1-9]|1[0-2])[-_/.](2\d)\b")
_ISO_DATE_RE = re.compile(r"^(20\d\d)-(\d{1,2})-(\d{1,2})$")
_END_TAG_RE = re.compile(r"\s*\(([^()]+)\)\s*$")
_NICKNAME_RE = re.compile(r'^(.*?)\s*[\("]([^()"]+)[\)"]\s*(.*?)$')


# =============================================================================
# Helper Functions
# =============================================================================

def parse_tab_month_year(sheet_name: str) -> Optional[Tuple[int, Optional[int]]]:
    """
    Extracts month (1-12) and year from tab name, supporting multiple conventions:
    - "November 2026", "Nov 2026", "Nov 26", "Nov '26", "November", "Nov"
    - "2026-11", "11/2026", "11-2026", "11/26", "11-26"
    """
    clean_name = sheet_name.strip()

    # 1. Text month names (e.g., "November 2026", "Nov 26", "November")
    name_match = _TEXT_MONTH_RE.search(clean_name)
    if name_match:
        month = MONTH_MAP[name_match.group(1).lower()]

        y4_match = _YEAR_4_RE.search(clean_name)
        if y4_match:
            return month, int(y4_match.group(1))

        y2_match = _YEAR_2_RE.search(clean_name)
        if y2_match:
            return month, 2000 + int(y2_match.group(1))

        return month, None

    # 2. Numeric patterns: YYYY-MM or YYYY/MM
    y_first = _Y_FIRST_RE.search(clean_name)
    if y_first:
        return int(y_first.group(2)), int(y_first.group(1))

    # 3. Numeric patterns: MM-YYYY or MM/YYYY
    m_first_4 = _M_FIRST_4_RE.search(clean_name)
    if m_first_4:
        return int(m_first_4.group(1)), int(m_first_4.group(2))

    # 4. Numeric patterns: MM-YY or MM/YY
    m_first_2 = _M_FIRST_2_RE.search(clean_name)
    if m_first_2:
        return int(m_first_2.group(1)), 2000 + int(m_first_2.group(2))

    return None


def _get_prev_month_year(current_year: int, current_month: int) -> Tuple[int, int]:
    """Returns (year, month) for the month preceding (current_year, current_month)."""
    if current_month == 1:
        return current_year - 1, 12
    return current_year, current_month - 1


def _match_month_year(
    month: int,
    year: Optional[int],
    current_year: int,
    current_month: int,
    prev_year: int,
    prev_month: int,
) -> Dict[str, Any]:
    """Evaluates whether (month, year) matches the current or previous month."""
    if year is not None:
        if year == current_year and month == current_month:
            return {"is_match": True, "matched_year": current_year, "matched_month": current_month, "is_current": True}
        if year == prev_year and month == prev_month:
            return {"is_match": True, "matched_year": prev_year, "matched_month": prev_month, "is_current": False}
        return {"is_match": False}

    if month == current_month:
        return {"is_match": True, "matched_year": current_year, "matched_month": current_month, "is_current": True}
    if month == prev_month:
        return {"is_match": True, "matched_year": prev_year, "matched_month": prev_month, "is_current": False}
    return {"is_match": False}


def _extract_cell_month_year(val: Any) -> Tuple[Optional[int], Optional[int]]:
    """Extracts (month, year) from a row 1 date header value."""
    if isinstance(val, (datetime, date)):
        return val.month, val.year

    s = str(val).strip()
    if not s:
        return None, None

    iso_match = _ISO_DATE_RE.match(s)
    if iso_match:
        return int(iso_match.group(2)), int(iso_match.group(1))

    parts = s.split("/")
    if len(parts) >= 2 and parts[0].strip().isdigit():
        cell_month = int(parts[0].strip())
        cell_year: Optional[int] = None
        if len(parts) >= 3 and parts[2].strip().isdigit():
            y_str = parts[2].strip()
            cell_year = 2000 + int(y_str) if len(y_str) == 2 else int(y_str)
        return cell_month, cell_year

    return None, None


def is_current_or_previous_month_tab(sheet: gspread.Worksheet, now: datetime) -> Dict[str, Any]:
    """
    Checks whether a worksheet corresponds to either the current month or previous month.
    Uses tab name parsing first; falls back to inspecting row 1 date cells if tab name is generic.
    """
    sheet_name = sheet.title.strip()
    current_year = now.year
    current_month = now.month
    prev_year, prev_month = _get_prev_month_year(current_year, current_month)

    parsed = parse_tab_month_year(sheet_name)
    if parsed is not None:
        month, year = parsed
        return _match_month_year(month, year, current_year, current_month, prev_year, prev_month)

    # Fallback: If tab name is unparseable (e.g. "Sheet1"), inspect date headers in row 1
    try:
        rows = sheet.get("A1:J1")
        if rows and len(rows[0]) > 1:
            for val in rows[0][1:]:
                if not val:
                    continue
                cell_month, cell_year = _extract_cell_month_year(val)
                if cell_month is not None and 1 <= cell_month <= 12:
                    return _match_month_year(
                        cell_month, cell_year, current_year, current_month, prev_year, prev_month
                    )
    except Exception as exc:
        print(f'[-] Could not inspect row 1 for tab "{sheet_name}": {exc}')

    return {"is_match": False}


def is_classroom_presence_mark(mark: str) -> bool:
    """
    Determines whether an attendance mark represents presence in the classroom.
    "P" (or "PRESENT") marks the student's first day in class.
    Also recognizes other present marks: "L" (Late), "KS" (Kid Space), "5" (Five Keys).
    """
    if not mark:
        return False
    m = mark.strip().upper()
    return m in {"P", "PRESENT", "L", "LATE", "KS", "KID SPACE", "5"}


def is_classroom_exit_mark(mark: str) -> bool:
    """
    Determines whether an attendance mark signifies that the student has exited the classroom.
    An "X" signifies that the student has exited the classroom.
    """
    if not mark:
        return False
    m = mark.strip().upper()
    return m in {"X", "EXIT", "EXITED"} or m.startswith(("X ", "X-", "X/", "X:", "X("))


def parse_student_name_and_cohort(raw_input: Any) -> Tuple[str, str]:
    """Parses student name (resolving nicknames) and optional trailing cohort tag."""
    text = str(raw_input or "").strip()
    cohort = "General"

    end_tag_match = _END_TAG_RE.search(text)
    if end_tag_match:
        cohort = end_tag_match.group(1).strip()
        text = text[: end_tag_match.start()].strip()

    nick_match = _NICKNAME_RE.match(text)
    student_name = text
    if nick_match:
        nickname = nick_match.group(2).strip()
        last_name = nick_match.group(3).strip()
        student_name = f"{nickname} {last_name}".strip() if last_name else nickname

    return student_name, cohort


def format_attendance_date(raw_date: Any, sheet_year: int, sheet_name: str) -> Tuple[str, str]:
    """Formats a raw header date into ('YYYY-MM-DD', 'Month YYYY')."""
    if isinstance(raw_date, (datetime, date)):
        return raw_date.strftime("%Y-%m-%d"), f"{MONTH_NAMES[raw_date.month - 1]} {raw_date.year}"

    s = str(raw_date).strip()

    iso_match = _ISO_DATE_RE.match(s)
    if iso_match:
        y = int(iso_match.group(1))
        m = int(iso_match.group(2))
        d = int(iso_match.group(3))
        if 1 <= m <= 12:
            return f"{y}-{m:02d}-{d:02d}", f"{MONTH_NAMES[m - 1]} {y}"

    parts = s.split("/")
    if len(parts) >= 2 and parts[0].strip().isdigit() and parts[1].strip().isdigit():
        m_str = parts[0].strip().zfill(2)
        d_str = parts[1].strip().zfill(2)
        if len(parts) >= 3 and parts[2].strip().isdigit():
            y_part = parts[2].strip()
            y_str = f"20{y_part}" if len(y_part) == 2 else y_part
        else:
            y_str = str(sheet_year)

        formatted_date = f"{y_str}-{m_str}-{d_str}"
        month_idx = int(m_str)
        month_label = f"{MONTH_NAMES[month_idx - 1]} {y_str}" if 1 <= month_idx <= 12 else sheet_name
        return formatted_date, month_label

    return str(raw_date), sheet_name


def classify_attendance_mark(raw_mark: str) -> Tuple[str, int, int, int, float]:
    """Maps a normalized attendance mark to (status, is_present, is_late, is_absent, score)."""
    if raw_mark in {"P", "PRESENT"}:
        return "Present", 1, 0, 0, 1.0
    if raw_mark in {"L", "LATE"}:
        return "Late", 1, 1, 0, 0.5
    if raw_mark in {"KS", "KID SPACE"}:
        return "Kid Space", 1, 0, 0, 1.0
    if raw_mark == "5":
        return "Five Keys", 1, 0, 0, 1.0
    return "Absent", 0, 0, 1, 0.0


def process_attendance_tabs(matched_tabs: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, int]:
    """
    Processes chronologically sorted attendance tab dictionaries and returns
    (processed DataFrame, processed_tabs_count).
    Each entry in matched_tabs must contain: sheet_name, month_match, sheet_year, sort_key, data.
    """
    matched_tabs_sorted = sorted(matched_tabs, key=lambda t: t["sort_key"])

    records: List[List[Any]] = []
    student_in_classroom: Dict[str, bool] = {}
    processed_tabs_count = 0

    for tab_info in matched_tabs_sorted:
        sheet_name = tab_info["sheet_name"]
        data = tab_info["data"]
        month_match = tab_info["month_match"]
        sheet_year = tab_info["sheet_year"]

        month_type = "Current Month" if month_match.get("is_current") else "Previous Month"
        print(f'[+] Processing tab: "{sheet_name}" ({len(data)} rows) [{month_type}]')
        processed_tabs_count += 1

        date_row = data[0]  # Row 1: Dates (11/2, 11/3...)
        day_row = data[1]   # Row 2: Days (Mon, Tues...)
        num_cols = len(date_row)
        num_rows = len(data)

        # Check which columns have attendance recorded
        has_attendance = [False] * num_cols
        for c in range(1, num_cols):
            for r in range(4, num_rows):
                cell_val = data[r][c] if c < len(data[r]) else ""
                if str(cell_val or "").strip() != "":
                    has_attendance[c] = True
                    break

        # Process student rows (starting at row 5 / index 4)
        for r in range(4, num_rows):
            raw_student = str(data[r][0] if data[r] else "").strip()
            if not raw_student or raw_student.startswith("(") or "legend" in raw_student.lower():
                continue

            student_name, cohort = parse_student_name_and_cohort(raw_student)
            student_key = student_name.lower()

            # Check if this student has any presence or exit marks in this tab
            has_classroom_activity = False
            for c in range(1, num_cols):
                if not has_attendance[c]:
                    continue
                cell_val = data[r][c] if c < len(data[r]) else ""
                mark = str(cell_val or "").strip().upper()
                if is_classroom_presence_mark(mark) or is_classroom_exit_mark(mark):
                    has_classroom_activity = True
                    break

            # If student was not active previously and has no presence/exit marks in this tab, skip
            if not has_classroom_activity and student_in_classroom.get(student_key) is not True:
                continue

            # Start in classroom if continuing from previous month and active; otherwise wait for first P
            in_classroom = student_in_classroom.get(student_key) is True

            for c in range(1, num_cols):
                raw_date = date_row[c]
                if not raw_date or not has_attendance[c]:
                    continue

                day_name = str(day_row[c] if c < len(day_row) else "").strip()
                cell_val = data[r][c] if c < len(data[r]) else ""
                raw_mark = str(cell_val or "").strip().upper()

                # 1. Check for Exit ("X"): student exited the classroom
                if is_classroom_exit_mark(raw_mark):
                    in_classroom = False
                    # X means exited the classroom; do not count X or any subsequent days
                    break

                # 2. Check for first P (first day in class)
                if not in_classroom:
                    if is_classroom_presence_mark(raw_mark):
                        in_classroom = True
                    else:
                        # Prior to the first P: do not count time in classroom
                        continue

                # 3. Format date and month label
                formatted_date, month_label = format_attendance_date(raw_date, sheet_year, sheet_name)

                # 4. Calculate attendance status & score
                status, is_present, is_late, is_absent, score = classify_attendance_mark(raw_mark)

                records.append([
                    formatted_date,
                    month_label,
                    sheet_name,
                    str(raw_date),
                    day_name,
                    student_name,
                    raw_student,
                    cohort,
                    status,
                    raw_mark,
                    is_present,
                    is_late,
                    is_absent,
                    score,
                ])

            # Update persistent classroom state for this student
            student_in_classroom[student_key] = in_classroom

    df_attendance = pd.DataFrame(records, columns=HEADERS)
    return df_attendance, processed_tabs_count


def _get_spreadsheet_now(ss: gspread.Spreadsheet) -> datetime:
    """Returns current datetime in the spreadsheet's configured timezone if available."""
    try:
        tz_name = ss.timezone
        if tz_name:
            return datetime.now(ZoneInfo(tz_name))
    except Exception:
        pass
    return datetime.now()


# =============================================================================
# Main Processing Entry Point
# =============================================================================

def run_attendance_processing(
    input_spreadsheet_id: str = TARGET_SPREADSHEET_ID,
    output_file: str = DEFAULT_OUTPUT_FILE,
    input_folder_name: Optional[str] = None,
    output_folder_name: Optional[str] = None,
):
    # =========================================================================
    # 1. Initialize Services & Resolve Output Folder
    # =========================================================================
    drive_service, gc, _ = get_services()
    output_folder_id = resolve_folder_id(drive_service, output_folder_name, "Output")

    # =========================================================================
    # 2. Open Input Spreadsheet & Collect Matching Monthly Tabs
    # =========================================================================
    input_ss = gc.open_by_key(input_spreadsheet_id)
    print(f'>>> Opened input file: "{input_ss.title}" ({input_spreadsheet_id})')

    now = _get_spreadsheet_now(input_ss)
    current_year = now.year
    all_sheets = input_ss.worksheets()

    matched_tabs: List[Dict[str, Any]] = []
    for sheet in all_sheets:
        sheet_name = sheet.title.strip()

        if sheet_name in IGNORE_TABS:
            print(f'[-] Skipping ignored tab: "{sheet_name}"')
            continue

        month_match = is_current_or_previous_month_tab(sheet, now)
        if not month_match.get("is_match"):
            print(f'[-] Skipping tab: "{sheet_name}" (not current or previous month)')
            continue

        data = sheet.get_all_values()
        if len(data) < 5:
            print(f'[-] Skipping tab "{sheet_name}" (fewer than 5 rows)')
            continue

        sheet_year = month_match.get("matched_year") or current_year
        year_match = _YEAR_4_RE.search(sheet_name)
        if year_match:
            sheet_year = int(year_match.group(1))

        month_num = month_match.get("matched_month") or 1
        sort_key = sheet_year * 12 + (month_num - 1)

        matched_tabs.append({
            "sheet_name": sheet_name,
            "month_match": month_match,
            "sheet_year": sheet_year,
            "sort_key": sort_key,
            "data": data,
        })

    # =========================================================================
    # 3. Process Attendance Data
    # =========================================================================
    print("Processing Attendance...")
    df_attendance, processed_tabs_count = process_attendance_tabs(matched_tabs)

    # =========================================================================
    # 4. Handle Output Spreadsheet ('wp-attendance-processed')
    # =========================================================================
    processed_file = find_file_id(
        drive_service,
        output_file,
        output_folder_id,
        "application/vnd.google-apps.spreadsheet",
    )
    if processed_file:
        spreadsheet = gc.open_by_key(processed_file["id"])
        print(f"📄 Opened existing: '{output_file}'")
    else:
        file_metadata: Dict[str, Any] = {
            "name": output_file,
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }
        if output_folder_id:
            file_metadata["parents"] = [output_folder_id]
        new_sheet = drive_service.files().create(body=file_metadata, supportsAllDrives=True).execute()
        spreadsheet = gc.open_by_key(new_sheet["id"])
        print(f"📄 Created new: '{output_file}'")

    # =========================================================================
    # 5. Export Tabs & Format Output
    # =========================================================================
    write_tab(spreadsheet, OUTPUT_TAB_NAME, df_attendance)

    try:
        target_ws = spreadsheet.worksheet(OUTPUT_TAB_NAME)
        target_ws.format("1:1", {"textFormat": {"bold": True}})
        spreadsheet.batch_update({
            "requests": [
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": target_ws.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": len(HEADERS),
                        }
                    }
                }
            ]
        })
    except Exception as exc:
        print(f"⚠️ Warning: Could not apply formatting to '{OUTPUT_TAB_NAME}': {exc}")

    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Sheet1"))
    except Exception:
        pass

    print(
        f"\n🎉 Processing complete! Wrote {len(df_attendance)} rows across "
        f"{processed_tabs_count} tab(s): {spreadsheet.url}"
    )
    return spreadsheet.url


if __name__ == "__main__":
    print("🚀 Starting Attendance Processing")
    run_attendance_processing(
        input_folder_name="Apricot Report Incoming",
        output_folder_name="KPI Processed Data",
    )
