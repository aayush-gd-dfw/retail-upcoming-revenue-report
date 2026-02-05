"""
Completed + Upcoming Revenue → Google Sheet Updater (IMAP + gspread)
===================================================================

Per your instructions:
- ONLY re-use the Gmail read + Google Sheet update style from the code you shared:
  - IMAP via app password
  - read first .xlsx attachment
  - gspread service account JSON
  - upsert by matching Date column in the sheet
- Your tab name is: "february"
- Sheet columns:
  Col1 Day (ignore)
  Col2 Date (match key; do NOT update)
  Col3 Expected Revenue (ignore)
  Col4 Completed Revenue (update from Completed email)
  Col5 Scheduled Revenue (update from Upcoming email)
  Col6 Revenue Needed (ignore)

Email subjects (contain-match):
- "Completed Revenue for Retail Excel Dashboard"
- "Upcoming Revenue for Retail Excel Dashboard"

Logic:
1) Upcoming: sum(Subtotal) by date for dates AFTER "today" date found in the attachment,
   then overwrite Col5 (Scheduled Revenue) for matching dates in the sheet.
2) Completed: sum(Subtotal) for "today" date found in the attachment,
   set Col4 (Completed Revenue) and clear Col5 for that date.

Environment variables required:
- GMAIL_ADDRESS
- GMAIL_APP_PW
- GSPREAD_SHEET_ID
- GDRIVE_SA_JSON  (either JSON string OR file path to service account json)

Install:
pip install gspread google-auth openpyxl
"""

import os, imaplib, email, io, json, re
from email.header import decode_header, make_header
from datetime import datetime, date, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook


# -------------------- CONFIG --------------------
TAB_NAME = "February"

SUBJECT_COMPLETED_PHRASE = "Completed Revenue for Retail Excel Dashboard"
SUBJECT_UPCOMING_PHRASE  = "Upcoming Revenue for Retail Excel Dashboard"

# Your sheet columns (1-based)
COL_DATE      = 2
COL_COMPLETED = 4
COL_SCHEDULED = 5

GMAIL_ADDRESS    = "apatilglassdoctordfw@gmail.com"
GMAIL_APP_PW     = "mird noii arle cnxb"
GSPREAD_SHEET_ID = "1GA6ug2EfshOdv-NVULcItZ0K8IJtOpUe75GgClm1lpk"
GDRIVE_SA_JSON   = r"C:\Users\Aayush Patil\Downloads\sheetsautomation-476714-3aa48c47b97e.json"

# How far back to search in inbox (most recent match is used)
# IMAP subject search is not as flexible as Gmail search; we just take the latest hit.
# ------------------------------------------------
# Base columns (your original 6-col table) keep as-is
COL_DATE      = 2   # B
COL_COMPLETED = 4   # D
COL_SCHEDULED = 5   # E

# -------------------- BU column mappings --------------------
# Each BU has its own Date col, Completed col, Scheduled col in the same tab.
BU_MAP = {
    "arlington":   {"date_col": 9,  "completed_col": 11, "scheduled_col": 12},  # I, K, L
    "carrollton":  {"date_col": 16, "completed_col": 18, "scheduled_col": 19},  # P, R, S
    "colleyville": {"date_col": 23, "completed_col": 25, "scheduled_col": 26},  # W, Y, Z
    "dallas":      {"date_col": 30, "completed_col": 32, "scheduled_col": 33},  # AD, AF, AG
    "denton":      {"date_col": 37, "completed_col": 39, "scheduled_col": 40},  # AK, AM, AN
}

# -------------------- Helpers (kept logic style) --------------------

def connect_imap():
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(GMAIL_ADDRESS, GMAIL_APP_PW)
    imap.select("INBOX")
    return imap

def get_gspread_client():
    raw = GDRIVE_SA_JSON
    if not raw:
        raise RuntimeError("Missing GDRIVE_SA_JSON env var (service account json string OR file path).")
    if os.path.exists(raw):
        with open(raw, "r", encoding="utf-8") as f:
            sa_info = json.load(f)
    else:
        sa_info = json.loads(raw)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def open_ws():
    gc = get_gspread_client()
    sh = gc.open_by_key(GSPREAD_SHEET_ID)
    return sh.worksheet(TAB_NAME)

def search_latest_matching(imap, phrase):
    typ, data = imap.search(None, f'(SUBJECT "{phrase}")')
    if typ != "OK":
        return None
    ids = data[0].split()
    return ids[-1] if ids else None

def get_first_xlsx_attachment(imap, msg_id):
    typ, data = imap.fetch(msg_id, "(RFC822)")
    if typ != "OK":
        return None, None

    msg = email.message_from_bytes(data[0][1])
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            fname = part.get_filename()
            if fname:
                fname = str(make_header(decode_header(fname)))
            if fname and fname.lower().endswith(".xlsx"):
                content = part.get_payload(decode=True)
                return fname, content

    return None, None

def read_xlsx_first_sheet(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append([("" if v is None else v) for v in r])
    return rows

def parse_money(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^0-9.\-]", "", str(val))
    return float(s) if s else 0.0

def try_parse_any_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        return None

def find_col_idx(header, target_names_lower):
    h = [str(x).strip().lower() for x in header]
    # exact match
    for i, name in enumerate(h):
        if name in target_names_lower:
            return i
    # contains match
    for i, name in enumerate(h):
        for t in target_names_lower:
            if t in name:
                return i
    return None

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def batch_update_cells(ws, updates):
    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")

def build_sheet_date_row_map(ws, date_col):
    """
    Reads a date column and returns {date: row_number}
    """
    col_vals = ws.col_values(date_col)
    mapping = {}
    for i, v in enumerate(col_vals, start=1):
        if i == 1:
            continue
        d = try_parse_any_date(v)
        if d:
            mapping[d] = i
    return mapping

# -------------------- Filename date extraction (for Completed) --------------------

def extract_date_from_filename(fname: str) -> date:
    if not fname:
        raise ValueError("Missing filename for date extraction.")
    m = re.search(r"(\d{1,4})[._-](\d{1,2})[._-](\d{1,4})", fname)
    if not m:
        raise ValueError(f"Could not extract date from filename: {fname}")

    a, b, c = m.groups()
    nums = list(map(int, (a, b, c)))

    if nums[0] > 31:  # YYYY first
        yyyy, mm, dd = nums[0], nums[1], nums[2]
    else:  # MM first
        mm, dd, yy = nums
        yyyy = 2000 + yy if yy < 100 else yy

    return date(yyyy, mm, dd)

# -------------------- Upcoming parsing (unchanged behavior + BU filter) --------------------

def subtotal_by_date_from_rows_upcoming(rows):
    """
    Upcoming: sum Jobs Subtotal by Next Appt Start Date, treat "today_in_file" as min date in file.
    """
    if not rows or len(rows) < 2:
        raise ValueError("Attachment sheet is empty or missing header/body.")

    header = rows[0]
    body = rows[1:]

    bu_col   = find_col_idx(header, {"business unit"})
    date_col = find_col_idx(header, {"next appt start date"})
    sub_col  = find_col_idx(header, {"jobs subtotal", "subtotal"})

    if bu_col is None or date_col is None or sub_col is None:
        raise ValueError(f"Could not find required columns in Upcoming. Header: {header}")

    # totals_by_bu_date = { bu: {date: sum} }
    totals_by_bu_date = {bu: {} for bu in BU_MAP.keys()}
    dates_seen = []

    for r in body:
        if len(r) <= max(bu_col, date_col, sub_col):
            continue

        bu_raw = str(r[bu_col]).strip().lower()
        if bu_raw not in totals_by_bu_date:
            continue

        d = try_parse_any_date(r[date_col])
        if not d:
            continue

        subtotal = parse_money(r[sub_col])
        m = totals_by_bu_date[bu_raw]
        m[d] = m.get(d, 0.0) + subtotal
        dates_seen.append(d)

    if not dates_seen:
        raise ValueError("No valid dates found in Upcoming attachment.")

    today_in_file = min(dates_seen)
    # round cents
    for bu in totals_by_bu_date:
        totals_by_bu_date[bu] = {k: round(v, 2) for k, v in totals_by_bu_date[bu].items()}

    return today_in_file, totals_by_bu_date

def apply_upcoming(ws, today_in_file, totals_by_bu_date):
    """
    Updates:
      - Global scheduled (Col5) based on overall totals_by_bu_date (summed across BUs) for dates > today_in_file
      - BU scheduled columns per BU_MAP for dates > today_in_file
    """
    # Global Date->Row map (col B)
    base_date_row = build_sheet_date_row_map(ws, COL_DATE)

    # BU Date->Row maps (different date cols)
    bu_date_row = {bu: build_sheet_date_row_map(ws, BU_MAP[bu]["date_col"]) for bu in BU_MAP}

    # 1) Global scheduled = sum across all BU totals for that date
    global_totals = {}
    for bu, dmap in totals_by_bu_date.items():
        for d, amt in dmap.items():
            global_totals[d] = global_totals.get(d, 0.0) + amt
    global_totals = {d: round(v, 2) for d, v in global_totals.items()}

    updates = []

    # Global scheduled updates
    for d, total in global_totals.items():
        if d <= today_in_file:
            continue
        row = base_date_row.get(d)
        if not row:
            continue
        updates.append({"range": f"{col_letter(COL_SCHEDULED)}{row}", "values": [[total]]})

    # BU scheduled updates
    for bu, dmap in totals_by_bu_date.items():
        sched_col = BU_MAP[bu]["scheduled_col"]
        row_map = bu_date_row[bu]
        for d, total in dmap.items():
            if d <= today_in_file:
                continue
            row = row_map.get(d)
            if not row:
                continue
            updates.append({"range": f"{col_letter(sched_col)}{row}", "values": [[total]]})

    batch_update_cells(ws, updates)
    return len(updates)

# -------------------- Completed parsing (YOUR RULE + BU filter) --------------------

def completed_values_from_rows_by_bu_last_jobs_subtotal(rows):
    """
    Completed:
      - For each BU in BU_MAP, take the last non-empty Jobs Subtotal value within that BU.
      - Also compute global completed = last non-empty Jobs Subtotal overall (kept simple).
    """
    if not rows or len(rows) < 2:
        raise ValueError("Attachment sheet is empty or missing header/body.")

    header = rows[0]
    body = rows[1:]

    bu_col  = find_col_idx(header, {"business unit"})
    sub_col = find_col_idx(header, {"jobs subtotal", "subtotal"})

    if bu_col is None or sub_col is None:
        raise ValueError(f'Could not find required columns in Completed. Header: {header}')

    # global last non-empty
    global_last = 0.0
    for r in reversed(body):
        if len(r) <= sub_col:
            continue
        v = r[sub_col]
        if v not in (None, "", " "):
            global_last = round(parse_money(v), 2)
            break

    # BU last non-empty (scan from bottom, first match per BU wins)
    bu_last = {bu: None for bu in BU_MAP.keys()}
    remaining = set(BU_MAP.keys())

    for r in reversed(body):
        if not remaining:
            break
        if len(r) <= max(bu_col, sub_col):
            continue

        bu_raw = str(r[bu_col]).strip().lower()
        if bu_raw not in remaining:
            continue

        v = r[sub_col]
        if v in (None, "", " "):
            continue

        bu_last[bu_raw] = round(parse_money(v), 2)
        remaining.remove(bu_raw)

    # fill missing with 0.0
    for bu in bu_last:
        if bu_last[bu] is None:
            bu_last[bu] = 0.0

    return global_last, bu_last

def apply_completed_filename_date(ws, file_date, global_completed_value, bu_completed_values):
    """
    Updates:
      - Global completed (Col4) for file_date in base Date col (B)
      - Clears global scheduled (Col5) for file_date
      - Updates BU completed columns per BU_MAP for file_date based on BU date cols
      - Clears BU scheduled columns for that date too (same rule as global)
    """
    updates = []

    # Base row map
    base_date_row = build_sheet_date_row_map(ws, COL_DATE)
    base_row = base_date_row.get(file_date)
    if base_row:
        updates.append({"range": f"{col_letter(COL_COMPLETED)}{base_row}", "values": [[global_completed_value]]})
        updates.append({"range": f"{col_letter(COL_SCHEDULED)}{base_row}", "values": [[""]]})

    # BU-specific rows
    for bu, cols in BU_MAP.items():
        row_map = build_sheet_date_row_map(ws, cols["date_col"])
        row = row_map.get(file_date)
        if not row:
            continue

        updates.append({"range": f"{col_letter(cols['completed_col'])}{row}", "values": [[bu_completed_values.get(bu, 0.0)]]})
        updates.append({"range": f"{col_letter(cols['scheduled_col'])}{row}", "values": [[""]]})

    batch_update_cells(ws, updates)
    return len(updates)

# -------------------- Main --------------------

def main():
    if not all([GMAIL_ADDRESS, GMAIL_APP_PW, GSPREAD_SHEET_ID, GDRIVE_SA_JSON]):
        raise RuntimeError("Missing env vars: GMAIL_ADDRESS, GMAIL_APP_PW, GSPREAD_SHEET_ID, GDRIVE_SA_JSON")

    imap = connect_imap()
    try:
        ws = open_ws()

        # 1) Upcoming (now also writes BU scheduled cols)
        up_id = search_latest_matching(imap, SUBJECT_UPCOMING_PHRASE)
        if up_id:
            up_fname, up_content = get_first_xlsx_attachment(imap, up_id)
            if up_content:
                up_rows = read_xlsx_first_sheet(up_content)
                up_today, totals_by_bu_date = subtotal_by_date_from_rows_upcoming(up_rows)
                n = apply_upcoming(ws, up_today, totals_by_bu_date)
                print(f"[Upcoming] {up_fname} | file_today={up_today} | total updated cells={n}")
            else:
                print("[Upcoming] Email found but no .xlsx attachment.")
        else:
            print("[Upcoming] No matching email found.")

        # 2) Completed (date from filename, value from last Jobs Subtotal; also writes BU completed cols)
        c_id = search_latest_matching(imap, SUBJECT_COMPLETED_PHRASE)
        if c_id:
            c_fname, c_content = get_first_xlsx_attachment(imap, c_id)
            if c_content:
                c_rows = read_xlsx_first_sheet(c_content)
                file_date = extract_date_from_filename(c_fname)

                global_last, bu_last = completed_values_from_rows_by_bu_last_jobs_subtotal(c_rows)
                n = apply_completed_filename_date(ws, file_date, global_last, bu_last)

                print(f"[Completed] {c_fname} | file_date={file_date} | global_completed={global_last} | updated cells={n}")
            else:
                print("[Completed] Email found but no .xlsx attachment.")
        else:
            print("[Completed] No matching email found.")

        print("Done.")

    finally:
        try:
            imap.close()
        except Exception:
            pass
        imap.logout()

if __name__ == "__main__":
    main()