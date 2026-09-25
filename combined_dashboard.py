"""
Production Ahead/Behind Dashboard — Flask API
Reads FG + WIP data from SQL Server, capacity from Excel attendance files.
"""

import os
import json
import time
from datetime import datetime, timedelta, date
from flask import Flask, jsonify, send_from_directory, request, render_template
try:
    import pyodbc
except ImportError:
    pyodbc = None  # SQL (Ahead/Behind) degrades to errors; the station report still works
import openpyxl

app = Flask(__name__, template_folder="templates", static_folder="static")

# =============================================================================
# CONFIGURATION — Edit these to match your environment
# =============================================================================

# SQL Server connection string — fill in your credentials
SQL_CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=VTN1PRDSQL002;"          # ← your server
    "DATABASE=LOCALNPV;"         # ← your database
    "UID=svcsqllocal;"          # ← your SQL login
    "PWD=KetnoilocalApp;"      # ← your SQL password
)

# Excel attendance files — point to the actual .xlsm files from SharePoint/local
# The server auto-detects which sheet to read based on the target date
CAPACITY_FILES = {
    "AD": r"C:\Users\Nguyen.Nguyen\OneDrive - Northstar Precision Vietnam\Northstar Precision Vietnam - Danh sach nhan su cac khu vuc\Attendance Matrix - Assembly V02.xlsm",   # ← path to Assembly file
    "PK": r"C:\Users\Nguyen.Nguyen\OneDrive - Northstar Precision Vietnam\Northstar Precision Vietnam - Danh sach nhan su cac khu vuc\Attendance Matrix - Packaging V02.xlsm",  # ← path to Packaging file
    "RP": r"C:\Users\Nguyen.Nguyen\OneDrive - Northstar Precision Vietnam\Northstar Precision Vietnam - Danh sach nhan su cac khu vuc\Attendance Matrix - Raw Part V02.xlsm",   # ← path to Raw Part file
}

INDIRECT_LABOR_IDS = {
    # Assembly GL/TL
    "NPV18019027",
    "NPV18021002",
    # Packaging GL/TL
    "NPV18018023",
    "NPV18021011",
    # Raw Part GL/TL
    "NPV18019028",
}

# =============================================================================
# STATION STATUS REPORT — reads the per-area JSON the control boards publish
# =============================================================================
# Same shared folder the control boards write to (app.py DASHBOARD_STATUS_DIR).
DASHBOARD_STATUS_DIR = r"\\npvshare\Data\06_Operation\02.Production\02.Component\06. Planning\dashboard_status"
STATION_AREAS = ['PK', 'AD', 'RP']
STATION_AREA_LABELS = {'PK': 'PACKAGING', 'AD': 'ASSEMBLY', 'RP': 'RAW PART'}
STATION_STALE_AFTER_SEC = 30

# Grace between an order hitting its ETC and being shown as OVERDUE — the
# warehouse pulls and scans FG in some minutes after the operator actually
# finishes the box. The authoritative value is published by each board in its
# {AREA}_status.json (app.py OVERDUE_GRACE_MIN); this is only the fallback for a
# board still running an older build that doesn't publish it.
DEFAULT_OVERDUE_GRACE_MIN = 30

def read_station_area(area):
    """Read one area's status JSON and reduce it to what the station report
    shows. Never raises — a missing/locked/half-written file degrades to an
    'unavailable' area. (Ported from shopfloor_dashboard.py.)"""
    path = os.path.join(DASHBOARD_STATUS_DIR, f'{area}_status.json')
    base = {'area': area, 'label': STATION_AREA_LABELS.get(area, area),
            'available': False, 'stale': True, 'age': None, 'ts': None,
            'file_area': None, 'mismatch': False, 'shift_windows': None,
            'overdue_grace_min': DEFAULT_OVERDUE_GRACE_MIN, 'stations': []}
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        return base
    age = time.time() - mtime
    base['age'] = round(age)
    try:
        with open(path, encoding='utf-8') as f:
            snap = json.load(f)
    except Exception:
        return base
    board = snap.get('board', {}) or {}
    stations = []
    # Collect EVERY order in the area (in-progress, queued at a station, and the
    # unassigned pool) with its ship date / days-left / remaining, so the at-risk
    # banner can cover all of them — not just what's running. Time-gated in
    # api_status. 'where' is the human location; 'station' is the card to pulse
    # (None for unassigned).
    risk_raw = []
    def _consider(p, where, station):
        if not p:
            return
        risk_raw.append({
            'pdo': p.get('id'),
            'part': p.get('part') or p.get('desc'),
            'ship_date': p.get('ship_date'),
            'days_left': p.get('days_left'),
            'remaining': p.get('remaining_qty'),
            'where': where,
            'station': station,
        })
    for sid in (snap.get('stations') or []):
        st = board.get(sid, {}) or {}
        ip = st.get('inprogress') or None
        timer = st.get('timer', {}) or {}
        workers = st.get('workers', []) or []
        stations.append({
            'id': sid,
            'pdo': ip.get('id') if ip else None,
            'part': ip.get('part') if ip else None,
            'completed': ip.get('completed_qty') if ip else None,
            'qty': ip.get('qty') if ip else None,
            'etc': timer.get('etc'),
            'paused': bool(timer.get('paused')),
            'ops': len(workers),
        })
        _consider(ip, sid, sid)
        for q in (st.get('queue') or []):
            _consider(q, f'{sid} · queue', sid)
    for q in (snap.get('queue') or []):   # unassigned pool
        _consider(q, 'unassigned', None)
    file_area = snap.get('area')
    return {
        'area': area,
        'label': STATION_AREA_LABELS.get(area, area),
        'file_area': file_area,
        'mismatch': bool(file_area) and file_area != area,
        'available': True,
        'stale': age > STATION_STALE_AFTER_SEC,
        'age': round(age),
        'ts': snap.get('ts'),
        'shift_windows': snap.get('shift_windows'),
        # Per-area, not global: a board still on an older build won't publish it,
        # and its stations should keep the fallback rather than borrow another
        # area's number.
        'overdue_grace_min': snap.get('overdue_grace_min', DEFAULT_OVERDUE_GRACE_MIN),
        'stations': stations,
        'risk_raw': risk_raw,
    }


# Shift parameters
BASE_HOURS_PER_SHIFT = 6.733  # Net work hours per shift (no OT) — Ca 1 / Ca 2
# Formula: (490min total - 50min breaks - 10min 5S) × (1 - 6% PFD) = 404min = 6.73h
# HC no OT = 6.90h (different breaks), handled separately in parser
BASE_HOURS_HC = 6.90          # Net work hours — Hành Chính, no OT
TARGET_EFFICIENCY = 0.95          # 95%

# Shift definitions: start/end in decimal hours, with per-shift break windows
# Each shift has its own break schedule so the pace line is accurate
SHIFTS = {
    "Ca 1": {
        "start": 6.0,     # 6:00 AM
        "end": 14.333,     # 14:20
        "breaks_no_ot": [
            (8.0, 8.167),      # 8:00-8:10   (10 min)
            (10.5, 11.0),      # 10:30-11:00 (30 min)
            (12.167, 12.333),   # 12:10-12:20 (10 min)
        ],
        "breaks_ot2": [
            (8.0, 8.167),      # 8:00-8:10
            (10.5, 11.0),      # 10:30-11:00
            (12.167, 12.333),   # 12:10-12:20
            (14.167, 14.333),   # 14:10-14:20
        ],
        "breaks_ot3": [
            (8.0, 8.167),      # 8:00-8:10
            (10.5, 11.167),     # 10:30-11:10 (40 min)
            (14.167, 14.667),   # 14:10-14:40 (30 min)
        ],
        "ot2_end": 16.333,  # 16:20
        "ot3_end": 17.333,  # 17:20
    },
    "Ca 2": {
        "start": 14.167,   # 14:10
        "end": 22.333,     # 22:20
        "breaks_no_ot": [
            (16.167, 16.333),   # 16:10-16:20 (10 min)
            (19.0, 19.5),       # 19:00-19:30 (30 min)
            (20.5, 20.667),     # 20:30-20:40 (10 min)
        ],
        "breaks_ot2": [
            (14.167, 14.333),   # 14:10-14:20
            (16.167, 16.333),   # 16:10-16:20
            (19.0, 19.5),       # 19:00-19:30
            (20.5, 20.667),     # 20:30-20:40
        ],
        "breaks_ot3": [
            (14.167, 14.333),   # 14:10-14:20
            (14.167, 14.667),   # 14:10-14:40 (30 min)
            (17.167, 17.333),   # 17:10-17:20
            (19.0, 19.5),       # 19:00-19:30
            (20.5, 20.667),     # 20:30-20:40
        ],
        "ot2_end": 22.333,  # same end, starts 2h earlier at 12:10
        "ot3_end": 22.333,  # same end, starts 3h earlier at 11:10
        "ot2_start": 12.167,
        "ot3_start": 11.167,
    },
    "HC": {
        "start": 8.0,      # 8:00 AM
        "end": 16.5,       # 16:30
        "breaks_no_ot": [
            (10.0, 10.167),     # 10:00-10:10 (10 min)
            (12.0, 12.667),     # 12:00-12:40 (40 min)
            (15.0, 15.167),     # 15:00-15:10 (10 min)
        ],
        "breaks_ot2": [
            (10.0, 10.167),     # 10:00-10:10
            (11.0, 11.167),     # 11:00-11:10 (10 min — from OT break image)
            (12.0, 12.667),     # 12:00-12:40
            (15.0, 15.167),     # 15:00-15:10
        ],
        "breaks_ot3": [
            (10.0, 10.167),
            (12.0, 12.667),
            (15.0, 15.167),
            (16.5, 16.667),     # 16:30-16:40
        ],
        "ot2_end": 18.5,    # 18:30
        "ot3_end": 19.5,    # 19:30
    },
}

# Overall day boundaries (for chart x-axis)
DAY_START = 6.0    # earliest shift start
DAY_END = 22.5     # 10:30 PM — chart extends to here

# Shift name mappings — the Ca column in Excel may use various labels
SHIFT_NAME_MAP = {
    "Ca 1": "Ca 1",
    "Ca 2": "Ca 2",
    "Ca Hành Chính": "HC",
    "HC": "HC",
    "Hành Chính": "HC",
}

# Status codes that mean "present/working"
PRESENT_CODES = {"ĐL", "ĐC", "HC"}

# Status codes that mean "absent"
ABSENT_CODES = {"NP(Y)", "NP(N)", "NO", "NKLD", "NVR(Y)", "NVR(N)",
                "NP 1/2(Y)", "NP 1/2(N)", "NVR 1/2(Y)", "NVR 1/2(N)",
                "NĐB"}


# =============================================================================
# CAPACITY PARSING — Reads Excel attendance files
# =============================================================================

def find_sheet_for_date(wb, target_date):
    """
    Auto-detect which sheet contains data for target_date.
    Strategy:
    1. Look for sheets with T{month}-{year} pattern matching target_date
    2. Among matches, prefer the one whose header row actually has the target date
    3. Fall back to any sheet that has the target date in its header row
    """
    import re
    target_ym = (target_date.year, target_date.month)

    def extract_ym(sheet_name):
        m = re.search(r'T(\d{1,2})-(\d{4})', sheet_name)
        if m:
            return (int(m.group(2)), int(m.group(1)))
        return None

    # First pass: sheets whose name matches target year/month
    name_matches = []
    for sname in wb.sheetnames:
        ym = extract_ym(sname)
        if ym == target_ym:
            name_matches.append(sname)

    # Among name matches, find one that has the target date in header
    for sname in reversed(name_matches):  # reversed = prefer latest (e.g. v2, suffix)
        ws = wb[sname]
        rows = list(ws.iter_rows(max_row=2, values_only=True))
        if len(rows) > 1:
            if any(isinstance(v, datetime) and v.date() == target_date for v in rows[1]):
                return sname, ws, rows

    # Second pass: any sheet with the target date in header row
    for sname in reversed(wb.sheetnames):
        ws = wb[sname]
        rows = list(ws.iter_rows(max_row=2, values_only=True))
        if len(rows) > 1:
            if any(isinstance(v, datetime) and v.date() == target_date for v in rows[1]):
                return sname, ws, rows

    return None, None, None


def parse_capacity_file(filepath, target_date):
    """
    Each worker has exactly 3 rows:
      Row 1: STT | MSNV | Ca | ... | daily status (ĐL/ĐC/HC/NP/etc.)
      Row 2: OT  | MSNV | Ca | ... | OT hours
      Row 3: Lend or Borrow | MSNV | ... | negative=lent out, positive=borrowed in
    
    Status: ĐL=Shift1, ĐC=Shift2, HC=General
    Absent: NP(Y/N), NO, NVR(Y/N), NĐB, NKLD → 0 hours
    Half-day: NP 1/2(Y/N), NVR 1/2(Y/N) → half base, shift from previous working day
    
    Auto-detects the correct month sheet based on target_date.
    Supports .xlsm files with multiple monthly sheets (e.g. DS Assembly T6-2026).
    """
    if not os.path.exists(filepath):
        return {"workers": 0, "available_hours": 0, "error": f"File not found: {filepath}"}

    # Copy to temp file to avoid OneDrive/SharePoint file locking issues
    import shutil, tempfile
    tmp_path = None
    try:
        suffix = os.path.splitext(filepath)[1]
        tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        tmp.close()
        tmp_path = tmp.name
        shutil.copy2(filepath, tmp_path)
        read_path = tmp_path
    except Exception as e:
        # If copy fails, try reading directly
        read_path = filepath

    try:
        wb = openpyxl.load_workbook(read_path, read_only=True, data_only=True)
    except Exception as e:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return {"workers": 0, "available_hours": 0, "error": f"Cannot open file: {e}"}

    sheet_name, ws, rows = find_sheet_for_date(wb, target_date)

    if ws is None:
        wb.close()
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return {"workers": 0, "available_hours": 0,
                "error": f"No sheet found for date {target_date} in {os.path.basename(filepath)}"}

    # Load remaining rows (find_sheet_for_date only loaded first 2)
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Clean up temp file
    if tmp_path and os.path.exists(tmp_path):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    if len(rows) < 2:
        return {"workers": 0, "available_hours": 0, "error": "File too short"}

    header = rows[1]
    date_col = None
    for col_idx, val in enumerate(header):
        if isinstance(val, datetime):
            if val.date() == target_date:
                date_col = col_idx
                break

    if date_col is None:
        return {"workers": 0, "available_hours": 0,
                "error": f"Date {target_date} not found in file columns"}

    # Build sorted list of date columns for half-day lookback
    date_columns = {}
    for col_idx, val in enumerate(header):
        if isinstance(val, datetime):
            date_columns[val.date()] = col_idx
    sorted_dates = sorted(date_columns.keys())

    data_rows = rows[2:]
    total_hours = 0
    worker_count = 0
    worker_details = []
    shift_hours = {"Ca 1": 0, "Ca 2": 0, "HC": 0}
    shift_workers = {"Ca 1": 0, "Ca 2": 0, "HC": 0}

    absent_codes = {"NP(Y)", "NP(N)", "NO", "NVR(Y)", "NVR(N)", "NĐB", "NKLD"}
    half_day_codes = {"NP 1/2(Y)", "NP 1/2(N)", "NVR 1/2(Y)", "NVR 1/2(N)"}

    i = 0
    while i + 2 < len(data_rows):
        main_row = data_rows[i]
        row2 = data_rows[i + 1]
        row3 = data_rows[i + 2]

        # Verify worker block: first cell is integer STT
        if not isinstance(main_row[0], (int, float)) or main_row[0] != int(main_row[0]):
            i += 1
            continue

        # Skip indirect labor (GL, TL, supervisors)
        msnv = str(main_row[2]).strip() if main_row[2] else ""
        if msnv in INDIRECT_LABOR_IDS:
            i += 3
            continue

        # Detect row order by checking labels in row 2 and row 3
        # Format A (regular): row2=OT, row3=Lend
        # Format B (borrowed): row2=Borrow, row3=OT
        row2_label = str(row2[0]).strip().lower() if row2[0] else ""
        row3_label = str(row3[0]).strip().lower() if row3[0] else ""

        if row2_label == "borrow":
            # Format B: row2=Borrow, row3=OT
            ot_row = row3
            third_row = row2  # Borrow row
        else:
            # Format A: row2=OT, row3=Lend (default)
            ot_row = row2
            third_row = row3

        # Get today's status
        status = main_row[date_col] if date_col < len(main_row) else None
        status_str = str(status).strip() if status else ""

        # Determine shift
        actual_shift = None
        is_half_day = status_str in half_day_codes

        if status_str == "ĐL":
            actual_shift = "Ca 1"
        elif status_str == "ĐC":
            actual_shift = "Ca 2"
        elif status_str == "HC":
            actual_shift = "HC"
        elif is_half_day:
            # Look back to find last working day's shift
            actual_shift = "Ca 1"  # fallback
            for prev_date in reversed(sorted_dates):
                if prev_date >= target_date:
                    continue
                prev_col = date_columns[prev_date]
                prev_status = main_row[prev_col] if prev_col < len(main_row) else None
                prev_str = str(prev_status).strip() if prev_status else ""
                if prev_str == "ĐL":
                    actual_shift = "Ca 1"
                    break
                elif prev_str == "ĐC":
                    actual_shift = "Ca 2"
                    break
                elif prev_str == "HC":
                    actual_shift = "HC"
                    break
        elif status_str in absent_codes or status_str in ("", "None"):
            i += 3
            continue
        else:
            # Unknown status code — skip
            i += 3
            continue

        # ── Hours calculation ──
        BASE_CLOCK = 8.0
        is_borrowed = (row2_label == "borrow")

        # Lend/Borrow clock hours from third_row
        lend_borrow_clock = 0
        if date_col < len(third_row):
            val = third_row[date_col]
            if isinstance(val, (int, float)) and val != 0:
                lend_borrow_clock = float(val)

        # OT clock hours from ot_row
        ot_clock = 0
        if date_col < len(ot_row):
            ot_val = ot_row[date_col]
            if isinstance(ot_val, (int, float)) and ot_val > 0:
                ot_clock = float(ot_val)

        # Weekend detection: if OT >= 8, the entire day is OT (Sat/Sun work)
        # No base shift — OT value IS their total clock hours
        is_weekend_ot = (ot_clock >= 8)

        if is_borrowed:
            # Borrowed worker: borrow value IS their base clock hours in this area
            # OT is additional clock hours on top
            remaining_base_clock = abs(lend_borrow_clock)
            remaining_ot_clock = ot_clock if not is_weekend_ot else 0
            if is_weekend_ot:
                # Weekend borrowed: borrow value is their clock hours, no separate OT
                remaining_base_clock = abs(lend_borrow_clock)
                remaining_ot_clock = 0
        elif is_weekend_ot:
            # Weekend: OT value is the total clock hours (replaces base)
            total_weekend_clock = ot_clock
            if lend_borrow_clock < 0:
                total_weekend_clock = max(0, ot_clock + lend_borrow_clock)
            elif lend_borrow_clock > 0:
                total_weekend_clock = ot_clock + lend_borrow_clock
            # Treat entire time as "base" for productive conversion
            remaining_base_clock = total_weekend_clock
            remaining_ot_clock = 0
        else:
            # Regular weekday worker
            base_clock = BASE_CLOCK if not is_half_day else BASE_CLOCK / 2

            if lend_borrow_clock == 0:
                remaining_base_clock = base_clock
                remaining_ot_clock = ot_clock
            else:
                # Lent out: subtract from base first, overflow to OT
                abs_lb = abs(lend_borrow_clock)
                base_lent = min(abs_lb, base_clock)
                ot_lent = max(0, abs_lb - base_clock)
                remaining_base_clock = base_clock - base_lent
                remaining_ot_clock = max(0, ot_clock - ot_lent)

        # Convert remaining clock hours to net productive hours
        # Values from shift capacity summary (includes 5S + 6% PFD deduction):
        #   Base shift (Ca 1/Ca 2): 6.73h no OT | HC: 6.90h no OT
        #   OT 2h: 8.47h total (all shifts same)
        #   OT 3h: 9.08h total (Ca 1/Ca 2 only)
        BASE_CLOCK = 8.0
        OT2_NET = 8.47   # net hours for base + 2h OT
        OT3_NET = 9.08   # net hours for base + 3h OT

        # Base productive (no OT portion)
        base_net = BASE_HOURS_HC if actual_shift == "HC" else BASE_HOURS_PER_SHIFT
        if is_half_day:
            base_net = base_net / 2

        base_productive = remaining_base_clock * (base_net / BASE_CLOCK)

        # OT productive: use delta from total net (OT2_NET - base, OT3_NET - base)
        ot_productive = 0
        if remaining_ot_clock > 0:
            base_no_ot = BASE_HOURS_HC if actual_shift == "HC" else BASE_HOURS_PER_SHIFT
            if remaining_ot_clock <= 2:
                # OT 2h total net = 8.47h, so OT portion = 8.47 - base
                ot_net_total = OT2_NET
                ot_productive = (ot_net_total - base_no_ot) * (remaining_ot_clock / 2.0)
            else:
                # OT 3h total net = 9.08h, so OT portion = 9.08 - base
                ot_net_total = OT3_NET
                ot_productive = (ot_net_total - base_no_ot) * (remaining_ot_clock / 3.0)
            ot_productive = max(0, ot_productive)

        worker_hours = max(0, base_productive + ot_productive)
        total_hours += worker_hours
        worker_count += 1
        shift_hours[actual_shift] += worker_hours
        shift_workers[actual_shift] += 1

        third_label = str(third_row[0]).strip() if third_row[0] else ""
        worker_details.append({
            "msnv": str(main_row[2]),   # col 2 = MSNV (new format: STT|Name|MSNV|Ca)
            "name": str(main_row[1]),   # col 1 = Full name
            "shift": actual_shift,
            "base": round(base_productive, 2),
            "ot": round(ot_productive, 2),
            "lend_borrow_clock": round(lend_borrow_clock, 2),
            "third_type": third_label,
            "total": round(worker_hours, 2),
        })

        i += 3

    return {
        "workers": worker_count,
        "available_hours": round(total_hours, 2),
        "target_hours": round(total_hours * TARGET_EFFICIENCY, 2),
        "shift_hours": {k: round(v, 2) for k, v in shift_hours.items()},
        "shift_workers": shift_workers,
        "details": worker_details,
    }


# Cache: capacity parsed from Excel, refreshed every 60 minutes
_capacity_cache = {}  # key: date_str -> {"data": ..., "built_at": datetime}
CAPACITY_CACHE_TTL = 60 * 60  # 1 hour in seconds

def get_all_capacity(target_date):
    """Parse all area files and return combined + per-area capacity, with shift breakdown.
    Results are cached for 1 hour to reduce Excel file reads, while still picking up
    same-day attendance updates from GL."""
    cache_key = target_date.isoformat()
    now = datetime.now()

    # Check cache
    if cache_key in _capacity_cache:
        cached = _capacity_cache[cache_key]
        age_seconds = (now - cached["built_at"]).total_seconds()
        if age_seconds < CAPACITY_CACHE_TTL:
            return cached["data"]

    # Build fresh
    result = {
        "areas": {},
        "total_available": 0,
        "total_target": 0,
        "total_workers": 0,
        "total_shift_hours": {"Ca 1": 0, "Ca 2": 0, "HC": 0},
        "total_shift_workers": {"Ca 1": 0, "Ca 2": 0, "HC": 0},
    }

    for area, filepath in CAPACITY_FILES.items():
        cap = parse_capacity_file(filepath, target_date)
        result["areas"][area] = cap
        if "error" not in cap:
            result["total_available"] += cap["available_hours"]
            result["total_target"] += cap["target_hours"]
            result["total_workers"] += cap["workers"]
            for s in ["Ca 1", "Ca 2", "HC"]:
                result["total_shift_hours"][s] += cap["shift_hours"].get(s, 0)
                result["total_shift_workers"][s] += cap["shift_workers"].get(s, 0)

    result["total_available"] = round(result["total_available"], 2)
    result["total_target"] = round(result["total_target"], 2)
    result["total_shift_hours"] = {k: round(v, 2) for k, v in result["total_shift_hours"].items()}

    # Store in cache
    _capacity_cache[cache_key] = {"data": result, "built_at": now}
    print(f"[Capacity] Refreshed for {target_date} — {result['total_workers']} workers, {result['total_available']}h available")
    return result


# =============================================================================
# SQL QUERIES — FG + WIP earned hours
# =============================================================================

def get_earned_hours(target_date):
    """
    Query SQL for today's earned hours:
      - FG_Database_All: completed boxes
      - Nhaplecuoingay_All: WIP at end of shift
      - Deduct yesterday's WIP for LPs that reappear today
    
    Returns intraday timeline (per FG scan) + totals per area.
    """
    today_str = target_date.strftime("%Y-%m-%d")

    try:
        conn = pyodbc.connect(SQL_CONN_STR, timeout=10)
        cursor = conn.cursor()

        # --- FG completed boxes today ---
        cursor.execute("""
            SELECT [Input time], [Earned hour], Station, PO, [License Plate], [Part no]
            FROM [dbo].[FG_Database_All]
            WHERE CAST([Production date] AS DATE) = ?
            ORDER BY [Input time]
        """, today_str)
        fg_rows = cursor.fetchall()

        # --- WIP logged today (Nhaplecuoingay for today) ---
        # These are boxes still in progress, logged at end of shift.
        # Their earned hours count toward today's total and appear on the chart.
        cursor.execute("""
            SELECT [Input time], CAST([Earned hour] AS float) AS earned, Station, [License Plate]
            FROM [dbo].[Nhaplecuoingay_All]
            WHERE CAST([Production date] AS DATE) = ?
            ORDER BY [Input time]
        """, today_str)
        wip_today_rows = cursor.fetchall()

        # --- Latest WIP entry per LP within 30 days (for deduction) ---
        # Only deduct the most recent Nhaplecuoingay record for each LP,
        # not the sum of all entries. The latest entry represents the
        # most recent WIP state for that box.
        cursor.execute("""
            SELECT w.[License Plate], CAST(w.[Earned hour] AS float) AS earned
            FROM [dbo].[Nhaplecuoingay_All] w
            INNER JOIN (
                SELECT [License Plate], MAX([Production date]) AS max_date
                FROM [dbo].[Nhaplecuoingay_All]
                WHERE CAST([Production date] AS DATE) < ?
                  AND CAST([Production date] AS DATE) >= DATEADD(day, -30, ?)
                GROUP BY [License Plate]
            ) latest ON w.[License Plate] = latest.[License Plate]
                    AND w.[Production date] = latest.max_date
        """, today_str, today_str)
        wip_prior_rows = cursor.fetchall()

        conn.close()

    except Exception as e:
        return {"error": str(e), "fg": [], "wip": [], "timeline": [],
                "totals": {"fg_hours": 0, "wip_hours": 0, "deduct_hours": 0, "net_hours": 0}}

    # Build prior WIP lookup: LP -> earned hours from latest Nhaplecuoingay entry
    prior_wip = {}
    for row in wip_prior_rows:
        lp = str(row[0]).strip()
        hours = float(row[1]) if row[1] else 0
        prior_wip[lp] = hours

    # Today's FG License Plates
    today_lps = set()
    for row in fg_rows:
        today_lps.add(str(row[4]).strip())

    # Deduction: sum of all prior WIP hours for LPs that reappear today
    # This ensures a box that was WIP across Mon+Tue gets both days deducted
    # when it appears as FG on Thursday
    deduct_hours = 0
    for lp, hours in prior_wip.items():
        if lp in today_lps:
            deduct_hours += hours

    # Collect all entries (FG + WIP) into one list, then sort and compute cumulative
    raw_entries = []  # list of (minutes_since_6am, time_str, net_earned, station, area)

    deducted_lps = set()

    # FG entries
    for row in fg_rows:
        input_time = row[0]
        earned = float(row[1]) if row[1] else 0
        station = str(row[2]).strip() if row[2] else ""
        lp = str(row[4]).strip() if row[4] else ""

        net_earned = earned
        if lp in prior_wip and lp not in deducted_lps:
            net_earned = earned - prior_wip[lp]
            deducted_lps.add(lp)

        time_str = ""
        minutes_since_6am = 0
        if isinstance(input_time, datetime):
            time_str = input_time.strftime("%H:%M:%S")
            minutes_since_6am = (input_time.hour - 6) * 60 + input_time.minute

        area = "AD"
        if station.upper().startswith("PK"):
            area = "PK"
        elif station.upper().startswith("RP"):
            area = "RP"

        part_no = str(row[5]).strip() if row[5] else ""
        raw_entries.append((minutes_since_6am, time_str, net_earned, station, area, "FG", part_no))

    # WIP entries (today's Nhaplecuoingay)
    # Query columns: [Input time], earned, Station, [License Plate]
    wip_total = 0
    wip_area = {"AD": 0, "PK": 0, "RP": 0}
    for row in wip_today_rows:
        input_time = row[0]
        earned = float(row[1]) if row[1] else 0
        station = str(row[2]).strip() if row[2] else ""
        lp = str(row[3]).strip() if row[3] else ""

        if lp in today_lps:
            continue

        net_earned = earned
        if lp in prior_wip:
            net_earned = earned - prior_wip[lp]
        net_earned = max(0, net_earned)

        if net_earned <= 0:
            continue

        wip_total += net_earned

        area = "AD"
        if station.upper().startswith("PK"):
            area = "PK"
        elif station.upper().startswith("RP"):
            area = "RP"
        wip_area[area] += net_earned

        time_str = ""
        minutes_since_6am = 0
        if isinstance(input_time, datetime):
            time_str = input_time.strftime("%H:%M:%S")
            minutes_since_6am = (input_time.hour - 6) * 60 + input_time.minute

        raw_entries.append((minutes_since_6am, time_str, net_earned, station, area, "WIP", ""))

    # Sort all entries by time
    raw_entries.sort(key=lambda e: e[0])

    # Build timeline with running cumulative
    timeline = []
    area_totals = {"AD": 0, "PK": 0, "RP": 0}
    cumulative = 0

    for minutes, time_str, net_earned, station, area, source, part_no in raw_entries:
        cumulative += net_earned
        area_totals[area] += net_earned

        timeline.append({
            "time": time_str,
            "minutes": minutes,
            "earned": round(net_earned, 2),
            "cumulative": round(cumulative, 2),
            "station": station,
            "area": area,
            "part": part_no,
        })

    raw_fg_total = sum(float(r[1]) for r in fg_rows if r[1])
    net_fg_total = cumulative - wip_total
    net_total = cumulative

    return {
        "timeline": timeline,
        "totals": {
            "raw_fg_hours": round(raw_fg_total, 2),
            "fg_hours": round(net_fg_total, 2),
            "wip_hours": round(wip_total, 2),
            "deduct_hours": round(deduct_hours, 2),
            "net_hours": round(net_total, 2),
        },
        "area_earned": {
            area: round(area_totals[area] + wip_area[area], 2)
            for area in ["AD", "PK", "RP"]
        },
        "fg_area": {k: round(v, 2) for k, v in area_totals.items()},
        "wip_area": {k: round(v, 2) for k, v in wip_area.items()},
    }


# =============================================================================
# TARGET PACE LINE — Layered by shift type
# =============================================================================

def _productive_minutes_for_shift(shift_key):
    """Count total productive minutes for a shift (excluding its breaks)."""
    shift = SHIFTS[shift_key]
    start = shift["start"]
    end = shift["end"]
    breaks = shift["breaks_no_ot"]
    
    productive = 0
    total_minutes = int((end - start) * 60)
    for m in range(total_minutes):
        hour_dec = start + m / 60
        in_break = any(bs <= hour_dec < be for bs, be in breaks)
        if not in_break:
            productive += 1
    return productive


def build_target_pace(target_hours, shift_hours=None):
    """
    Build a minute-by-minute target pace line from 6:00 to 22:20.
    
    If shift_hours is provided (dict with "Ca 1", "Ca 2", "HC" keys),
    the pace line layers each shift's contribution during its active window,
    flattening during that shift's breaks. This means:
      - Ca 1 hours only accumulate during 6:00-14:20 (minus Ca 1 breaks)
      - HC hours only accumulate during 8:00-16:30 (minus HC breaks)
      - Ca 2 hours only accumulate during 14:10-22:20 (minus Ca 2 breaks)
    
    If shift_hours is None, falls back to even distribution across full day.
    """
    # Minutes from DAY_START (6:00) — chart x-axis
    total_chart_minutes = int((DAY_END - DAY_START) * 60)

    if shift_hours is None or all(v == 0 for v in shift_hours.values()):
        # Fallback: spread target_hours evenly across all productive minutes
        # Use combined breaks from all shifts
        all_breaks = []
        for s in SHIFTS.values():
            all_breaks.extend(s["breaks_no_ot"])
        
        productive_total = 0
        for m in range(total_chart_minutes):
            hour_dec = DAY_START + m / 60
            in_break = any(bs <= hour_dec < be for bs, be in all_breaks)
            if not in_break:
                productive_total += 1

        points = []
        work_elapsed = 0
        for m in range(total_chart_minutes + 1):
            hour_dec = DAY_START + m / 60
            in_break = any(bs <= hour_dec < be for bs, be in all_breaks)
            if not in_break and m > 0:
                work_elapsed += 1
            target_val = (work_elapsed / max(productive_total, 1)) * target_hours
            points.append({"minutes": m, "target": round(target_val, 2)})
        return points

    # --- Layered pace line ---
    # For each shift, calculate its productive minutes and rate per minute
    shift_rates = {}  # target hours per productive minute for each shift
    for shift_key in ["Ca 1", "Ca 2", "HC"]:
        shift_target = shift_hours.get(shift_key, 0) * TARGET_EFFICIENCY
        if shift_target <= 0:
            shift_rates[shift_key] = 0
            continue
        prod_mins = _productive_minutes_for_shift(shift_key)
        shift_rates[shift_key] = shift_target / max(prod_mins, 1)

    # Build pace line minute by minute
    points = []
    cumulative_target = 0

    for m in range(total_chart_minutes + 1):
        hour_dec = DAY_START + m / 60
        
        if m > 0:
            # For each shift, check if this minute is productive (active + not on break)
            for shift_key in ["Ca 1", "Ca 2", "HC"]:
                if shift_rates[shift_key] <= 0:
                    continue
                shift = SHIFTS[shift_key]
                # Is this minute within the shift's active window?
                if shift["start"] <= hour_dec < shift["end"]:
                    # Is this minute NOT in the shift's break?
                    in_break = any(bs <= hour_dec < be for bs, be in shift["breaks_no_ot"])
                    if not in_break:
                        cumulative_target += shift_rates[shift_key]

        points.append({"minutes": m, "target": round(cumulative_target, 2)})

    return points


# =============================================================================
# STATION PACE MODEL — Inter-arrival cycle time per station/part/hour bracket
# =============================================================================

# Cache: built once at startup, refreshed every hour
_station_model_cache = {"data": None, "built_at": None}

def get_hour_bracket(hour):
    """Map hour of day to shift bracket for cycle time lookup."""
    if 6 <= hour < 10:
        return "06-10"   # Ca 1 early (slow ramp)
    elif 10 <= hour < 14:
        return "10-14"   # Ca 1 peak (may have OT overlap ops)
    elif 14 <= hour < 18:
        return "14-18"   # Ca 2 early + Ca 1 OT tail
    else:
        return "18-23"   # Ca 2 peak


def build_station_cycle_model():
    """
    Query last 30 days of FG data to build per-station, per-part, per-hour-bracket
    average cycle time (minutes between consecutive scans) and average earned hours.

    Filters out cross-shift gaps (>240 min) to avoid shift-boundary contamination.
    Returns dict: {(station, part_no, bracket): {avg_minutes, avg_earned, count}}
    """
    start_date = (date.today() - timedelta(days=30)).strftime("%Y-%m-%d")

    try:
        conn = pyodbc.connect(SQL_CONN_STR, timeout=15)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT
                t.Station,
                t.[Part no],
                CASE
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 6 AND 9  THEN '06-10'
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 10 AND 13 THEN '10-14'
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 14 AND 17 THEN '14-18'
                    ELSE '18-23'
                END AS bracket,
                AVG(CAST(t.gap_minutes AS float))  AS avg_minutes,
                AVG(CAST(t.[Earned hour] AS float)) AS avg_earned,
                COUNT(*)                            AS scan_count
            FROM (
                SELECT
                    Station, [Part no], [Input time],
                    CAST([Earned hour] AS float) AS [Earned hour],
                    DATEDIFF(minute,
                        LAG([Input time]) OVER (
                            PARTITION BY Station, CAST([Production date] AS DATE)
                            ORDER BY [Input time]
                        ),
                        [Input time]
                    ) AS gap_minutes
                FROM [dbo].[FG_Database_All]
                WHERE CAST([Production date] AS DATE) >= ?
            ) t
            WHERE t.gap_minutes BETWEEN 5 AND 240
            GROUP BY t.Station, t.[Part no],
                CASE
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 6 AND 9  THEN '06-10'
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 10 AND 13 THEN '10-14'
                    WHEN DATEPART(hour, t.[Input time]) BETWEEN 14 AND 17 THEN '14-18'
                    ELSE '18-23'
                END
        """, start_date)
        rows = cursor.fetchall()
        conn.close()
    except Exception as e:
        print(f"[StationModel] Build error: {e}")
        return {}

    model = {}
    for row in rows:
        station = str(row[0]).strip()
        part = str(row[1]).strip()
        bracket = str(row[2]).strip()
        key = (station, part, bracket)
        model[key] = {
            "avg_minutes": float(row[3]),
            "avg_earned":  float(row[4]),
            "count":       int(row[5]),
        }
    print(f"[StationModel] Built {len(model)} station/part/bracket combinations")
    return model


def get_station_model():
    """Return cached station model, rebuilding if older than 1 hour or if previous build failed."""
    now = datetime.now()
    cache = _station_model_cache
    built_at = cache.get("built_at")

    # Rebuild if: never built, older than 1 hour, or last build returned empty (connection failed)
    needs_rebuild = (
        cache["data"] is None or
        len(cache["data"]) == 0 or
        built_at is None or
        (now - built_at).total_seconds() > 3600
    )
    if needs_rebuild:
        result = build_station_cycle_model()
        cache["data"] = result
        cache["built_at"] = now
        if result:
            print(f"[StationModel] Loaded {len(result)} station/part/bracket combinations")
        else:
            # Failed — clear built_at so it retries on next request (5 min cooldown)
            cache["built_at"] = now  # still set to avoid hammering SQL every request
    return cache["data"]


def estimate_station_wip(timeline, model, target_date):
    """
    Estimate WIP currently accumulating at each active station based on
    time elapsed since last scan + historical cycle time for that station/part.

    Rules:
    - Only count stations active in the CURRENT shift (no cross-shift carryover)
    - Cap estimate at 1.0× average earned hours per box (can't exceed one box)
    - Stop estimating if elapsed > 2× avg cycle time (station likely idle/changeover)
    - Returns {station: estimated_wip_hours} and total_estimated_wip

    For historical dates, returns empty (no live estimation needed).
    """
    if target_date != date.today():
        return {}, 0.0

    now = datetime.now()
    now_decimal = now.hour + now.minute / 60

    # Determine current active shifts
    active_shifts = []
    for shift_key, shift in SHIFTS.items():
        if shift["start"] <= now_decimal <= shift["end"]:
            active_shifts.append(shift_key)

    if not active_shifts:
        return {}, 0.0

    # Find earliest start of current active shifts
    current_shift_start_decimal = min(SHIFTS[s]["start"] for s in active_shifts)
    current_shift_start_dt = now.replace(
        hour=int(current_shift_start_decimal),
        minute=int((current_shift_start_decimal % 1) * 60),
        second=0, microsecond=0
    )

    # Build last scan per station within current shift
    last_scan_per_station = {}
    for point in timeline:
        # Parse time back to datetime
        if not point["time"]:
            continue
        try:
            t = datetime.strptime(
                target_date.strftime("%Y-%m-%d") + " " + point["time"],
                "%Y-%m-%d %H:%M:%S"
            )
        except Exception:
            continue

        # Only consider scans within current shift window
        if t >= current_shift_start_dt:
            station = point["station"]
            if station not in last_scan_per_station or t > last_scan_per_station[station]["time"]:
                last_scan_per_station[station] = {
                    "time": t,
                    "earned": point["earned"],
                    "part": point.get("part", ""),
                    "minutes": point["minutes"],
                }

    if not last_scan_per_station:
        return {}, 0.0

    # For each active station, estimate WIP since last scan
    station_wip = {}
    total_wip = 0.0
    hour_bracket = get_hour_bracket(now.hour)

    for station, last in last_scan_per_station.items():
        elapsed_min = (now - last["time"]).total_seconds() / 60
        part = last["part"]

        # Look up cycle time: prefer station+part+bracket, fall back to station+part or station
        cycle_data = (
            model.get((station, part, hour_bracket)) or
            model.get((station, part, "10-14")) or   # fallback to peak bracket
            None
        )

        if cycle_data is None or cycle_data["avg_minutes"] <= 0:
            # No model data — use average earned / 60min as rough rate
            rate_per_min = last["earned"] / 60.0 if last["earned"] > 0 else 0
            estimated = min(elapsed_min * rate_per_min, last["earned"])
        else:
            avg_minutes = cycle_data["avg_minutes"]
            avg_earned = cycle_data["avg_earned"]

            # Stop estimating if elapsed > 2× average cycle (changeover/idle)
            if elapsed_min > avg_minutes * 2.0:
                continue

            fraction = min(elapsed_min / avg_minutes, 1.0)
            estimated = avg_earned * fraction

        if estimated > 0:
            station_wip[station] = round(estimated, 3)
            total_wip += estimated

    return station_wip, round(total_wip, 2)


# =============================================================================
# API ENDPOINTS
# =============================================================================

@app.route("/")
def index():
    # The rotator: full-screen page that switches between the two reports.
    return render_template("rotator.html")

@app.route("/station")
def station_view():
    # Station status report (reads the boards' JSON snapshots).
    return render_template("dashboard.html")

@app.route("/aheadbehind")
def aheadbehind_view():
    # Ahead/Behind report (SQL + Excel earned-hours vs target pace).
    return render_template("aheadbehind.html")

@app.route("/api/status")
def api_status():
    # Station report data — one entry per area, read from the shared JSON files.
    areas = [read_station_area(a) for a in STATION_AREAS]
    # Same-day at-risk: any order due TODAY (days_left==0), still unfinished
    # (remaining>0), once it's past 14:00 (server local time). Covers in-progress,
    # queued-at-a-station, and unassigned orders — the banner lists them all; a
    # station's stripe pulses when it holds an at-risk order (running or queued).
    now = datetime.now()
    past_2pm = (now.hour + now.minute / 60.0) >= 14.0
    at_risk = []
    station_risk = set()   # (area, station_id) to pulse
    for a in areas:
        for c in a.pop('risk_raw', []):
            if past_2pm and c.get('days_left') == 0 and (c.get('remaining') or 0) > 0:
                at_risk.append({'area': a.get('area'), 'label': a.get('label'),
                                'station': c.get('station'), 'where': c.get('where'),
                                'pdo': c.get('pdo'), 'part': c.get('part'),
                                'ship_date': c.get('ship_date')})
                if c.get('station'):
                    station_risk.add((a.get('area'), c.get('station')))
        for s in a.get('stations', []):
            s['at_risk'] = (a.get('area'), s.get('id')) in station_risk
    return jsonify({
        'areas': areas,
        'at_risk': at_risk,
        'past_2pm': past_2pm,
        'server_time': now.strftime('%H:%M:%S'),
        'stale_after': STATION_STALE_AFTER_SEC,
    })


def build_dashboard_payload(target_date):
    """Build the full Ahead/Behind payload for a date. Shared by the /api/dashboard
    route AND the background writer that publishes aheadbehind_status.js to the
    share (so the firewall-free shared viewer can show this report too)."""
    is_today = (target_date == date.today())

    # Get capacity from Excel
    capacity = get_all_capacity(target_date)

    # Get earned hours from SQL
    earned = get_earned_hours(target_date)

    # Build target pace line (layered by shift)
    target_hours = capacity["total_target"]
    total_shift_hours = capacity["total_shift_hours"]
    pace_line = build_target_pace(target_hours, total_shift_hours)

    # Per-area pace lines (layered by shift within each area)
    area_pace = {}
    for area in ["AD", "PK", "RP"]:
        area_cap = capacity["areas"].get(area, {})
        area_target = area_cap.get("target_hours", 0)
        area_shift_hours = area_cap.get("shift_hours", {"Ca 1": 0, "Ca 2": 0, "HC": 0})
        area_pace[area] = build_target_pace(area_target, area_shift_hours)

    # Per-area timelines (filter main timeline by area)
    area_timelines = {"AD": [], "PK": [], "RP": []}
    area_cumulatives = {"AD": 0, "PK": 0, "RP": 0}
    for point in earned.get("timeline", []):
        area = point["area"]
        area_cumulatives[area] += point["earned"]
        area_timelines[area].append({
            "time": point["time"],
            "minutes": point["minutes"],
            "earned": point["earned"],
            "cumulative": round(area_cumulatives[area], 2),
            "station": point["station"],
        })

    # Station-level WIP estimation (live only)
    station_model = get_station_model()
    station_wip, total_estimated_wip = estimate_station_wip(
        earned.get("timeline", []), station_model, target_date
    )

    # Build smooth estimated line: FG cumulative + estimated station WIP
    # at each minute from last FG scan to now
    estimated_smooth_line = []
    if is_today and total_estimated_wip > 0:
        timeline = earned.get("timeline", [])
        last_fg_cumulative = timeline[-1]["cumulative"] if timeline else 0
        last_fg_minute = timeline[-1]["minutes"] if timeline else 0
        now_minute = (datetime.now().hour - int(DAY_START)) * 60 + datetime.now().minute

        # Build minute-by-minute estimated line from last FG scan to now
        for m in range(last_fg_minute, now_minute + 1):
            fraction = (m - last_fg_minute) / max(now_minute - last_fg_minute, 1)
            estimated_smooth_line.append({
                "minutes": m,
                "estimated": round(last_fg_cumulative + total_estimated_wip * fraction, 2)
            })

    return {
        "date": target_date.isoformat(),
        "is_today": is_today,
        "timestamp": datetime.now().isoformat(),
        "capacity": capacity,
        "earned": earned,
        "pace_line": pace_line,
        "estimated_smooth_line": estimated_smooth_line,
        "station_wip": station_wip,
        "total_estimated_wip": total_estimated_wip,
        "area_pace": area_pace,
        "area_timelines": area_timelines,
        "config": {
            "target_efficiency": TARGET_EFFICIENCY,
            "base_hours": BASE_HOURS_PER_SHIFT,
            "shifts": {k: {"start": v["start"], "end": v["end"]} for k, v in SHIFTS.items()},
            "day_start": DAY_START,
            "day_end": DAY_END,
        }
    }


@app.route("/api/dashboard")
def dashboard_data():
    """Main endpoint — returns everything the dashboard needs.
    Optional query param: ?date=YYYY-MM-DD to view a specific day."""
    date_str = request.args.get("date")
    if date_str:
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            target_date = date.today()
    else:
        target_date = date.today()
    return jsonify(build_dashboard_payload(target_date))


# =============================================================================
# SHARED-DRIVE AHEAD/BEHIND FEED — publish aheadbehind_status.js (JSONP) so the
# firewall-free shared viewer can render this report off the share, no server.
# =============================================================================
import threading

def _atomic_write_text(path, text):
    """Write text to path atomically (temp file + replace) so a reader on the
    share never sees a half-written file."""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(text)
    os.replace(tmp, path)

def write_aheadbehind_status():
    """Compute today's Ahead/Behind payload and write it to the share as a JSONP
    file (window.NPV_AHEADBEHIND = {...}) the shared viewer loads via <script>."""
    try:
        payload = build_dashboard_payload(date.today())
        payload["__written"] = int(time.time() * 1000)  # epoch ms → staleness
        js = "window.NPV_AHEADBEHIND=" + json.dumps(payload, default=str) + ";"
        _atomic_write_text(os.path.join(DASHBOARD_STATUS_DIR, "aheadbehind_status.js"), js)
    except Exception as e:
        print(f"[aheadbehind feed] write FAILED: {e}")

def _aheadbehind_writer_loop(interval_sec=60):
    while True:
        write_aheadbehind_status()
        time.sleep(interval_sec)

def start_aheadbehind_feed():
    t = threading.Thread(target=_aheadbehind_writer_loop, daemon=True)
    t.start()
    print("[aheadbehind feed] publishing aheadbehind_status.js to the share every 60s")


@app.route("/api/refresh_capacity")
def refresh_capacity():
    """Force-clear capacity cache so next request re-reads Excel files immediately."""
    _capacity_cache.clear()
    return jsonify({"status": "ok", "message": "Capacity cache cleared — will reload on next request"})


@app.route("/api/capacity")
def capacity_only():
    """Debug endpoint — capacity parsed from attendance files, shown as HTML table."""
    date_str = request.args.get("date")
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else date.today()
    except ValueError:
        target_date = date.today()

    cap = get_all_capacity(target_date)

    rows_html = ""
    grand_workers = 0
    grand_avail = 0
    grand_target = 0

    for area, data in cap["areas"].items():
        if "error" in data:
            rows_html += f"""
            <tr class="error-row">
                <td colspan="8">⚠️ <strong>{area}</strong>: {data['error']}</td>
            </tr>"""
            continue

        shift_h = data.get("shift_hours", {})
        rows_html += f"""
        <tr class="area-header">
            <td><strong>{area}</strong></td>
            <td>{data['workers']}</td>
            <td>{data['available_hours']:.2f}h</td>
            <td>{data['target_hours']:.2f}h</td>
            <td>{shift_h.get('Ca 1', 0):.2f}h</td>
            <td>{shift_h.get('Ca 2', 0):.2f}h</td>
            <td>{shift_h.get('HC', 0):.2f}h</td>
            <td>{data.get('borrow_hours', 0):.2f}h</td>
        </tr>"""

        # Worker detail rows
        for w in data.get("details", []):
            lend_str = f"{w['lend_borrow_clock']:+.0f}h" if w.get("lend_borrow_clock") else "—"
            lend_color = "red" if w.get("lend_borrow_clock", 0) < 0 else ("green" if w.get("lend_borrow_clock", 0) > 0 else "")
            rows_html += f"""
        <tr>
            <td class="indent">{w.get('name', '')} <span class="msnv">{w['msnv']}</span></td>
            <td>—</td>
            <td>{w['total']:.2f}h</td>
            <td>—</td>
            <td colspan="2">{w['shift']}</td>
            <td>OT: {w['ot']:.2f}h</td>
            <td><span class="{lend_color}">{lend_str} {w.get('third_type','')}</span></td>
        </tr>"""

        grand_workers += data['workers']
        grand_avail += data['available_hours']
        grand_target += data['target_hours']

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Capacity Debug — {target_date}</title>
<style>
  body {{ font-family: 'Segoe UI', sans-serif; background: #0f1117; color: #e4e6ed; padding: 24px; }}
  h2 {{ color: #34d399; margin-bottom: 4px; }}
  .sub {{ color: #8b8fa3; font-size: 13px; margin-bottom: 20px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th {{ background: #1a1d27; color: #8b8fa3; text-transform: uppercase; font-size: 11px;
        letter-spacing: 0.06em; padding: 10px 12px; text-align: left; border-bottom: 1px solid #2a2e3a; }}
  td {{ padding: 7px 12px; border-bottom: 1px solid #1e2130; }}
  tr:hover td {{ background: #1a1d27; }}
  .area-header td {{ background: #1a1d27; font-weight: 600; color: #60a5fa; border-top: 2px solid #2a2e3a; }}
  .error-row td {{ color: #f87171; background: rgba(248,113,113,0.08); }}
  .indent {{ padding-left: 28px; color: #c4c9d8; }}
  .msnv {{ color: #565a6e; font-size: 11px; font-family: monospace; margin-left: 6px; }}
  .total-row td {{ background: #22262f; font-weight: 700; color: #fbbf24; border-top: 2px solid #2a2e3a; }}
  .red {{ color: #f87171; }}
  .green {{ color: #34d399; }}
  .nav {{ margin-bottom: 16px; display: flex; gap: 12px; align-items: center; }}
  .nav a {{ color: #60a5fa; text-decoration: none; padding: 6px 14px;
             background: #1a1d27; border: 1px solid #2a2e3a; border-radius: 6px; font-size: 13px; }}
  .nav a:hover {{ background: #2a2e3a; }}
  .date-badge {{ color: #fbbf24; font-family: monospace; font-size: 14px; }}
</style>
</head>
<body>
<h2>Capacity Debug</h2>
<div class="sub">Parsed from attendance files · Target efficiency: {TARGET_EFFICIENCY*100:.0f}%</div>

<div class="nav">
  <a href="/api/capacity?date={( target_date - __import__('datetime').timedelta(days=1)).isoformat()}">◀ Prev Day</a>
  <span class="date-badge">{target_date.strftime('%A, %d %b %Y')}</span>
  <a href="/api/capacity?date={( target_date + __import__('datetime').timedelta(days=1)).isoformat()}">Next Day ▶</a>
  <a href="/api/capacity">Today</a>
</div>

<table>
  <thead>
    <tr>
      <th>Worker / Area</th>
      <th>Count</th>
      <th>Available</th>
      <th>Target (95%)</th>
      <th>Ca 1 Hours</th>
      <th>Ca 2 Hours</th>
      <th>HC Hours</th>
      <th>Lend / Borrow</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
    <tr class="total-row">
      <td>TOTAL</td>
      <td>{grand_workers}</td>
      <td>{grand_avail:.2f}h</td>
      <td>{grand_target:.2f}h</td>
      <td colspan="4">&nbsp;</td>
    </tr>
  </tbody>
</table>
</body>
</html>"""

    return html


PORT = 8080  # avoid Chrome-blocked ports (e.g. 5060); 8080 is safe

if __name__ == "__main__":
    print("=" * 60)
    print("Combined Production Dashboard (Station status + Ahead/Behind)")
    print(f"Date: {date.today()}")
    print(f"Station JSON folder : {DASHBOARD_STATUS_DIR}")
    print(f"Capacity files      : {list(CAPACITY_FILES.keys())}")
    print(f"pyodbc available    : {pyodbc is not None}")
    print(f"Open fullscreen     : http://<this-pc-ip>:{PORT}/   (auto-rotates every 1 min; pick a screen from the top-right selector)")
    print("=" * 60)
    start_aheadbehind_feed()   # publish aheadbehind_status.js to the share for the firewall-free shared viewer
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True)
