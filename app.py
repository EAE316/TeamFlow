# -*- coding: utf-8 -*-
"""
TeamFlow / ניהול משימות צוות
Backend server using standard-library HTTP server and openpyxl.
"""

import os
import sys
import json
import time
PROCESS_STARTED = time.perf_counter()
import shutil
import tempfile
import logging
import base64
import hashlib
import copy
import io
import http.client
import subprocess
import uuid
import zipfile
import socket
import threading
import webbrowser
from datetime import datetime, date, timedelta
from pathlib import Path
from contextlib import contextmanager
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
import urllib.parse
# The HTTP shell does not need Excel libraries. Import only on the local data path.
openpyxl = None
_EXCEL_IMPORT_LOCK = threading.Lock()

def load_excel_libraries():
    global openpyxl, PatternFill, Font, Alignment, Border, Side, get_column_letter, Translator, TranslatorError
    if openpyxl is not None:
        return
    with _EXCEL_IMPORT_LOCK:
        if openpyxl is None:
            import openpyxl as module
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from openpyxl.formula.translate import Translator, TranslatorError
            openpyxl = module


APP_NAME = "ניהול משימות צוות"
APP_VERSION = '1.4.2'
PORT = 8765
TASK_SHEET = "משימות"
LISTS_SHEET = "צוות ורשימות"
HISTORY_SHEET = "היסטוריה"
HEADERS = [
    "משימה",
    "אחראי",
    "עדיפות",
    "סטטוס",
    "קטגוריה",
    "תאריך יעד",
    "% ביצוע",
    "חסום?",
    "הערות",
    "תאריך התחלה",
    "תאריך סיום",
    "מספר משימה",
    "ימים שנותרו",
    "מצב",
    "חודש יעד",
    "נוצר ע\"י",
    "תאריך יצירה",
    "עודכן ע\"י",
    "זמן עדכון",
]
TASK_FIELDS = ('title', 'owner', 'priority', 'status', 'category', 'dueDate', 'progress',
               'blocked', 'notes', 'startDate', 'completedDate', 'id', 'daysLeft', 'health',
               'dueMonth', 'createdBy', 'createdAt', 'updatedBy', 'updatedAt')
HISTORY_HEADERS = ["זמן", "משתמש", "פעולה", "מספר משימה", "משימה", "פרטים"]
START_ROW = 7
# Column U: manual card order. Rows never move in SharePoint mode; only this number changes.
ORDER_COLUMN = 21
MAX_ROW = 500
DEFAULT_PEOPLE = ["אלכס", "מאיה", "ירדן", "סם", "פריה"]
DEFAULT_CATEGORIES = ["פיתוח", "בדיקות", "מוצר", "תפעול", "תיעוד", "תמיכה", "אבטחה", "שחרור גרסה"]
DEFAULT_STATUSES = ["לא התחיל", "בתהליך", "בהמתנה", "בבדיקה", "בוצע", "בוטל"]
DEFAULT_PRIORITIES = ["P0", "P1", "P2", "P3"]
YES = "כן"
NO = "לא"
DONE = "בוצע"
CANCELLED = "בוטל"
HEALTH_CLOSED = "סגור"
HEALTH_BLOCKED = "חסום"

def get_logged_in_excel_user() -> str:
    """
    Automatically detects the logged-in Microsoft Office / Excel / OneDrive user
    from the Windows Registry and environment.
    """
    # 1. Try Microsoft Office UserInfo (exact name configured in Excel / Office)
    try:
        import winreg
        for subpath in [
            r"Software\Microsoft\Office\Common\UserInfo",
            r"Software\Microsoft\Office\16.0\Common\UserInfo",
            r"Software\Microsoft\Office\15.0\Common\UserInfo",
        ]:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subpath) as key:
                    val, _ = winreg.QueryValueEx(key, "UserName")
                    if val and str(val).strip():
                        return str(val).strip()
            except Exception:
                pass
    except Exception:
        pass

    # 2. Try OneDrive Account display name or email
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\OneDrive\Accounts") as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                subname = winreg.EnumKey(key, i)
                try:
                    with winreg.OpenKey(key, subname) as subkey:
                        for field in ["UserName", "UserEmail"]:
                            try:
                                u, _ = winreg.QueryValueEx(subkey, field)
                                if u and str(u).strip():
                                    return str(u).strip()
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass

    # 3. Fallback to Windows username
    try:
        import getpass
        u = getpass.getuser()
        if u and str(u).strip():
            return str(u).strip()
    except Exception:
        pass

    return os.environ.get("USERNAME") or os.environ.get("USER") or "משתמש Excel"
HEALTH_OVERDUE = "באיחור"
# Opening a local file through desktop Excel runs the SharePoint code on this PC.
EXCEL_TEST_MODE = '--excel-test' in sys.argv or os.environ.get('TEAMFLOW_EXCEL_TEST') == '1'


def is_tombstone(value):
    from cloud_excel import is_tombstone as check
    return check(value)


def task_events(value):
    from cloud_excel import parse_events
    return parse_events(value)

HEALTH_DUE_SOON = "קרוב ליעד"
HEALTH_ON_TRACK = "במסלול"

# In PyInstaller windowed mode (console=False), sys.stdout and sys.stderr are None
if sys.stdout is None:
    try:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    except Exception:
        pass
if sys.stderr is None:
    try:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    except Exception:
        pass

# Global threading lock for workbook operations
THREAD_LOCK = threading.RLock()
PICKER_LOCK = threading.Lock()
LOCK_STATE = threading.local()
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAX_REQUEST_BYTES = 28 * 1024 * 1024
DATA_CACHE = {}  # Only the most recently read workbook; never trust timestamps alone.
INSTANCE_ID = uuid.uuid4().hex


class WorkbookConflictError(Exception):
    pass

# The server is a background process, so Windows will not bring its windows to the front.
# A visible, top-most, invisible owner keeps the dialog above the browser and gives it a
# taskbar button. TopMost on an owner that is never shown has no effect.
FILE_PICKER_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$owner = New-Object System.Windows.Forms.Form
$dialog = New-Object System.Windows.Forms.OpenFileDialog
try {
    $area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $owner.Text = 'TeamFlow - בחירת קובץ Excel'
    $owner.FormBorderStyle = 'None'
    $owner.StartPosition = 'Manual'
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Location = New-Object System.Drawing.Point(($area.X + $area.Width / 2), ($area.Y + $area.Height / 3))
    $owner.Opacity = 0
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $true
    $owner.Show()
    $owner.Activate()
    $dialog.Filter = 'Excel Files (*.xlsx)|*.xlsx'
    $dialog.Title = 'TeamFlow - Select Excel workbook'
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
        [Console]::Out.WriteLine($dialog.FileName)
    }
} finally {
    $dialog.Dispose()
    $owner.Dispose()
}
"""


def pick_workbook():
    encoded = base64.b64encode(FILE_PICKER_SCRIPT.encode('utf-16le')).decode('ascii')
    result = subprocess.run(
        ['powershell.exe', '-STA', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
        capture_output=True, encoding='utf-8', timeout=300,
        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
    )
    if result.returncode:
        log(f'Windows file picker failed: {result.stderr}', logging.ERROR)
        raise RuntimeError('לא ניתן לפתוח את סייר הקבצים. סגור חלונות פתוחים ונסה שוב.')
    return result.stdout.strip().lstrip('\ufeff')


def workbook_revision(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def literal_cell(cell, value):
    if isinstance(value, str):
        if len(value) > 32767:
            raise ValueError('הטקסט ארוך מדי לתא Excel (עד 32767 תווים).')
        if any(ord(c) < 32 and c not in '\t\n\r' for c in value):
            raise ValueError('הטקסט מכיל תווי בקרה שאינם נתמכים ב-Excel.')
    cell.value = value
    if isinstance(value, str) and value.startswith('='):
        cell.data_type = 's'


def validate_archive(content):
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        if sum(info.file_size for info in archive.infolist()) > 100 * 1024 * 1024:
            raise ValueError('קובץ האקסל גדול מדי לאחר פריסה')


def validate_task_headers(wb):
    if TASK_SHEET not in wb.sheetnames:
        raise ValueError('הקובץ אינו בתבנית TeamFlow: חסר גיליון משימות')
    ws = wb[TASK_SHEET]
    if [ws.cell(6, c).value for c in range(1, 13)] != HEADERS[:12]:
        raise ValueError('הקובץ אינו בתבנית TeamFlow: כותרות המשימות בשורה 6 אינן תואמות')


def validate_workbook(path):
    load_excel_libraries()
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('יש לבחור קובץ Excel קיים עם סיומת .xlsx')
    with zipfile.ZipFile(path) as archive:
        if sum(info.file_size for info in archive.infolist()) > 100 * 1024 * 1024:
            raise ValueError('קובץ האקסל גדול מדי לאחר פריסה')
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        validate_task_headers(wb)
        ws = wb[TASK_SHEET]
        ids = set()
        for row in ws.iter_rows(min_row=START_ROW, max_col=12, values_only=True):
            if row[11]:
                tid = str(row[11]).strip()
                if tid in ids:
                    raise ValueError('קובץ האקסל מכיל מספרי משימה כפולים')
                ids.add(tid)
    finally:
        wb.close()

# Watchdog heartbeat tracker
LAST_PING_TIME = time.time()
SERVER_START_TIME = time.time()
HAS_CONNECTED = False
SHUTDOWN_REQUESTED = False

def get_app_dir() -> Path:
    """Returns directory where executable or script resides."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

APP_DIR = get_app_dir()

def get_base_dir() -> Path:
    """
    Returns user data directory where config, logs, and default workbook reside.
    If app_dir is user-writable and NOT inside Program Files, uses app_dir (portable mode).
    Otherwise, uses %LOCALAPPDATA%\\TeamFlow.
    """
    prog_files = os.environ.get("ProgramFiles", "C:\\Program Files").lower()
    prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)").lower()
    app_str = str(APP_DIR).lower()

    installed_dir = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData' / 'Local'))) / 'Programs' / 'TeamFlow'
    is_in_prog_files = (app_str.startswith(prog_files) or app_str.startswith(prog_files_x86)
                        or APP_DIR == installed_dir.resolve())

    if not is_in_prog_files:
        try:
            test_file = APP_DIR / ".write_test"
            with open(test_file, "w") as f:
                f.write("1")
            test_file.unlink()
            return APP_DIR
        except Exception:
            pass

    # Program Files or read-only directory: use LocalAppData\TeamFlow
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        data_dir = Path(local_app_data) / "TeamFlow"
    else:
        data_dir = Path.home() / "AppData" / "Local" / "TeamFlow"
    try:
        data_dir.mkdir(parents=True, exist_ok=True)
        return data_dir
    except Exception:
        fallback = Path(tempfile.gettempdir()) / "TeamFlow"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback

def get_web_dir() -> Path:
    """Returns directory where web frontend static files reside."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass and (Path(meipass) / "web").is_dir():
            return Path(meipass) / "web"
        exe_parent = Path(sys.executable).resolve().parent
        if (exe_parent / "web").is_dir():
            return exe_parent / "web"
        if (exe_parent / "_internal" / "web").is_dir():
            return exe_parent / "_internal" / "web"
    return Path(__file__).resolve().parent / "web"

BASE_DIR = get_base_dir()
WEB_DIR = get_web_dir()
CONFIG_FILE = BASE_DIR / "config.json"
LOG_FILE = BASE_DIR / "startup.log"

try:
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        encoding="utf-8",
    )
except Exception:
    fallback_log = Path(tempfile.gettempdir()) / "TeamFlow_startup.log"
    try:
        logging.basicConfig(
            filename=str(fallback_log),
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            encoding="utf-8",
        )
    except Exception:
        pass

def log(msg: str, level=logging.INFO):
    try:
        logging.log(level, msg)
    except Exception:
        pass
    if sys.stdout is not None:
        try:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)
        except Exception:
            pass

def get_configured_workbook_path() -> Path:
    prog_files = os.environ.get("ProgramFiles", "C:\\Program Files").lower()
    prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)").lower()

    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
                if "workbook" in data and data["workbook"]:
                    wb_path = Path(data["workbook"])
                    wb_str = str(wb_path).lower()
                    # Divert any path inside Program Files to BASE_DIR
                    if wb_str.startswith(prog_files) or wb_str.startswith(prog_files_x86):
                        log(f"Workbook path in Program Files detected ({wb_path}), diverting to BASE_DIR")
                    else:
                        if wb_path.is_absolute():
                            return wb_path
                        elif not wb_path.is_absolute():
                            return (BASE_DIR / wb_path).resolve()
        except Exception as e:
            log(f"Error reading config.json: {e}", logging.WARNING)
            raise ValueError('לא ניתן לקרוא את הגדרות הקובץ. בחר מחדש את קובץ האקסל דרך בחירת קובץ.') from e

    default_wb = (BASE_DIR / "Team_Task_Manager.xlsx").resolve()
    # If default workbook doesn't exist yet in BASE_DIR, copy from APP_DIR template if available
    if not default_wb.exists() and APP_DIR != BASE_DIR:
        template_wb = APP_DIR / "Team_Task_Manager.xlsx"
        if template_wb.exists():
            try:
                shutil.copy2(template_wb, default_wb)
            except Exception as e:
                log(f"Could not copy template workbook: {e}", logging.WARNING)
    return default_wb

def set_configured_workbook_path(path: Path):
    write_config({'workbook': str(path.resolve()), 'mode': 'local'})
    stop_cloud()

@contextmanager
def file_lock(lock_path: Path, timeout: float = 3.0):
    key = str(lock_path.resolve())
    held = getattr(LOCK_STATE, 'held', set())
    if key in held:
        yield
        return
    LOCK_STATE.held = held
    if not lock_path.parent.is_dir():
        raise FileNotFoundError('תיקיית קובץ האקסל אינה זמינה. בדוק את החיבור או בחר קובץ אחר.')
    fd = os.open(key, os.O_CREAT | os.O_RDWR)
    start = time.monotonic()
    try:
        if os.fstat(fd).st_size == 0:
            os.write(fd, b'0')
        while True:
            try:
                os.lseek(fd, 0, os.SEEK_SET)
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() - start >= timeout:
                    raise PermissionError('קובץ האקסל בשימוש. נסה שוב בעוד מספר שניות.')
                time.sleep(0.1)
        held.add(key)
        yield
    finally:
        held.discard(key)
        os.close(fd)  # OS releases the lock even if the process crashes.

def format_date_str(val) -> str:
    if not val:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    return s

def compute_health_and_days_left(status: str, blocked: str, due_date_str: str):
    today = datetime.now().date()
    days_left = ""
    due_date = None
    if due_date_str:
        try:
            if isinstance(due_date_str, datetime):
                due_date = due_date_str.date()
            elif isinstance(due_date_str, date):
                due_date = due_date_str
            else:
                due_date = datetime.strptime(str(due_date_str)[:10], "%Y-%m-%d").date()
        except Exception:
            due_date = None

    if status == DONE:
        days_left = DONE
    elif due_date is not None:
        days_left = (due_date - today).days

    if status in (DONE, CANCELLED):
        health = HEALTH_CLOSED
    elif blocked == YES:
        health = HEALTH_BLOCKED
    elif due_date is not None and due_date < today:
        health = HEALTH_OVERDUE
    elif due_date is not None and 0 <= (due_date - today).days <= 2:
        health = HEALTH_DUE_SOON
    else:
        health = HEALTH_ON_TRACK

    return health, days_left


def decode_workbook(wb, workbook_path, revision):
    people = []
    categories = []
    statuses = []
    priorities = []

    # Read lists
    if LISTS_SHEET in wb.sheetnames:
        ws_l = wb[LISTS_SHEET]
        # People: Col A, row 6+
        for r in range(6, ws_l.max_row + 1):
            val = ws_l.cell(row=r, column=1).value
            if val is not None and str(val).strip():
                people.append(str(val).strip())
        # Categories: Col F, row 6+
        for r in range(6, ws_l.max_row + 1):
            val = ws_l.cell(row=r, column=6).value
            if val is not None and str(val).strip():
                categories.append(str(val).strip())
        # Statuses: Col H, row 6+
        for r in range(6, ws_l.max_row + 1):
            val = ws_l.cell(row=r, column=8).value
            if val is not None and str(val).strip():
                statuses.append(str(val).strip())
        # Priorities: Col J, row 6+
        for r in range(6, ws_l.max_row + 1):
            val = ws_l.cell(row=r, column=10).value
            if val is not None and str(val).strip():
                priorities.append(str(val).strip())

    if not people:
        people = list(DEFAULT_PEOPLE)
    if not categories:
        categories = list(DEFAULT_CATEGORIES)
    if not statuses:
        statuses = list(DEFAULT_STATUSES)
    if not priorities:
        priorities = list(DEFAULT_PRIORITIES)

    # Read tasks
    tasks = []
    task_rows = {}
    tombstone_rows = []
    if TASK_SHEET in wb.sheetnames:
        ws_t = wb[TASK_SHEET]
        reserved_ids = {str(ws_t.cell(r, 12).value).strip()
                        for r in range(START_ROW, ws_t.max_row + 1)
                        if ws_t.cell(r, 12).value}
        seen_ids = set()
        next_id = 1
        for r in range(START_ROW, ws_t.max_row + 1):
            title = ws_t.cell(row=r, column=1).value
            task_id = ws_t.cell(row=r, column=12).value
            if title is None and task_id is None:
                continue
            if task_id and is_tombstone(ws_t.cell(row=r, column=20).value):
                tombstone_rows.append(r)  # Deleted through SharePoint; its ID stays reserved.
                continue
            if title is not None and not str(title).strip() and not task_id:
                continue

            owner = str(ws_t.cell(row=r, column=2).value or "").strip()
            priority = str(ws_t.cell(row=r, column=3).value or "P2").strip()
            status = str(ws_t.cell(row=r, column=4).value or DEFAULT_STATUSES[0]).strip()
            category = str(ws_t.cell(row=r, column=5).value or "").strip()
            due_date = format_date_str(ws_t.cell(row=r, column=6).value)

            # Progress
            prog_raw = ws_t.cell(row=r, column=7).value
            prog = 0
            if prog_raw is not None:
                try:
                    prog_float = float(prog_raw)
                    if prog_float <= 1.0 and prog_float > 0.0:
                        prog = int(round(prog_float * 100))
                    else:
                        prog = int(round(prog_float))
                except Exception:
                    prog = 0
            prog = max(0, min(100, prog))

            blocked_val = str(ws_t.cell(row=r, column=8).value or "").strip()
            blocked = YES if blocked_val == YES else NO
            notes = str(ws_t.cell(row=r, column=9).value or "").strip()
            start_date = format_date_str(ws_t.cell(row=r, column=10).value)
            completed_date = format_date_str(ws_t.cell(row=r, column=11).value)
            if task_id:
                tid = str(task_id).strip()
            else:
                while f'T-{next_id:03d}' in reserved_ids:
                    next_id += 1
                tid = f'T-{next_id:03d}'
                reserved_ids.add(tid)
            if tid in seen_ids:
                wb.close()
                raise ValueError('קובץ האקסל מכיל מספרי משימה כפולים. תקן אותם באקסל לפני שמירה.')
            seen_ids.add(tid)
            task_rows[tid] = r

            health, days_left = compute_health_and_days_left(status, blocked, due_date)

            created_by = str(ws_t.cell(row=r, column=16).value or "").strip()
            created_at = format_date_str(ws_t.cell(row=r, column=17).value) or str(ws_t.cell(row=r, column=17).value or "").strip()
            updated_by = str(ws_t.cell(row=r, column=18).value or "").strip()
            updated_at = str(ws_t.cell(row=r, column=19).value or "").strip()
            # Changes and comments stored with the task itself (column T).
            activity = task_events(ws_t.cell(row=r, column=20).value)
            order = ws_t.cell(row=r, column=ORDER_COLUMN).value
            order = float(order) if isinstance(order, (int, float)) and not isinstance(order, bool) else None

            tasks.append({
                "id": tid,
                "title": str(title or "").strip(),
                "owner": owner,
                "priority": priority,
                "status": status,
                "category": category,
                "dueDate": due_date,
                "progress": prog,
                "blocked": blocked,
                "notes": notes,
                "startDate": start_date,
                "completedDate": completed_date,
                "health": health,
                "daysLeft": days_left,
                "createdBy": created_by,
                "createdAt": created_at,
                "updatedBy": updated_by,
                "updatedAt": updated_at,
                "activity": activity,
                "order": order,
            })
        # Manual order: the number in column U when set, otherwise the row itself.
        tasks.sort(key=lambda t: t["order"] if t["order"] is not None else task_rows[t["id"]])

    # Read history (last 60, newest first)
    raw_history = []
    if HISTORY_SHEET in wb.sheetnames:
        ws_h = wb[HISTORY_SHEET]
        header_col2 = str(ws_h.cell(row=1, column=2).value or "").strip()
        has_user_col = (header_col2 == "משתמש")

        for r in range(2, ws_h.max_row + 1):
            t_time = str(ws_h.cell(row=r, column=1).value or "").strip()
            if has_user_col:
                user = str(ws_h.cell(row=r, column=2).value or "").strip()
                act = str(ws_h.cell(row=r, column=3).value or "").strip()
                tid = str(ws_h.cell(row=r, column=4).value or "").strip()
                title = str(ws_h.cell(row=r, column=5).value or "").strip()
                det = str(ws_h.cell(row=r, column=6).value or "").strip()
            else:
                user = ""
                act = str(ws_h.cell(row=r, column=2).value or "").strip()
                tid = str(ws_h.cell(row=r, column=3).value or "").strip()
                title = str(ws_h.cell(row=r, column=4).value or "").strip()
                det = str(ws_h.cell(row=r, column=5).value or "").strip()

            if any([t_time, user, act, tid, title, det]):
                raw_history.append({
                    "time": t_time,
                    "user": user,
                    "action": act,
                    "id": tid,
                    "title": title,
                    "details": det,
                })

    # Take last 60, newest first
    history = list(reversed(raw_history[-60:]))

    wb.close()
    lists_revision = hashlib.sha256(json.dumps([people, categories, statuses, priorities],
                                               ensure_ascii=False).encode()).hexdigest()[:16]
    data = {
        "workbook": str(workbook_path),
        "revision": revision,
        "listsRevision": lists_revision,
        "systemUser": get_logged_in_excel_user(),
        "people": people,
        "categories": categories,
        "statuses": statuses,
        "priorities": priorities,
        "tasks": tasks,
        "taskRows": task_rows,
        "tombstoneRows": tombstone_rows,
        "history": history,
    }
    return data

def build_report(data, start, end, today=None):
    """The weekly report figures, shared by the screen and the Excel download."""
    today = today or date.today().isoformat()
    week_end = (date.fromisoformat(today) + timedelta(days=7)).isoformat()
    tasks = data.get('tasks', [])
    open_tasks = [t for t in tasks if t.get('status') not in (DONE, CANCELLED)]
    in_range = lambda value: bool(value) and start <= str(value)[:10] <= end
    completed = [t for t in tasks if t.get('status') == DONE and in_range(t.get('completedDate'))]
    created = [t for t in tasks if in_range(t.get('createdAt'))]
    overdue = sorted((t for t in open_tasks if t.get('health') == HEALTH_OVERDUE), key=lambda t: t.get('dueDate') or '')
    upcoming = sorted((t for t in open_tasks if t.get('dueDate') and today < t['dueDate'] <= week_end),
                      key=lambda t: t['dueDate'])
    owners = list(dict.fromkeys(list(data.get('people', [])) + [t.get('owner') or '' for t in tasks]))
    people = []
    for owner in owners:
        mine = [t for t in tasks if (t.get('owner') or '') == owner]
        row = {'owner': owner or 'ללא שיוך',
               'open': sum(t in open_tasks for t in mine),
               'done': sum(t in completed for t in mine),
               'overdue': sum(t in overdue for t in mine)}
        if row['open'] or row['done']:
            people.append(row)
    return {'from': start, 'to': end, 'completed': completed, 'created': created,
            'overdue': overdue, 'upcoming': upcoming, 'people': people, 'openCount': len(open_tasks)}


def report_date(value):
    """A yyyy-mm-dd text as a real date (shown dd/mm/yyyy), anything else unchanged."""
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date() if value else ''
    except ValueError:
        return value


def display_date(value):
    parsed = report_date(value)
    return parsed.strftime("%d/%m/%Y") if isinstance(parsed, date) else str(value or '')


def report_workbook_bytes(data, start, end):
    """The report as an .xlsx file: a summary sheet and one sheet per list."""
    load_excel_libraries()
    report = build_report(data, start, end)
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="18181B", end_color="18181B", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    def table(ws, first_row, headers, rows, widths):
        for col, title in enumerate(headers, 1):
            cell = ws.cell(first_row, col, title)
            cell.fill, cell.font = header_fill, header_font
        for r, values in enumerate(rows, first_row + 1):
            for col, value in enumerate(values, 1):
                cell = ws.cell(r, col)
                literal_cell(cell, value)
                if isinstance(value, date):
                    cell.number_format = 'DD/MM/YYYY'

        for col, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = ws.cell(first_row + 1, 1)

    summary = wb.active
    summary.title = 'סיכום'
    summary.sheet_view.rightToLeft = True
    summary['A1'] = f'דוח TeamFlow {display_date(start)} עד {display_date(end)}'
    summary['A1'].font = Font(bold=True, size=14)
    figures = [('הושלמו בתקופה', len(report['completed'])), ('נוצרו בתקופה', len(report['created'])),
               ('פתוחות כעת', report['openCount']), ('באיחור כעת', len(report['overdue'])),
               ('יעד בשבוע הקרוב', len(report['upcoming']))]
    for r, (label, value) in enumerate(figures, 3):
        summary.cell(r, 1, label)
        summary.cell(r, 2, value).font = Font(bold=True)
    table(summary, 10, ['אחראי', 'פתוחות', 'הושלמו בתקופה', 'באיחור'],
          [[p['owner'], p['open'], p['done'], p['overdue']] for p in report['people']], [24, 12, 16, 12])
    summary.freeze_panes = None

    task_columns = ['מספר', 'משימה', 'אחראי', 'עדיפות', 'סטטוס', 'תאריך יעד', 'תאריך סיום']
    for title, rows in (('הושלמו', report['completed']), ('באיחור', report['overdue']),
                        ('יעד בשבוע הקרוב', report['upcoming']), ('נוצרו', report['created'])):
        ws = wb.create_sheet(title)
        ws.sheet_view.rightToLeft = True
        table(ws, 1, task_columns,
              [[t.get('id', ''), t.get('title', ''), t.get('owner', ''), t.get('priority', ''),
                t.get('status', ''), report_date(t.get('dueDate')), report_date(t.get('completedDate'))] for t in rows],
              [10, 44, 18, 10, 14, 13, 13])
    buffer = io.BytesIO()
    wb.save(buffer)
    wb.close()
    return buffer.getvalue()


class TaskManagerWorkbook:
    """Manages the Excel workbook operations."""

    def __init__(self, path: Path):
        load_excel_libraries()
        self.path = path.resolve()
        self.lock_path = self.path.parent / f"{self.path.name}.lock"
        self.loaded_revision = None

    @staticmethod
    def new_task_id(tasks):
        numbers = [int(t['id'][2:]) for t in tasks if t.get('id', '').startswith('T-') and t['id'][2:].isdigit()]
        return f"T-{max(numbers, default=0) + 1:03d}"

    def ensure_workbook_exists(self):
        with THREAD_LOCK:
            if not self.path.exists():
                if self.path != (BASE_DIR / 'Team_Task_Manager.xlsx').resolve():
                    raise FileNotFoundError('קובץ האקסל שנבחר אינו זמין. בדוק את החיבור או בחר קובץ אחר.')
                with file_lock(self.lock_path):
                    if not self.path.exists():
                        self.create_default_workbook()

    def create_default_workbook(self):
        wb = openpyxl.Workbook()
        ws_tasks = wb.active
        ws_tasks.title = TASK_SHEET
        ws_lists = wb.create_sheet(title=LISTS_SHEET)
        ws_history = wb.create_sheet(title=HISTORY_SHEET)

        # Populate lists
        self._write_lists_sheet(ws_lists, DEFAULT_PEOPLE, DEFAULT_CATEGORIES, DEFAULT_STATUSES, DEFAULT_PRIORITIES)

        # Initial sample tasks
        sample_tasks = [
            {
                "id": "T-001",
                "title": "הגדרת סביבת עבודה ותשתיות צוות",
                "owner": "אלכס",
                "priority": "P0",
                "status": "בתהליך",
                "category": "פיתוח",
                "dueDate": (datetime.now().date() + timedelta(days=3)).strftime("%Y-%m-%d"),
                "progress": 65,
                "blocked": NO,
                "notes": "סביבת עבודה מקומית ואינטגרציה לשרת",
                "startDate": datetime.now().strftime("%Y-%m-%d"),
                "completedDate": "",
            },
            {
                "id": "T-002",
                "title": "סקירת ארכיטקטורה ועיצוב ממשק",
                "owner": "מאיה",
                "priority": "P1",
                "status": "בתהליך",
                "category": "מוצר",
                "dueDate": (datetime.now().date() + timedelta(days=1)).strftime("%Y-%m-%d"),
                "progress": 85,
                "blocked": NO,
                "notes": "אישור דשבורד ולוח עבודה מול מנהל מוצר",
                "startDate": datetime.now().strftime("%Y-%m-%d"),
                "completedDate": "",
            },
            {
                "id": "T-003",
                "title": "תיקון תקלה בשירות אימות הרשאות",
                "owner": "ירדן",
                "priority": "P0",
                "status": "בהמתנה",
                "category": "אבטחה",
                "dueDate": (datetime.now().date() - timedelta(days=1)).strftime("%Y-%m-%d"),
                "progress": 30,
                "blocked": YES,
                "notes": "ממתין לקבלת טוקן סודי מצוות אבטחת מידע",
                "startDate": datetime.now().strftime("%Y-%m-%d"),
                "completedDate": "",
            },
            {
                "id": "T-004",
                "title": "כתיבת בדיקות יחידה ואינטגרציה",
                "owner": "סם",
                "priority": "P2",
                "status": "לא התחיל",
                "category": "בדיקות",
                "dueDate": (datetime.now().date() + timedelta(days=5)).strftime("%Y-%m-%d"),
                "progress": 0,
                "blocked": NO,
                "notes": "כיסוי תרחישי שגיאה ואימות ביצועים",
                "startDate": datetime.now().strftime("%Y-%m-%d"),
                "completedDate": "",
            },
            {
                "id": "T-005",
                "title": "עדכון תיעוד והוראות התקנה",
                "owner": "פריה",
                "priority": "P3",
                "status": DONE,
                "category": "תיעוד",
                "dueDate": datetime.now().strftime("%Y-%m-%d"),
                "progress": 100,
                "blocked": NO,
                "notes": "נבדק בהצלחה",
                "startDate": datetime.now().strftime("%Y-%m-%d"),
                "completedDate": datetime.now().strftime("%Y-%m-%d"),
            },
        ]
        self._write_tasks_sheet(ws_tasks, sample_tasks)

        # Initial history entries
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        history_rows = [
            [now_str, get_logged_in_excel_user(), "יצירת קובץ", "T-001", "הקמת קובץ משימות חדש", "נוצר קובץ ראשוני עם הגדרות ברירת מחדל"]
        ]
        self._write_history_sheet(ws_history, history_rows)

        self.path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(self.path))
        wb.close()
        log(f"Created default workbook at {self.path}")

    def format_existing_workbook(self):
        try:
            with file_lock(self.lock_path):
                wb = openpyxl.load_workbook(str(self.path))
                changed = False
                for sheet_name in [TASK_SHEET, LISTS_SHEET, HISTORY_SHEET]:
                    if sheet_name not in wb.sheetnames:
                        wb.create_sheet(title=sheet_name)
                        changed = True

                ws_tasks = wb[TASK_SHEET]
                ws_lists = wb[LISTS_SHEET]
                ws_history = wb[HISTORY_SHEET]

                self._apply_tasks_formatting(ws_tasks)
                self._apply_lists_formatting(ws_lists)
                self._apply_history_formatting(ws_history)

                wb.save(str(self.path))
        except Exception as e:
            log(f"Warning during existing workbook format: {e}", logging.WARNING)

    def _apply_tasks_formatting(self, ws):
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False

        dark_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        white_bold = Font(name="Arial", size=11, color="FFFFFF", bold=True)

        ws["A1"] = "משימות"
        ws["A1"].fill = dark_fill
        ws["A1"].font = white_bold
        ws["A1"].alignment = Alignment(horizontal="right", vertical="center")

        for col_idx, h in enumerate(HEADERS, 1):
            cell = ws.cell(row=6, column=col_idx, value=h)
            cell.fill = dark_fill
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="center" if col_idx in (3, 4, 7, 8, 12, 13, 14, 16, 17, 18, 19) else "right", vertical="center")

        ws.freeze_panes = "A7"
        widths = [38, 16, 11, 16, 18, 13, 10, 11, 42, 13, 15, 12, 12, 12, 10, 18, 14, 18, 20]
        for idx, w in enumerate(widths, 1):
            col_letter = get_column_letter(idx)
            ws.column_dimensions[col_letter].width = w

    def _apply_lists_formatting(self, ws):
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False

        dark_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        white_bold = Font(name="Arial", size=11, color="FFFFFF", bold=True)
        section_font = Font(name="Arial", size=11, color="0F172A", bold=True)

        ws["A1"] = "צוות ורשימות"
        ws["A1"].fill = dark_fill
        ws["A1"].font = white_bold

        ws["A4"] = "אנשים"
        ws["A4"].font = section_font
        ws["F4"] = "קטגוריות"
        ws["F4"].font = section_font
        ws["H4"] = "סטטוסים"
        ws["H4"].font = section_font
        ws["J4"] = "עדיפויות"
        ws["J4"].font = section_font

        headers_row5 = {
            1: "שם",
            2: "צוות",
            3: "אימייל",
            4: "פעיל",
            6: "קטגוריה",
            8: "סטטוס",
            10: "קוד",
        }
        for col_idx, text in headers_row5.items():
            cell = ws.cell(row=5, column=col_idx, value=text)
            cell.fill = dark_fill
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="center", vertical="center")

        widths = {
            "A": 22, "B": 18, "C": 26, "D": 10, "E": 6,
            "F": 22, "G": 6, "H": 18, "I": 6, "J": 14
        }
        for col_l, w in widths.items():
            ws.column_dimensions[col_l].width = w

    def _apply_history_formatting(self, ws):
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False

        dark_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        white_bold = Font(name="Arial", size=11, color="FFFFFF", bold=True)

        for col_idx, h in enumerate(HISTORY_HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = dark_fill
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="right" if col_idx in (1, 5, 6) else "center", vertical="center")

        ws.freeze_panes = "A2"
        widths = [21, 18, 16, 14, 38, 55]
        for idx, w in enumerate(widths, 1):
            col_letter = get_column_letter(idx)
            ws.column_dimensions[col_letter].width = w

    def _write_lists_sheet(self, ws, people, categories, statuses, priorities):
        self._apply_lists_formatting(ws)

        # Preserve team/email/active metadata for people who remain in the list.
        metadata = {str(ws.cell(r, 1).value): [ws.cell(r, c).value for c in (2, 3, 4)]
                    for r in range(6, ws.max_row + 1) if ws.cell(r, 1).value}
        # Clear existing entries from row 6 down
        max_r = max(len(people) + 5, len(categories) + 5, len(statuses) + 5,
                    len(priorities) + 5, ws.max_row)
        for r in range(6, max_r + 1):
            for c in [1, 2, 3, 4, 6, 8, 10]:
                ws.cell(row=r, column=c).value = None

        # Write people
        for i, person in enumerate(people):
            r = 6 + i
            literal_cell(ws.cell(row=r, column=1), person)
            for c, val in zip((2, 3, 4), metadata.get(person, [None, None, YES])):
                ws.cell(row=r, column=c).value = val

        # Write categories
        for i, cat in enumerate(categories):
            r = 6 + i
            literal_cell(ws.cell(row=r, column=6), cat)

        # Write statuses
        for i, st in enumerate(statuses):
            r = 6 + i
            literal_cell(ws.cell(row=r, column=8), st)

        # Write priorities
        for i, pr in enumerate(priorities):
            r = 6 + i
            literal_cell(ws.cell(row=r, column=10), pr)

    def _write_tasks_sheet(self, ws, tasks, previous=None, previous_rows=None, tombstone_rows=()):
        previous = previous or {}
        previous_rows = previous_rows or {}
        # Capture whole rows by task ID before anything is overwritten.
        last_column = max(ws.max_column, len(HEADERS))
        snapshots = {
            tid: [(c.value, c.data_type, copy.copy(c._style), c.coordinate)
                  for c in (ws.cell(row, col) for col in range(1, last_column + 1))]
            for tid, row in previous_rows.items()
        }
        occupied = set(previous_rows.values()) | set(tombstone_rows)
        self._apply_tasks_formatting(ws)

        fill_overdue = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")
        fill_blocked = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

        priority_styles = {
            "P0": (PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"), Font(color="991B1B", bold=True)),
            "P1": (PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid"), Font(color="9A3412", bold=True)),
            "P2": (PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"), Font(color="92400E", bold=True)),
            "P3": (PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"), Font(color="1D4ED8", bold=True)),
        }

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
        )

        max_row_to_clear = max(ws.max_row, len(tasks) + START_ROW - 1)

        # Columns past S belong to the task on that row (T holds SharePoint history).
        for row_idx in range(START_ROW, max_row_to_clear + 1):
            index = row_idx - START_ROW
            saved = snapshots.get(tasks[index].get('id')) if index < len(tasks) else None
            for col_idx in range(len(HEADERS) + 1, last_column + 1):
                cell = ws.cell(row_idx, col_idx)
                if type(cell).__name__ == 'MergedCell':
                    continue
                carried = saved[col_idx - 1] if saved else None
                if carried and carried[0] is not None:
                    self._restore_cell(cell, carried)
                elif row_idx in occupied:
                    cell.value = None
                    cell.style = 'Normal'

        if any(t.get('order') is not None for t in tasks):
            ws.cell(6, ORDER_COLUMN).value = 'סדר'
            for index, t in enumerate(tasks):
                ws.cell(START_ROW + index, ORDER_COLUMN).value = index + 1

        for row_idx in range(START_ROW, max_row_to_clear + 1):
            task_idx = row_idx - START_ROW
            if task_idx < len(tasks):
                t = tasks[task_idx]
                status = t.get("status", DEFAULT_STATUSES[0])
                blocked = t.get("blocked", NO)
                due_date_str = t.get("dueDate", "")
                health, days_left = compute_health_and_days_left(status, blocked, due_date_str)
                due_month = due_date_str[:7] if due_date_str else ""

                progress_val = t.get("progress", 0)
                try:
                    progress_num = max(0, min(100, int(progress_val)))
                except Exception:
                    progress_num = 0

                row_values = [
                    t.get("title", ""),
                    t.get("owner", ""),
                    t.get("priority", "P2"),
                    status,
                    t.get("category", ""),
                    due_date_str,
                    progress_num / 100.0,
                    blocked,
                    t.get("notes", ""),
                    t.get("startDate", ""),
                    t.get("completedDate", ""),
                    t.get("id", f"T-{task_idx + 1:03d}"),
                    days_left,
                    health,
                    due_month,
                    t.get("createdBy", ""),
                    t.get("createdAt", ""),
                    t.get("updatedBy", ""),
                    t.get("updatedAt", ""),
                ]

                # Determine base row fill
                row_fill = None
                if blocked == YES:
                    row_fill = fill_blocked
                elif health == HEALTH_OVERDUE:
                    row_fill = fill_overdue

                for col_idx, val in enumerate(row_values, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
                    cell.fill = row_fill if row_fill else PatternFill(fill_type=None)

                    # Date formatting
                    if col_idx in (6, 10, 11):
                        cell.number_format = "yyyy-mm-dd"
                        if val:
                            try:
                                cell.value = datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
                            except Exception:
                                literal_cell(cell, str(val))
                        else:
                            cell.value = ""
                    elif col_idx == 7:
                        cell.number_format = "0%"
                        cell.value = val
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 3:  # Priority
                        literal_cell(cell, str(val))
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if str(val) in priority_styles:
                            p_fill, p_font = priority_styles[str(val)]
                            cell.fill = p_fill
                            cell.font = p_font
                        else:
                            cell.font = Font(name="Arial", size=10)
                    else:
                        literal_cell(cell, val)
                        if col_idx in (4, 8, 12, 13, 14, 15, 16, 17, 18, 19):
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.font = Font(name="Arial", size=10)

                # Keep a formula unless that field changed. Calculated columns M-O always keep it.
                saved = snapshots.get(t.get('id'))
                if saved:
                    before = previous.get(t.get('id'), {})
                    for col_idx, key in enumerate(TASK_FIELDS, 1):
                        carried = saved[col_idx - 1]
                        if carried[1] == 'f' and (col_idx in (13, 14, 15) or t.get(key) == before.get(key)):
                            self._restore_cell(ws.cell(row_idx, col_idx), carried, keep_style=False)
            else:
                # Clear empty row cells
                for col_idx in range(1, len(HEADERS) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = None

    @staticmethod
    def _restore_cell(cell, saved, keep_style=True):
        """Write a captured cell, shifting relative formula references to its new row."""
        value, data_type, style, origin = saved
        if data_type == 'f' and isinstance(value, str) and origin != cell.coordinate:
            try:
                value = Translator(value, origin=origin).translate_formula(cell.coordinate)
            except TranslatorError:
                pass
        cell.value = value
        if data_type == 's':
            cell.data_type = 's'  # Literal text that starts with "=" stays text.
        if keep_style:
            cell._style = copy.copy(style)

    def _write_history_sheet(self, ws, history_rows):
        self._apply_history_formatting(ws)
        # Clear existing
        max_r = max(ws.max_row, len(history_rows) + 2)
        for r in range(2, max_r + 1):
            for c in range(1, 7):
                ws.cell(row=r, column=c).value = None

        for i, row_data in enumerate(history_rows):
            r = 2 + i
            for c, val in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=c, value=val)
                if c in (2, 3, 4):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")

    def read_all_data(self):
        with THREAD_LOCK:
            self.ensure_workbook_exists()
            with file_lock(self.lock_path):
                # Parse and hash exactly the same bytes, even when Excel replaces the file.
                content = self.path.read_bytes()
                revision = hashlib.sha256(content).hexdigest()
                self.loaded_revision = revision
                cache_key = (str(self.path), revision, date.today())
                if cache_key in DATA_CACHE:
                    return copy.deepcopy(DATA_CACHE[cache_key])
                validate_archive(content)
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                try:
                    validate_task_headers(wb)
                except Exception:
                    wb.close()
                    raise
                data = decode_workbook(wb, str(self.path), revision)
                DATA_CACHE.clear()
                DATA_CACHE[cache_key] = copy.deepcopy(data)
                return data

    def save_all_data(self, tasks, lists=None, new_history_entry=None):
        with THREAD_LOCK:
            self.ensure_workbook_exists()
            with file_lock(self.lock_path):
                # 1. Read existing lists and history if not provided
                if self.loaded_revision and workbook_revision(self.path) != self.loaded_revision:
                    raise WorkbookConflictError('קובץ האקסל השתנה בזמן העריכה. רענן ובדוק את השינויים לפני שמירה.')
                revision_before_save = workbook_revision(self.path)
                before = self.read_all_data()
                wb = openpyxl.load_workbook(str(self.path))
                if TASK_SHEET not in wb.sheetnames:
                    wb.create_sheet(TASK_SHEET)
                if LISTS_SHEET not in wb.sheetnames:
                    wb.create_sheet(LISTS_SHEET)
                if HISTORY_SHEET not in wb.sheetnames:
                    wb.create_sheet(HISTORY_SHEET)

                ws_tasks = wb[TASK_SHEET]
                ws_lists = wb[LISTS_SHEET]
                ws_history = wb[HISTORY_SHEET]
                # Upgrade five-column legacy history before adding six-column rows.
                if ws_history.cell(1, 2).value != HISTORY_HEADERS[1]:
                    ws_history.insert_cols(2)
                self._apply_history_formatting(ws_history)

                # Update lists if provided
                if lists:
                    people = lists.get("people", DEFAULT_PEOPLE)
                    categories = lists.get("categories", DEFAULT_CATEGORIES)
                    statuses = lists.get("statuses", DEFAULT_STATUSES)
                    priorities = lists.get("priorities", DEFAULT_PRIORITIES)
                    self._write_lists_sheet(ws_lists, people, categories, statuses, priorities)

                # Write tasks
                self._write_tasks_sheet(ws_tasks, tasks,
                                        previous={t['id']: t for t in before['tasks']},
                                        previous_rows=before['taskRows'],
                                        tombstone_rows=before['tombstoneRows'])

                # Append history entry if provided
                if new_history_entry:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    user = new_history_entry.get("user") or get_logged_in_excel_user()
                    row_data = [
                        now_str,
                        user,
                        new_history_entry.get("action", ""),
                        new_history_entry.get("id", ""),
                        new_history_entry.get("title", ""),
                        new_history_entry.get("details", ""),
                    ]
                    # Append at the next available row in history
                    next_r = ws_history.max_row + 1
                    if next_r < 2:
                        next_r = 2
                    for col_idx, val in enumerate(row_data, 1):
                        cell = ws_history.cell(row=next_r, column=col_idx)
                        literal_cell(cell, val)
                        if col_idx in (2, 3, 4):
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="right", vertical="center")

                self._replace_file(wb, revision_before_save)

    def _replace_file(self, wb, revision_before_save):
        """Save atomically, refusing if the file changed underneath us."""
        tmp_path = self.path.parent / f".tmp_{os.getpid()}_{time.time_ns()}.xlsx"
        try:
            wb.save(str(tmp_path))
            wb.close()
            if workbook_revision(self.path) != revision_before_save:
                raise WorkbookConflictError('קובץ האקסל השתנה בזמן השמירה. השינוי החיצוני נשמר; רענן ונסה שוב.')
            os.replace(str(tmp_path), str(self.path))
            self.loaded_revision = workbook_revision(self.path)
            DATA_CACHE.clear()
            log(f"Workbook saved atomically at {self.path}")
        except PermissionError:
            tmp_path.unlink(missing_ok=True)
            raise PermissionError("אי אפשר לשמור. סגור את קובץ האקסל ונסה שוב.")
        except Exception:
            tmp_path.unlink(missing_ok=True)
            raise
        finally:
            wb.close()

    def add_task_event(self, task_id, event):
        """Append to one task's history cell (column T) without rewriting the sheet."""
        from cloud_excel import append_event
        with THREAD_LOCK:
            self.ensure_workbook_exists()
            with file_lock(self.lock_path):
                revision_before_save = workbook_revision(self.path)
                current = self.read_all_data()
                if not any(t['id'] == task_id for t in current['tasks']):
                    raise ValueError('המשימה אינה קיימת עוד. רענן את הנתונים.')
                wb = openpyxl.load_workbook(str(self.path))
                try:
                    cell = wb[TASK_SHEET].cell(current['taskRows'][task_id], 20)
                    literal_cell(cell, append_event(cell.value, dict(event, id=task_id)))
                    cell.alignment = Alignment(wrap_text=False, vertical="center")
                except Exception:
                    wb.close()
                    raise
                self._replace_file(wb, revision_before_save)

_CLOUD = None
_PENDING = None

def read_config():
    if not CONFIG_FILE.exists():
        return {}
    data = json.loads(CONFIG_FILE.read_text(encoding='utf-8-sig'))
    if not isinstance(data, dict):
        raise ValueError('קובץ ההגדרות אינו תקין. בחר מחדש מקור נתונים.')
    return data

def write_config(data):
    temporary = CONFIG_FILE.with_name(f'.config-{uuid.uuid4().hex}.tmp')
    try:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, CONFIG_FILE)
    finally:
        temporary.unlink(missing_ok=True)

def new_connection(url, local_test=False):
    from cloud_excel import ExcelConnection

    def activate(connection):
        global _CLOUD, _PENDING
        if connection is _CLOUD:
            return  # Restored from saved settings; nothing to switch.
        with THREAD_LOCK:
            if connection is not _PENDING:
                connection.stop()
                return
            try:
                config = read_config()
            except (ValueError, OSError):
                config = {}
            config.update(mode='excel', cloud_url=connection.url, excelTest=connection.local_test)
            if connection.local_test:
                config['workbook'] = connection.url
            write_config(config)
            previous, _CLOUD, _PENDING = _CLOUD, connection, None
            if previous is not None:
                previous.stop()

    return ExcelConnection(url, decode_workbook,
        {'task': TASK_SHEET, 'lists': LISTS_SHEET, 'history': HISTORY_SHEET, 'headers': HEADERS},
        BASE_DIR / 'excel-recovery', activate, local_test=local_test)

def start_cloud(url, local_test=False):
    global _PENDING
    from cloud_excel import normalize_url
    url = str(Path(url).resolve()) if local_test else normalize_url(url)
    with THREAD_LOCK:
        if _PENDING is not None:
            if _PENDING.status()['state'] == 'connecting':
                raise ValueError('Excel כבר מתחבר. השלם את הכניסה בחלון Excel והמתן.')
            _PENDING.stop()
        _PENDING = new_connection(url, local_test)
        return _PENDING

def stop_cloud(wait=False):
    """Stop Excel connections. On exit, wait for test-mode Excel, which is hidden, to close itself."""
    global _CLOUD, _PENDING
    connections = [c for c in (_CLOUD, _PENDING) if c is not None]
    for connection in connections:
        connection.stop()
    _CLOUD = _PENDING = None
    if wait:
        for connection in connections:
            thread = getattr(connection, '_thread', None)
            if getattr(connection, 'local_test', False) and thread is not None:
                thread.join(15)

def active_cloud():
    """The connection for the saved data source, or None in local mode."""
    global _CLOUD
    with THREAD_LOCK:
        config = read_config()
        if config.get('mode') != 'excel':
            return None
        if config.get('excelTest') and not EXCEL_TEST_MODE:
            return None  # A test connection reopens as the plain local file.
        if _CLOUD is None:
            _CLOUD = new_connection(config['cloud_url'], bool(config.get('excelTest')))
        return _CLOUD

def connection_status():
    pending = _PENDING
    attempt = pending.status() if pending is not None else None
    if attempt and attempt['state'] != 'error':
        return {**attempt, 'state': 'connecting'}
    cloud = active_cloud()
    status = cloud.status() if cloud is not None else {
        'mode': 'local', 'state': 'ready', 'message': 'קובץ מקומי • ללא אישור סנכרון ענן'}
    if attempt:
        status['attempt'] = attempt
    return status

class CloudTaskWorkbook:
    """Turns a handler's full task list back into one Excel operation."""

    def __init__(self, cloud, data, task_id=None, expected=None, lists_revision=None):
        self.cloud, self.data, self.task_id, self.expected = cloud, copy.deepcopy(data), task_id, expected
        self.lists_revision = lists_revision
        self.written = False

    def read_all_data(self):
        # After a write, hand back what Excel now holds rather than the snapshot we started from.
        return self.cloud.snapshot() if self.written else copy.deepcopy(self.data)

    def new_task_id(self, tasks):
        return self.cloud.new_task_id()

    def save_all_data(self, tasks, lists=None, new_history_entry=None):
        from cloud_excel import CloudConflict, CLOUD_LIMIT
        if lists is not None:
            if [t['id'] for t in tasks] != [t['id'] for t in self.data['tasks']]:
                raise CloudConflict(CLOUD_LIMIT)
            self.cloud.submit('lists', lists, self.lists_revision)
            self.written = True
            return
        old = {t['id']: t for t in self.data['tasks']}
        new = {t['id']: t for t in tasks}
        added = [i for i in new if i not in old]
        removed = [i for i in old if i not in new]
        changed = [i for i in new if i in old and new[i] != old[i]]
        same_order = [t['id'] for t in tasks if t['id'] in old] == [i for i in old if i in new]
        if not same_order:
            raise CloudConflict(CLOUD_LIMIT)
        self.written = True
        if self.task_id is None and len(added) == 1 and not removed and not changed:
            self.cloud.submit('create', new[added[0]], new_history_entry)
        elif self.task_id is not None and removed == [self.task_id] and not added and not changed:
            self.cloud.submit('delete', self.task_id, self.expected, new_history_entry)
        elif self.task_id is not None and not added and not removed and changed in ([], [self.task_id]):
            if changed:
                self.cloud.submit('update', new[self.task_id], self.expected, new_history_entry)
        else:
            raise CloudConflict(CLOUD_LIMIT)


def apply_bulk_action(t, action, val, user, now_str, today_str):
    """One bulk change applied to one task; shared by local and SharePoint mode."""
    t["updatedBy"] = user
    t["updatedAt"] = now_str
    if action == "done" or (action == "status" and val == DONE):
        t.update(status=DONE, progress=100, completedDate=today_str, blocked=NO,
                 health=HEALTH_CLOSED, daysLeft=DONE)
    elif action == "status":
        t["status"] = val
        t["completedDate"] = ''
        t["health"], t["daysLeft"] = compute_health_and_days_left(t["status"], t["blocked"], t.get("dueDate", ""))
    elif action == "owner":
        t["owner"] = val
    elif action == "priority":
        t["priority"] = val
    elif action == "blocked":
        t["blocked"] = val if val else YES
        t["health"], t["daysLeft"] = compute_health_and_days_left(t["status"], t["blocked"], t.get("dueDate", ""))
    return t


BULK_ACTIONS = ("done", "status", "owner", "priority", "blocked", "delete")


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True
    allow_reuse_address = False

    def server_bind(self):
        # Windows SO_REUSEADDR permits another process to claim a live port.
        # Use exclusive binding so a second launch cannot serve the old copy's UI.
        if hasattr(socket, 'SO_EXCLUSIVEADDRUSE'):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()

class TaskRequestHandler(BaseHTTPRequestHandler):
    server_version = "TeamFlowHTTP/1.0"

    def trusted_request(self):
        expected = f'127.0.0.1:{self.server.server_port}'
        if self.headers.get('Host') not in (expected, f'localhost:{self.server.server_port}'):
            self.send_json({'ok': False, 'error': 'Invalid host'}, 403)
            return False
        origin = self.headers.get('Origin')
        if origin and origin not in (f'http://{expected}', f'http://localhost:{self.server.server_port}'):
            self.send_json({'ok': False, 'error': 'Invalid origin'}, 403)
            return False
        if self.headers.get('Sec-Fetch-Site') == 'cross-site':
            self.send_json({'ok': False, 'error': 'Cross-site request rejected'}, 403)
            return False
        return True

    def log_message(self, format, *args):
        msg = "%s - - [%s] %s" % (
            self.address_string(),
            self.log_date_time_string(),
            format % args,
        )
        if sys.stderr is not None:
            try:
                sys.stderr.write(msg + "\n")
            except Exception:
                pass
        log(msg, logging.DEBUG)

    def do_GET(self):
        if not self.trusted_request():
            return
        global LAST_PING_TIME, HAS_CONNECTED
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/ping":
            LAST_PING_TIME = time.time()
            HAS_CONNECTED = True
            self.send_json({"ok": True})
            return
        elif path == "/api/health":
            self.send_json({"ok": True, 'application': 'TeamFlow', 'version': APP_VERSION, 'instance': INSTANCE_ID})
            return
        elif path == "/api/connection":
            try:
                self.send_json({'ok': True, **connection_status()})
            except Exception as exc:
                self.send_json({'ok': False, 'error': str(exc)}, 503)
            return
        elif path == "/api/data":
            LAST_PING_TIME = time.time()
            try:
                cloud = active_cloud()
                if cloud is not None:
                    data = cloud.snapshot()
                else:
                    wb_mgr = TaskManagerWorkbook(get_configured_workbook_path())
                    data = wb_mgr.read_all_data()
                    data['connection'] = {'mode': 'local', 'state': 'ready'}
                self.send_json(data)
            except PermissionError as pe:
                self.send_json({"ok": False, "error": str(pe)}, status=500)
            except Exception as e:
                log(f"Error in /api/data: {e}", logging.ERROR)
                self.send_json({"ok": False, "error": str(e)}, status=500)
            return
        elif path == "/api/workbook-info":
            LAST_PING_TIME = time.time()
            self.handle_workbook_info()
            return
        elif path == "/api/select-workbook":
            LAST_PING_TIME = time.time()
            self.handle_select_workbook()
            return
        elif path == "/api/open-excel":
            LAST_PING_TIME = time.time()
            try:
                self.handle_open_excel()
            except Exception as e:
                self.send_json({'ok': False, 'error': str(e)}, 500)
            return
        elif path == "/api/report.xlsx":
            self.handle_report_download(urllib.parse.parse_qs(parsed.query))
            return
        elif path == "/" or not path.startswith("/api/"):
            LAST_PING_TIME = time.time()
            HAS_CONNECTED = True
            self.serve_static(path)
            return
        else:
            self.send_json({"ok": False, "error": "נתיב לא מוכר"}, status=404)
            return

    def do_POST(self):
        if not self.trusted_request():
            return
        global LAST_PING_TIME, HAS_CONNECTED
        LAST_PING_TIME = time.time()
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/exit":
            self.send_json({"ok": True})
            def immediate_exit():
                time.sleep(0.1)
                log("Server shutdown requested via exit button.")
                # Finish active workbook transactions before stopping the server.
                with THREAD_LOCK:
                    self.server.shutdown()
            threading.Thread(target=immediate_exit, daemon=True).start()
            return

        if path == "/api/close":
            self.send_json({"ok": True})
            return

        body = self.read_body_json()
        if body is None:
            return

        try:
            if path == '/api/cloud/connect':
                connection = start_cloud(body.get('url'))
                self.send_json({'ok': True, **connection.status()}, 202)
                return
            if path == '/api/cloud/show':
                cloud = active_cloud()
                if cloud is None:
                    raise ValueError('אין חיבור Excel פעיל.')
                self.send_json(cloud.submit('show'))
                return
            if path == '/api/browse-workbook':
                self.handle_browse_workbook()
                return
            mutations = ('/api/task', '/api/delete', '/api/done', '/api/bulk', '/api/reorder', '/api/lists', '/api/comment')
            cloud = active_cloud() if path in mutations else None
            if cloud is not None:
                self.handle_cloud_mutation(cloud, path, body)
                return
            # Serialize the complete read-modify-write transaction, across processes too.
            with THREAD_LOCK:
                if path in mutations:
                    wb_mgr = TaskManagerWorkbook(get_configured_workbook_path())
                    with file_lock(wb_mgr.lock_path):
                        if body.get('workbook') and Path(body['workbook']).resolve() != wb_mgr.path:
                            self.send_json({'ok': False, 'error': 'קובץ האקסל הוחלף. רענן את הנתונים ונסה שוב.'}, 409)
                            return
                        if body.get('revision') and body['revision'] != workbook_revision(wb_mgr.path):
                            self.send_json({'ok': False, 'error': 'הנתונים השתנו מאז הטעינה. רענן ובדוק את השינויים לפני שמירה.'}, 409)
                            return
                        self.dispatch_post(path, wb_mgr, body)
                else:
                    self.dispatch_post(path, None, body)
        except WorkbookConflictError as e:
            self.send_json({'ok': False, 'error': str(e)}, status=409)
        except (ValueError, zipfile.BadZipFile) as e:
            self.send_json({'ok': False, 'error': str(e)}, status=400)
        except PermissionError as pe:
            self.send_json({"ok": False, "error": str(pe)}, status=500)
        except Exception as e:
            if type(e).__name__ in ('CloudConflict', 'CloudError'):
                self.send_json({'ok': False, 'error': str(e)}, 409 if type(e).__name__ == 'CloudConflict' else 503)
                return
            log(f"API Error ({path}): {e}", logging.ERROR)
            self.send_json({"ok": False, "error": str(e)}, status=500)

    def handle_cloud_mutation(self, cloud, path, body):
        from cloud_excel import CloudConflict, CLOUD_LIMIT
        data = cloud.snapshot()
        if body.get('workbook') != data['workbook']:
            raise CloudConflict('מקור הנתונים הוחלף. יש לפתוח מחדש את המשימה.')
        if path == '/api/bulk':
            self.handle_cloud_bulk(cloud, body)
            return
        if path == '/api/comment':
            event = self.comment_event(body)
            if event:
                cloud.submit('comment', body.get('id', ''), event)
                self.send_json({'ok': True})
            return
        if path == '/api/reorder':
            ids = body.get('ids') or []
            if not ids:
                self.send_json({'ok': False, 'error': 'חסר סדר משימות'}, 400)
                return
            cloud.submit('reorder', ids)
            self.send_json({'ok': True})
            return
        if path == '/api/lists':
            self.handle_save_lists(CloudTaskWorkbook(cloud, data, lists_revision=body.get('listsRevision')), body)
            return
        if path == '/api/task' and not body.get('id'):
            self.dispatch_post(path, CloudTaskWorkbook(cloud, data), body)
            return
        current = next((t for t in data['tasks'] if t['id'] == body.get('id')), None)
        if not current or not body.get('taskRevision') or body['taskRevision'] != current['taskRevision']:
            raise CloudConflict('המשימה השתנתה. הטיוטה נשארה בטופס; יש להשוות עם הגרסה העדכנית.')
        self.dispatch_post(path, CloudTaskWorkbook(cloud, data, body['id'], body['taskRevision']), body)

    def comment_event(self, body):
        """The history event for a comment, or None after sending a 400."""
        text = (body.get('text') or '').strip()
        if not body.get('id') or not text:
            self.send_json({'ok': False, 'error': 'חסר טקסט לתגובה'}, 400)
            return None
        if len(text) > 2000:
            self.send_json({'ok': False, 'error': 'התגובה ארוכה מדי (עד 2000 תווים).'}, 400)
            return None
        user = (body.get('user') or '').strip() or get_logged_in_excel_user()
        return {'action': 'תגובה', 'user': user, 'comment': text}

    def handle_cloud_bulk(self, cloud, body):
        """Bulk changes as one Excel operation per task, each checked against the latest data."""
        from cloud_excel import CloudConflict
        ids = list(dict.fromkeys(body.get("ids") or []))
        action = (body.get("action") or "").strip()
        value = (body.get("value") or "").strip()
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        if not ids:
            self.send_json({"ok": False, "error": "לא נבחרו משימות"}, status=400)
            return
        if action not in BULK_ACTIONS:
            self.send_json({"ok": False, "error": "פעולה לא מוכרת"}, status=400)
            return
        now = datetime.now()
        done, failed = 0, []
        for task_id in ids:
            task = next((t for t in cloud.snapshot()['tasks'] if t['id'] == task_id), None)
            if task is None:
                failed.append(task_id)
                continue
            history = {"user": user, "action": "מחיקת משימה" if action == "delete" else "פעולה מרובה",
                       "id": task_id, "title": task.get("title", ""),
                       "details": f"{action}: '{value}' ע\"י {user}"}
            try:
                if action == "delete":
                    cloud.submit('delete', task_id, task['taskRevision'], history)
                else:
                    changed = apply_bulk_action(copy.deepcopy(task), action, value, user,
                                                now.strftime("%Y-%m-%d %H:%M:%S"), now.strftime("%Y-%m-%d"))
                    cloud.submit('update', changed, task['taskRevision'], history)
                done += 1
            except CloudConflict:
                failed.append(task_id)
        self.send_json({"ok": True, "count": done, "failed": failed})

    def dispatch_post(self, path, wb_mgr, body):
        if path == "/api/task":
            self.handle_save_task(wb_mgr, body)
        elif path == "/api/delete":
            self.handle_delete_task(wb_mgr, body)
        elif path == "/api/done":
            self.handle_mark_done(wb_mgr, body)
        elif path == "/api/bulk":
            self.handle_bulk(wb_mgr, body)
        elif path == "/api/reorder":
            self.handle_reorder(wb_mgr, body)
        elif path == "/api/lists":
            self.handle_save_lists(wb_mgr, body)
        elif path == "/api/comment":
            event = self.comment_event(body)
            if event:
                wb_mgr.add_task_event(body.get('id', ''), event)
                self.send_json({'ok': True})
        elif path == "/api/set-workbook":
            self.handle_set_workbook(body)
        elif path == "/api/upload-workbook":
            self.handle_upload_workbook(body)
        elif path == "/api/reset-workbook":
            self.handle_reset_workbook()
        elif path == '/api/open-excel':
            self.handle_open_excel()
        else:
            self.send_json({"ok": False, "error": "נתיב לא מוכר"}, status=404)

    def handle_report_download(self, query):
        try:
            start, end = (query.get(k, [''])[0] for k in ('from', 'to'))
            if date.fromisoformat(start) > date.fromisoformat(end):
                raise ValueError('טווח התאריכים אינו תקין')
            cloud = active_cloud()
            data = cloud.snapshot() if cloud is not None else TaskManagerWorkbook(get_configured_workbook_path()).read_all_data()
            content = report_workbook_bytes(data, start, end)
        except ValueError as exc:
            self.send_json({'ok': False, 'error': str(exc) or 'טווח התאריכים אינו תקין'}, 400)
            return
        except Exception as exc:
            log(f'Report failed: {exc}', logging.ERROR)
            self.send_json({'ok': False, 'error': str(exc)}, 500)
            return
        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', f'attachment; filename="TeamFlow-report-{end}.xlsx"')
        self.send_header('Content-Length', str(len(content)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(content)

    def handle_open_excel(self):
        cloud = active_cloud()
        if cloud is not None:
            self.send_json(cloud.submit('show'))
            return
        wb_path = get_configured_workbook_path()
        if not wb_path.is_file():
            self.send_json({'ok': False, 'error': 'קובץ האקסל לא נמצא'}, 404)
            return
        os.startfile(wb_path)
        self.send_json({'ok': True, 'path': str(wb_path)})

    def read_body_json(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            if content_length < 0 or content_length > MAX_REQUEST_BYTES:
                self.send_json({'ok': False, 'error': 'הקובץ גדול מדי. הגודל המרבי הוא 20 MB.'}, 413)
                return None
            raw_body = self.rfile.read(content_length)
            data = json.loads(raw_body.decode("utf-8")) if raw_body else {}
            if not isinstance(data, dict):
                raise ValueError('Expected a JSON object')
            for key in ('id', 'title', 'owner', 'priority', 'status', 'category', 'dueDate',
                        'startDate', 'completedDate', 'blocked', 'notes', 'user', 'action',
                        'value', 'workbook', 'revision', 'path', 'filename', 'content', 'text'):
                if key in data and not isinstance(data[key], str):
                    raise ValueError('Expected text for ' + key)
            if 'ids' in data and (not isinstance(data['ids'], list) or
                                  any(not isinstance(tid, str) for tid in data['ids'])):
                raise ValueError('Expected a list of task IDs')
            return data
        except Exception:
            self.send_json({"ok": False, "error": "JSON לא תקין"}, status=400)
            return None

    def send_json(self, data, status=200):
        content = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        try:
            self.wfile.write(content)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            pass

    def serve_static(self, rel_path):
        if rel_path == "/":
            rel_path = "/index.html"
        clean_path = urllib.parse.unquote(rel_path).lstrip("/").replace("/", os.sep)
        file_path = (WEB_DIR / clean_path).resolve()

        # Prevent directory traversal
        try:
            file_path.relative_to(WEB_DIR.resolve())
        except ValueError:
            self.send_response(403)
            self.end_headers()
            return

        if not file_path.is_file():
            self.send_response(404)
            self.end_headers()
            return

        mime_types = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
            ".json": "application/json; charset=utf-8",
            ".png": "image/png",
            ".ttf": "font/ttf",
        }
        content_type = mime_types.get(file_path.suffix.lower(), "application/octet-stream")

        try:
            with open(file_path, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            log(f"Error serving static file {file_path}: {e}", logging.ERROR)
            self.send_response(500)
            self.end_headers()

    def handle_workbook_info(self):
        try:
            cloud = active_cloud()
            if cloud is not None:
                status = cloud.status()
                self.send_json({'ok': True, 'path': cloud.url,
                    'name': urllib.parse.unquote(cloud.url.rsplit('/', 1)[-1]),
                    'exists': status['state'] == 'ready', 'writable': status['state'] == 'ready',
                    'mode': 'excel', 'sizeBytes': 0, 'modified': status.get('lastReadAt'),
                    'isDefault': False, 'connection': status})
                return
            wb_path = get_configured_workbook_path()
            exists = wb_path.exists()
            size = wb_path.stat().st_size if exists else 0
            mod_time = datetime.fromtimestamp(wb_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S") if exists else ""
            default_wb = (BASE_DIR / "Team_Task_Manager.xlsx").resolve()

            writable = True
            if exists:
                try:
                    with open(str(wb_path), "a"):
                        pass
                except Exception:
                    writable = False

            self.send_json({
                "ok": True,
                "path": str(wb_path),
                "name": wb_path.name,
                "directory": str(wb_path.parent),
                "exists": exists,
                "sizeBytes": size,
                "modified": mod_time,
                "writable": writable,
                "isDefault": str(wb_path.resolve()) == str(default_wb.resolve()),
                "defaultPath": str(default_wb),
            })
        except Exception as e:
            log(f"Error in handle_workbook_info: {e}", logging.ERROR)
            self.send_json({"ok": False, "error": str(e)}, status=500)

    def handle_set_workbook(self, body: dict):
        raw_path = (body.get("path") or "").strip()
        raw_path = raw_path.strip('\"\'')
        if not raw_path:
            self.send_json({"ok": False, "error": "נא להזין נתיב מלא לקובץ אקסל"}, status=400)
            return

        if raw_path.startswith(('https://', 'ms-excel:')):
            connection = start_cloud(raw_path)
            self.send_json({'ok': True, **connection.status()}, 202)
            return
        target_path = Path(raw_path).resolve()
        if not target_path.exists():
            self.send_json({"ok": False, "error": f"הקובץ לא קיים בנתיב המבוקש:\n{target_path}"}, status=400)
            return

        if target_path.suffix.lower() != ".xlsx":
            self.send_json({"ok": False, "error": "הקובץ חייב להיות קובץ אקסל עם סיומת .xlsx"}, status=400)
            return

        # Check if user accidentally pointed into Program Files
        prog_files = os.environ.get("ProgramFiles", "C:\\Program Files").lower()
        prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)").lower()
        t_str = str(target_path).lower()
        if t_str.startswith(prog_files) or t_str.startswith(prog_files_x86):
            self.send_json({
                "ok": False,
                "error": "תיקיית Program Files נעולה לכתיבה ע\"י מערכת ההפעלה. יש לבחור קובץ בתיקיית המסמכים, שולחן העבודה, כונן משותף או תיקיית SharePoint."
            }, status=400)
            return

        if EXCEL_TEST_MODE:
            validate_workbook(target_path)
            connection = start_cloud(str(target_path), local_test=True)
            self.send_json({'ok': True, **connection.status()}, 202)
            return
        self.activate_workbook(target_path)

    def activate_workbook(self, target_path):
        validate_workbook(target_path)
        wb_mgr = TaskManagerWorkbook(target_path)
        full_data = wb_mgr.read_all_data()
        set_configured_workbook_path(target_path)
        log(f"Workbook successfully switched to: {target_path}")
        self.send_json({"ok": True, "path": str(target_path), "data": full_data})

    def handle_browse_workbook(self):
        if not PICKER_LOCK.acquire(blocking=False):
            self.send_json({'ok': False, 'error': 'חלון בחירת קובץ כבר פתוח.'}, 409)
            return
        try:
            selected = pick_workbook()
            if not selected:
                self.send_json({'ok': True, 'cancelled': True})
                return
            with THREAD_LOCK:
                self.handle_set_workbook({'path': selected})
        except subprocess.TimeoutExpired:
            self.send_json({'ok': False, 'error': 'תם הזמן לבחירת קובץ. נסה שוב.'}, 408)
        except Exception as e:
            log(f"PowerShell file picker error: {e}", logging.WARNING)
            self.send_json({'ok': False, 'error': str(e)}, 500)
        finally:
            PICKER_LOCK.release()

    def handle_upload_workbook(self, body: dict):
        filename = (body.get("filename") or "Team_Task_Manager.xlsx").strip()
        if not filename.lower().endswith(".xlsx"):
            raise ValueError('יש להעלות קובץ .xlsx בלבד')
        safe_name = "".join(c for c in filename if c.isalnum() or c in ("-", "_", ".", " ")).strip()
        if not safe_name:
            safe_name = "Team_Task_Manager.xlsx"

        b64_content = body.get("content")
        if not b64_content:
            self.send_json({"ok": False, "error": "לא התקבל תוכן קובץ"}, status=400)
            return

        try:
            file_bytes = base64.b64decode(b64_content, validate=True)
        except Exception:
            self.send_json({"ok": False, "error": "שגיאה בפענוח נתוני הקובץ"}, status=400)
            return

        if len(file_bytes) > MAX_UPLOAD_BYTES:
            self.send_json({'ok': False, 'error': 'הקובץ גדול מדי. הגודל המרבי הוא 20 MB.'}, 413)
            return
        upload_dir = BASE_DIR / 'uploads'
        upload_dir.mkdir(exist_ok=True)
        safe_name = safe_name[:120]
        dest_path = upload_dir / f'{uuid.uuid4().hex}_{safe_name}'
        if dest_path.suffix.lower() != '.xlsx':
            dest_path = dest_path.with_suffix('.xlsx')
        try:
            dest_path.write_bytes(file_bytes)
            self.activate_workbook(dest_path)
        except Exception:
            dest_path.unlink(missing_ok=True)
            raise

    def handle_reset_workbook(self):
        default_wb = (BASE_DIR / "Team_Task_Manager.xlsx").resolve()
        if not default_wb.exists() and APP_DIR != BASE_DIR:
            template_wb = APP_DIR / "Team_Task_Manager.xlsx"
            if template_wb.exists():
                try:
                    shutil.copy2(template_wb, default_wb)
                except Exception as e:
                    log(f"Could not copy template workbook: {e}", logging.WARNING)

        wb_mgr = TaskManagerWorkbook(default_wb)
        full_data = wb_mgr.read_all_data()
        set_configured_workbook_path(default_wb)
        log(f"Workbook reset to default: {default_wb}")
        self.send_json({"ok": True, "path": str(default_wb), "data": full_data})

    def handle_select_workbook(self):
        # Legacy fallback delegating to handle_browse_workbook
        self.handle_browse_workbook()

    def handle_save_task(self, wb_mgr: TaskManagerWorkbook, body: dict):
        title = (body.get("title") or "").strip()
        if not title:
            self.send_json({"ok": False, "error": "חסר שם משימה"}, status=400)
            return

        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]
        task_id = (body.get("id") or "").strip()
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()

        now_dt = datetime.now()
        now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
        today_date_str = now_dt.strftime("%Y-%m-%d")

        # Generate new ID if needed
        is_new = False
        if not task_id:
            is_new = True
            task_id = wb_mgr.new_task_id(tasks)
        else:
            if not any(t.get("id") == task_id for t in tasks):
                self.send_json({'ok': False, 'error': 'המשימה אינה קיימת עוד. רענן את הנתונים.'}, 404)
                return

        # Look for existing task to calculate change diffs
        existing_task = None
        for t in tasks:
            if t.get("id") == task_id:
                existing_task = t
                break

        # Defaults
        people = all_data["people"] or DEFAULT_PEOPLE
        statuses = all_data["statuses"] or DEFAULT_STATUSES
        categories = all_data["categories"] or DEFAULT_CATEGORIES
        priorities = all_data["priorities"] or DEFAULT_PRIORITIES

        owner = body.get('owner', (existing_task or {}).get('owner', people[0])).strip()
        status = (body.get("status") or "").strip() or statuses[0]
        category = body.get('category', (existing_task or {}).get('category', categories[0])).strip()
        priority = (body.get("priority") or "").strip() or priorities[0]
        due_date = format_date_str(body.get("dueDate"))
        previous = existing_task or {}
        start_date = format_date_str(body.get('startDate', previous.get('startDate'))) or today_date_str
        completed_date = format_date_str(body.get('completedDate', previous.get('completedDate')))
        for val in (due_date, start_date, completed_date):
            if val:
                date.fromisoformat(val)
        blocked = YES if body.get("blocked") == YES else NO
        notes = (body.get("notes") or "").strip()

        try:
            progress = int(round(float(body.get("progress", 0))))
        except Exception:
            progress = 0
        progress = max(0, min(100, progress))

        if status == DONE:
            progress = 100
            blocked = NO
            if not completed_date:
                completed_date = today_date_str
        else:
            completed_date = ''

        health, days_left = compute_health_and_days_left(status, blocked, due_date)

        if is_new or not existing_task:
            created_by = user
            created_at = today_date_str
            updated_by = user
            updated_at = now_str
            hist_action = "יצירת משימה"
            hist_details = f"נוצרה ע\"י {user}: {title} (אחראי: {owner}, סטטוס: {status}, עדיפות: {priority})"
        else:
            created_by = existing_task.get("createdBy") or user
            created_at = existing_task.get("createdAt") or today_date_str
            updated_by = user
            updated_at = now_str
            hist_action = "עדכון משימה"

            changes = []
            if existing_task.get("title") != title:
                changes.append(f"שם: '{existing_task.get('title')}' ➔ '{title}'")
            if existing_task.get("owner") != owner:
                changes.append(f"אחראי: '{existing_task.get('owner')}' ➔ '{owner}'")
            if existing_task.get("status") != status:
                changes.append(f"סטטוס: '{existing_task.get('status')}' ➔ '{status}'")
            if existing_task.get("priority") != priority:
                changes.append(f"עדיפות: '{existing_task.get('priority')}' ➔ '{priority}'")
            if existing_task.get("progress") != progress:
                changes.append(f"ביצוע: {existing_task.get('progress')}% ➔ {progress}%")
            if existing_task.get("dueDate") != due_date:
                changes.append(f"יעד: '{existing_task.get('dueDate')}' ➔ '{due_date}'")
            if existing_task.get("blocked") != blocked:
                changes.append(f"חסום: '{existing_task.get('blocked')}' ➔ '{blocked}'")

            if changes:
                hist_details = f"עודכן ע\"י {user}: " + ", ".join(changes)
            else:
                hist_details = f"עודכן ע\"י {user}: ללא שינוי שדות עיקריים"

        task_record = {
            "id": task_id,
            "title": title,
            "owner": owner,
            "priority": priority,
            "status": status,
            "category": category,
            "dueDate": due_date,
            "progress": progress,
            "blocked": blocked,
            "notes": notes,
            "startDate": start_date,
            "completedDate": completed_date,
            "health": health,
            "daysLeft": days_left,
            "createdBy": created_by,
            "createdAt": created_at,
            "updatedBy": updated_by,
            "updatedAt": updated_at,
        }

        # Update or append
        found = False
        for idx, t in enumerate(tasks):
            if t.get("id") == task_id:
                tasks[idx] = task_record
                found = True
                break
        if not found:
            tasks.append(task_record)

        wb_mgr.save_all_data(
            tasks,
            new_history_entry={
                "user": user,
                "action": hist_action,
                "id": task_id,
                "title": title,
                "details": hist_details,
            },
        )
        self.send_json({"ok": True, "task": task_record})

    def handle_delete_task(self, wb_mgr: TaskManagerWorkbook, body: dict):
        task_id = (body.get("id") or "").strip()
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]

        target = None
        new_tasks = []
        for t in tasks:
            if t.get("id") == task_id:
                target = t
            else:
                new_tasks.append(t)

        if target:
            wb_mgr.save_all_data(
                new_tasks,
                new_history_entry={
                    "user": user,
                    "action": "מחיקת משימה",
                    "id": task_id,
                    "title": target.get("title", ""),
                    "details": f"נמחקה ע\"י {user}: {target.get('title', '')}",
                },
            )
        self.send_json({"ok": bool(target), **({} if target else {'error': 'המשימה אינה קיימת עוד'})}, 200 if target else 404)

    def handle_mark_done(self, wb_mgr: TaskManagerWorkbook, body: dict):
        task_id = (body.get("id") or "").strip()
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today_str = datetime.now().strftime("%Y-%m-%d")

        target = None
        for t in tasks:
            if t.get("id") == task_id:
                t["status"] = DONE
                t["progress"] = 100
                t["completedDate"] = today_str
                t["blocked"] = NO
                t["health"] = HEALTH_CLOSED
                t["daysLeft"] = DONE
                t["updatedBy"] = user
                t["updatedAt"] = now_str
                target = t
                break

        if target:
            wb_mgr.save_all_data(
                tasks,
                new_history_entry={
                    "user": user,
                    "action": "סימון כבוצע",
                    "id": task_id,
                    "title": target.get("title", ""),
                    "details": f"סומנה כבוצעה (100%) ע\"י {user}",
                },
            )
        self.send_json({"ok": bool(target), **({} if target else {'error': 'המשימה אינה קיימת עוד'})}, 200 if target else 404)

    def handle_bulk(self, wb_mgr: TaskManagerWorkbook, body: dict):
        ids = body.get("ids") or []
        if not ids or not isinstance(ids, list):
            self.send_json({"ok": False, "error": "לא נבחרו משימות"}, status=400)
            return

        action = (body.get("action") or "").strip()
        val = (body.get("value") or "").strip()
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if action not in BULK_ACTIONS:
            self.send_json({"ok": False, "error": "פעולה לא מוכרת"}, status=400)
            return

        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]
        id_set = set(ids)
        today_str = datetime.now().strftime("%Y-%m-%d")

        if action == "delete":
            remaining_tasks = [t for t in tasks if t.get("id") not in id_set]
            deleted_count = len(tasks) - len(remaining_tasks)
            wb_mgr.save_all_data(
                remaining_tasks,
                new_history_entry={
                    "user": user,
                    "action": "מחיקה מרובה",
                    "id": f"{deleted_count} משימות",
                    "title": "מחיקה מרובה",
                    "details": f"נמחקו {deleted_count} משימות ע\"י {user}",
                },
            )
            self.send_json({"ok": True, "count": deleted_count})
            return

        affected_count = 0
        for t in tasks:
            if t.get("id") in id_set:
                affected_count += 1
                apply_bulk_action(t, action, val, user, now_str, today_str)

        wb_mgr.save_all_data(
            tasks,
            new_history_entry={
                "user": user,
                "action": "פעולה מרובה",
                "id": f"{affected_count} משימות",
                "title": f"פעולה: {action}",
                "details": f"עודכנו {affected_count} משימות בערך '{val}' ע\"י {user}",
            },
        )
        self.send_json({"ok": True, "count": affected_count})

    def handle_reorder(self, wb_mgr: TaskManagerWorkbook, body: dict):
        ids = body.get("ids") or []
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        if not ids or not isinstance(ids, list):
            self.send_json({"ok": False, "error": "חסר סדר משימות"}, status=400)
            return

        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]
        task_map = {t.get("id"): t for t in tasks}

        ordered_tasks = []
        for tid in ids:
            if tid in task_map:
                ordered_tasks.append(task_map.pop(tid))

        if not ordered_tasks:
            self.send_json({"ok": False, "error": "לא נמצאו משימות לסידור"}, status=400)
            return

        ordered_set = set(ids)
        for t in tasks:
            if t.get("id") not in ordered_set:
                ordered_tasks.append(t)

        wb_mgr.save_all_data(
            ordered_tasks,
            new_history_entry={
                "user": user,
                "action": "סידור ידני",
                "id": f"{len(ids)} משימות",
                "title": "סידור ידני בלוח",
                "details": f"סידור ידני של {len(ids)} משימות ע\"י {user}",
            },
        )
        self.send_json({"ok": True})

    def handle_save_lists(self, wb_mgr: TaskManagerWorkbook, body: dict):
        user = (body.get("user") or "").strip() or get_logged_in_excel_user()
        def parse_items(raw_val, defaults):
            if isinstance(raw_val, str):
                items = [x.strip() for x in raw_val.splitlines() if x.strip()]
            elif isinstance(raw_val, list):
                items = [str(x).strip() for x in raw_val if str(x).strip()]
            else:
                items = []
            return list(dict.fromkeys(items)) if items else list(defaults)

        people = parse_items(body.get("people"), DEFAULT_PEOPLE)
        categories = parse_items(body.get("categories"), DEFAULT_CATEGORIES)
        statuses = parse_items(body.get("statuses"), DEFAULT_STATUSES)
        priorities = parse_items(body.get("priorities"), DEFAULT_PRIORITIES)

        all_data = wb_mgr.read_all_data()
        tasks = all_data["tasks"]

        wb_mgr.save_all_data(
            tasks,
            lists={
                "people": people,
                "categories": categories,
                "statuses": statuses,
                "priorities": priorities,
            },
            new_history_entry={
                "user": user,
                "action": "עדכון רשימות",
                "id": "-",
                "title": "הגדרות צוות",
                "details": f"עודכנו רשימות ע\"י {user}: {len(people)} אנשים, {len(categories)} קטגוריות",
            },
        )
        full_data = wb_mgr.read_all_data()
        self.send_json({"ok": True, "data": full_data})

def existing_instance_url():
    """Only reuse the server published by the owner of this data directory's lock."""
    try:
        info = json.loads((BASE_DIR / '.instance.json').read_text(encoding='utf-8'))
        port = info['port']
        if type(port) is not int or not 8765 <= port <= 8784:
            return None
        connection = http.client.HTTPConnection('127.0.0.1', port, timeout=0.3)
        try:
            connection.request('GET', '/api/health')
            response = connection.getresponse()
            health = json.loads(response.read())
            if response.status == 200 and health.get('application') == 'TeamFlow' and health.get('instance') == info['instance']:
                return f'http://127.0.0.1:{port}'
        finally:
            connection.close()
    except (OSError, ValueError, KeyError, TypeError, http.client.HTTPException):
        pass
    return None


def main():
    # One server per data directory, including simultaneous double-clicks.
    # OS locks are released on a crash; a stale endpoint file cannot block startup.
    guard = file_lock(BASE_DIR / '.instance.lock', timeout=0)
    deadline = time.monotonic() + 15
    while True:
        try:
            guard.__enter__()
            break
        except PermissionError:
            url = existing_instance_url()
            if url:
                log(f'Reusing running application at {url}')
                if '--no-browser' not in sys.argv:
                    webbrowser.open(url)
                return
            if time.monotonic() >= deadline:
                raise RuntimeError('TeamFlow is already starting or shutting down. Try again shortly.')
            time.sleep(0.1)
            guard = file_lock(BASE_DIR / '.instance.lock', timeout=0)
    try:
        run_server()
    finally:
        (BASE_DIR / '.instance.json').unlink(missing_ok=True)
        guard.__exit__(None, None, None)


def run_server():
    global LAST_PING_TIME, SERVER_START_TIME
    SERVER_START_TIME = time.time()
    LAST_PING_TIME = time.time()

    log(f"Starting {APP_NAME}...")
    log(f"Base Directory: {BASE_DIR}")
    log(f"Web Directory: {WEB_DIR}")

    # Serve the UI first. Workbook creation/reading happens on /api/data so a
    # slow or unavailable shared drive does not prevent the browser from opening.

    # Find free port from 8765 to 8784
    server = None
    actual_port = PORT
    for p in range(8765, 8785):
        try:
            server = ThreadedHTTPServer(("127.0.0.1", p), TaskRequestHandler)
            actual_port = p
            break
        except OSError:
            continue

    if not server:
        err = f"Failed to bind any port in range 8765-8784."
        log(err, logging.FATAL)
        sys.exit(1)

    url = f"http://127.0.0.1:{actual_port}"
    log(f"Server listening on {url}")
    log(f"Launch: HTTP ready in {time.perf_counter() - PROCESS_STARTED:.3f}s")

    endpoint = BASE_DIR / '.instance.json'
    temporary = BASE_DIR / f'.instance-{INSTANCE_ID}.tmp'
    temporary.write_text(json.dumps({'port': actual_port, 'instance': INSTANCE_ID}), encoding='utf-8')
    os.replace(temporary, endpoint)

    # Automatically open browser
    def launch_browser():
        log(f"Opening browser at {url}")
        webbrowser.open(url)

    if '--no-browser' not in sys.argv:
        threading.Thread(target=launch_browser, daemon=True).start()

    try:
        server.serve_forever()
    except (KeyboardInterrupt, SystemExit):
        log("Server stopped.")
    finally:
        server.server_close()
        stop_cloud(wait=True)

if __name__ == "__main__":
    main()
