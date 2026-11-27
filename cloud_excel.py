"""Desktop Excel connection. COM objects never leave the owning STA thread.

Cloud mode deliberately supports existing-task edits only. Structural operations
need a separate allocation protocol; Excel coauthoring is not a distributed lock.
No local-file writes, browser cookies, Graph credentials, or cloud-sync claims.
"""
import copy
import hashlib
import json
import logging
import queue
import secrets
import threading
import time
import uuid
from concurrent.futures import Future, TimeoutError
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qs, unquote, quote

MAX_ROWS = 15000
# Unrecognised read errors are retried for this long before the connection is dropped.
RETRY_SECONDS = 60
# RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, VBA_E_IGNORE: Excel is editing or showing a dialog.
BUSY_CODES = {-2147418111, -2147417846, -2146777998}
FIELDS = ('title', 'owner', 'priority', 'status', 'category', 'dueDate',
          'progress', 'blocked', 'notes', 'startDate', 'completedDate', 'id',
          'daysLeft', 'health', 'dueMonth', 'createdBy', 'createdAt',
          'updatedBy', 'updatedAt')
CLOUD_LIMIT = 'הפעולה אינה זמינה בחיבור SharePoint. רענן ונסה שוב.'
# A created task is checked for this long, in case a simultaneous creation on another PC took its row.
VERIFY_SECONDS = 120
ID_ALPHABET = '23456789ABCDEFGHJKMNPQRSTUVWXYZ'
WRITES = ('update', 'create', 'delete', 'comment', 'lists', 'reorder')
ORDER_COLUMN = 21
LIST_COLUMNS = {'people': 1, 'categories': 6, 'statuses': 8, 'priorities': 10}


class CloudError(Exception):
    pass


class CloudConflict(CloudError):
    pass


def is_busy(exc):
    codes = [getattr(exc, 'hresult', None)] + list(exc.args[:1])
    if len(exc.args) > 2 and isinstance(exc.args[2], (tuple, list)) and exc.args[2]:
        codes.append(exc.args[2][-1])
    return any(isinstance(code, int) and code in BUSY_CODES for code in codes)


def normalize_url(value):
    if not isinstance(value, str) or len(value) > 4096:
        raise ValueError('יש להזין כתובת של קובץ Excel ב-SharePoint.')
    value = value.strip()
    if value.startswith('ms-excel:ofe|u|'):
        value = value[len('ms-excel:ofe|u|'):]
    parts = urlsplit(value)
    host = (parts.hostname or '').lower()
    if host in ('onedrive.live.com', '1drv.ms'):
        raise ValueError('זהו קישור שיתוף של OneDrive. פתח את הקובץ ב-Excel, בחר קובץ > מידע > העתק נתיב, והדבק כאן את הכתובת שמתחילה ב-https://d.docs.live.net.')
    # A personal OneDrive address (d.docs.live.net) is accepted for testing coauthoring without a company account.
    personal = host == 'd.docs.live.net'
    if (parts.scheme.lower() != 'https' or parts.username or parts.password
            or parts.port not in (None, 443)
            or not (personal or any(host.endswith(s) for s in ('.sharepoint.com', '.sharepoint.us', '.sharepoint.de', '.sharepoint.cn')))
            or any(ord(c) < 32 for c in value)):
        raise ValueError('נדרשת כתובת HTTPS של SharePoint ארגוני או של OneDrive.')
    path = unquote(parts.path)
    if path.startswith('/:x:/r/'):
        path = path[len('/:x:/r'):]
    query = parse_qs(parts.query)
    if not path.lower().endswith('.xlsx'):
        selected = (query.get('id') or query.get('RootFolder') or [''])[0]
        if selected.startswith('/') and selected.lower().endswith('.xlsx'):
            path = selected
        else:
            raise ValueError('זהו קישור שיתוף מקוצר. פתח את הקובץ ב-Excel, בחר קובץ > מידע > העתק נתיב, והדבק כאן את הכתובת המלאה המסתיימת ב-.xlsx.')
    if any(p in ('.', '..') for p in path.split('/')) or '\\' in path or '|' in path:
        raise ValueError('כתובת הקובץ אינה תקינה.')
    return urlunsplit(('https', host, quote(path, safe='/'), '', ''))


def is_tombstone(value):
    """Column T of a task deleted through SharePoint: its row stays, marked deleted."""
    if not isinstance(value, str) or '"deleted"' not in value:
        return False
    try:
        events = json.loads(value)
    except ValueError:
        return False
    return isinstance(events, list) and any(isinstance(e, dict) and e.get('deleted') is True for e in events)


def parse_events(value):
    """The per-task history list kept as JSON in column T; [] when empty or unreadable."""
    if not isinstance(value, str) or not value.strip():
        return []
    try:
        events = json.loads(value)
    except ValueError:
        return []
    return [e for e in events if isinstance(e, dict)] if isinstance(events, list) else []


def append_event(raw, event):
    """Column T with one more event, trimmed to fit an Excel cell. ValueError if T holds other data."""
    try:
        records = json.loads(raw or '[]')
    except (TypeError, ValueError):
        raise ValueError('עמודה T משמשת לנתונים אחרים. נדרשת עמודה T פנויה להיסטוריית TeamFlow.')
    if not isinstance(records, list):
        raise ValueError('עמודה T אינה זמינה להיסטוריית TeamFlow.')
    event = dict(event, time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # Comments are kept longer than field-change records.
    comments = [r for r in records if isinstance(r, dict) and r.get('comment')]
    changes = [r for r in records if not (isinstance(r, dict) and r.get('comment'))]
    keep = set(map(id, comments[-30:] + changes[-15:]))
    records = [r for r in records if id(r) in keep] + [event]
    encoded = json.dumps(records, ensure_ascii=False)
    while len(encoded) > 30000 and len(records) > 1:
        records.pop(0)
        encoded = json.dumps(records, ensure_ascii=False)
    if len(encoded) > 32767:
        raise ValueError('רשומת ההיסטוריה גדולה מדי.')
    return encoded


def task_revision(task):
    # Derived calendar values are not editable task content.
    record = {k: task.get(k) for k in FIELDS if k not in ('daysLeft', 'health', 'dueMonth')}
    return hashlib.sha256(json.dumps(record, ensure_ascii=False, sort_keys=True, default=str).encode()).hexdigest()


class GridCell:
    def __init__(self, value):
        self.value = value


class GridSheet:
    def __init__(self, rows):
        self.rows = rows
        self.max_row = len(rows)

    def cell(self, row, column):
        try:
            return GridCell(self.rows[row - 1][column - 1])
        except IndexError:
            return GridCell(None)


class GridWorkbook:
    def __init__(self, sheets):
        self.sheets = sheets
        self.sheetnames = list(sheets)

    def __getitem__(self, key):
        return self.sheets[key]

    def close(self):
        pass


class ExcelConnection:
    def __init__(self, url, decode, schema, recovery_dir, on_ready=None, local_test=False):
        # Test mode drives the same Excel code against a local copy: no SharePoint, no AutoSave.
        self.local_test = local_test
        self.url = str(Path(url).resolve()) if local_test else normalize_url(url)
        self.decode, self.schema = decode, schema
        self.recovery_dir = Path(recovery_dir)
        self.on_ready = on_ready
        self._lock = threading.RLock()
        self._queue = queue.Queue()
        self._stop = threading.Event()
        self._data = None
        self._sheets = set()
        self._failing_since = None
        self._writing = False
        self._last_read = 0
        self._rows = {}
        self._next_row = 7
        self._pending_creates = []
        self._state = dict(mode='excel', state='connecting', url=self.url,
                           message='Excel מתחבר לקובץ. אם מופיע חלון Microsoft, יש להשלים כניסה.',
                           cloudConfirmed=False, lastReadAt=None, autosave=False)
        self._thread = threading.Thread(target=self._run, name='TeamFlow-Excel', daemon=True)
        self._thread.start()

    def status(self):
        with self._lock:
            return copy.deepcopy(self._state)

    def _status(self, **values):
        with self._lock:
            self._state.update(values)

    def snapshot(self):
        with self._lock:
            if self._state['state'] != 'ready' or self._data is None:
                raise CloudError(self._state['message'])
            data = copy.deepcopy(self._data)
            data['connection'] = copy.deepcopy(self._state)
            return data

    def stop(self):
        # Do not discard a workbook, pending AutoSave, or the user's other Excel windows.
        self._stop.set()

    def submit(self, operation, *args):
        if self.status()['state'] != 'ready':
            raise CloudError(self.status()['message'])
        with self._lock:
            if operation in WRITES and self._state.get('writePending'):
                raise CloudError('Excel עדיין מטפל בשמירה קודמת. יש להמתין לסיום.')
            if operation in WRITES:
                self._state['writePending'] = True
        future = Future()
        self._queue.put((future, operation, args))
        try:
            return future.result(timeout=35)
        except TimeoutError:
            if future.cancel():
                self._status(writePending=False)
                raise CloudError('Excel עסוק. הפעולה בוטלה לפני שהתחילה; נסה שוב לאחר סגירת חלונות Excel.')
            raise CloudError('Excel עדיין מטפל בפעולה. אין לשלוח אותה שוב. בדוק את הקובץ ב-Excel לפני ניסיון נוסף.')

    def _run(self):
        excel = workbook = pythoncom = None
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx('Excel.Application')
            # This is an interactive desktop session: keep authentication/recovery visible.
            # Test mode stays hidden unless the user asks to see Excel.
            excel.Visible = not self.local_test
            excel.AutomationSecurity = 3  # Force-disable macros for this dedicated instance.
            if self.local_test:
                excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(self.url, UpdateLinks=0, ReadOnly=False, AddToMru=False)
            if self.local_test:
                same = str(Path(str(workbook.FullName)).resolve()).casefold() == self.url.casefold()
            else:
                same = normalize_url(str(workbook.FullName)).casefold() == self.url.casefold()
            if not same:
                raise CloudError('Excel פתח קובץ אחר מהכתובת המבוקשת. החיבור לא הופעל.')
            if workbook.ReadOnly:
                raise CloudError('הקובץ נפתח לקריאה בלבד. בדוק הרשאת עריכה ב-SharePoint.')
            if not self.local_test:
                if not workbook.AutoSaveOn:
                    workbook.AutoSaveOn = True
                if not workbook.AutoSaveOn:
                    raise CloudError('AutoSave אינו זמין. נדרשת גרסת Microsoft 365 התומכת בעריכה משותפת.')
            self.excel, self.workbook = excel, workbook
            self._read()
            if self.on_ready:
                self.on_ready(self)
            try:
                excel.WindowState = -4140  # Minimize only our dedicated Excel instance.
            except Exception:
                logging.warning('Could not minimize Excel', exc_info=True)
            while not self._stop.is_set() or not self._queue.empty():
                pythoncom.PumpWaitingMessages()
                try:
                    future, operation, args = self._queue.get(timeout=0.15)
                except queue.Empty:
                    if time.monotonic() - self._last_read >= 2:
                        self._refresh()
                    continue
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    self._writing = False
                    if operation == 'show':
                        excel.Visible = True
                        excel.WindowState = -4143
                        result = {'ok': True}
                    elif operation == 'update':
                        result = self._update(*args)
                    elif operation == 'create':
                        result = self._create(*args)
                    elif operation == 'delete':
                        result = self._delete(*args)
                    elif operation == 'comment':
                        result = self._comment(*args)
                    elif operation == 'lists':
                        result = self._lists(*args)
                    elif operation == 'reorder':
                        result = self._reorder(*args)
                    elif operation == 'run' and self.local_test:
                        result = args[0](self)  # Test hook: runs on the Excel thread.
                    else:
                        raise ValueError('Unknown Excel operation')
                    self._status(writePending=False)
                    future.set_result(result)
                except CloudConflict as exc:
                    self._status(writePending=False)
                    future.set_exception(exc)
                except Exception as exc:
                    if not self._writing and not isinstance(exc, CloudError):
                        # Nothing was written, so keep the connection; the refresh will judge it.
                        logging.warning('Excel operation failed before writing', exc_info=True)
                        self._status(writePending=False)
                        future.set_exception(CloudError('Excel עסוק ולא בוצעו שינויים. סיים את העריכה או סגור את החלון ב-Excel ונסה שוב.'))
                        continue
                    future.set_exception(exc)
                    raise
        except Exception as exc:
            message = str(exc)
            if not isinstance(exc, (CloudError, ValueError)):
                message = 'לא ניתן להמשיך בחיבור ל-Excel. בדוק את חלון Excel, הכניסה לחשבון ו-AutoSave. הפרטים נשמרו ביומן.'
                logging.exception('Excel connection failed')
            else:
                logging.error('Excel connection stopped: %s', exc)
            self._status(state='error', message=message)
        finally:
            while True:
                try:
                    future, _, _ = self._queue.get_nowait()
                    if not future.done():
                        future.set_exception(CloudError('החיבור ל-Excel הסתיים.'))
                except queue.Empty:
                    break
            # Leave Office open for its own save/recovery UI; never call Quit/Close on cloud data.
            # A local test copy is saved after every change, so its private Excel can close.
            if self.local_test and excel is not None:
                try:
                    if workbook is not None:
                        workbook.Close(SaveChanges=False)
                    excel.Quit()
                except Exception:
                    logging.warning('Could not close test Excel', exc_info=True)
                excel = None
            if excel is not None:
                try:
                    excel.Visible = True
                except Exception:
                    pass
            self.workbook = self.excel = None
            workbook = excel = None
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _refresh(self):
        """Background read that tolerates Excel being temporarily unavailable."""
        try:
            self._read()
            self._verify_creates()
        except CloudError:
            raise  # Read-only, AutoSave off, bad template: retrying cannot help.
        except Exception as exc:
            now = time.monotonic()
            if self._failing_since is None:
                self._failing_since = now
                logging.warning('Excel read failed; retrying', exc_info=True)
            if is_busy(exc):
                self._failing_since = now  # A busy Excel never counts toward the limit.
            elif now - self._failing_since >= RETRY_SECONDS:
                raise
            self._last_read = now
            self._status(busy=True, message='Excel עסוק כרגע. החיבור נשמר וינסה שוב אוטומטית.')

    def _read(self):
        workbook = self.workbook
        if workbook.ReadOnly or not (self.local_test or workbook.AutoSaveOn):
            raise CloudError('העריכה נעצרה: Excel במצב קריאה בלבד או AutoSave כבוי. בדוק ב-Excel והתחבר מחדש.')
        sheets, row_ids, audit, last_row = {}, {}, [], 6
        task_sheet = self.schema['task']
        for name, columns in ((task_sheet, ORDER_COLUMN), (self.schema['lists'], 10), (self.schema['history'], 6)):
            try:
                sheet = workbook.Worksheets(name)
            except Exception as exc:
                # A sheet that was there before is unavailable, not missing.
                if name in self._sheets or is_busy(exc):
                    raise
                if name == task_sheet:
                    raise CloudError('חסר גיליון משימות בתבנית TeamFlow.')
                continue
            end = int(sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1)
            if end > MAX_ROWS:
                raise CloudError('הקובץ חורג ממגבלת 15,000 השורות של החיבור הניסיוני.')
            rows = [list(r) for r in sheet.Range(sheet.Cells(1, 1), sheet.Cells(max(end, 6), columns)).Value2]
            if name == task_sheet:
                if rows[5][:12] != self.schema['headers'][:12]:
                    raise CloudError('כותרות המשימות אינן בתבנית TeamFlow.')
                epoch = datetime(1904, 1, 1) if workbook.Date1904 else datetime(1899, 12, 30)
                for n, row in enumerate(rows[6:], 7):
                    if any(v not in (None, '') for v in row[:20]):
                        last_row = n
                    if row[0] and not row[11]:
                        raise CloudError('למשימה חסר מזהה קבוע בעמודה L. יש להשלים מזהים ייחודיים לפני חיבור משותף.')
                    if row[11]:
                        tid = str(row[11]).strip()
                        if tid in row_ids:
                            raise CloudError('נמצאו מזהי משימה כפולים. החיבור נעצר כדי למנוע עריכה בשורה שגויה.')
                        row_ids[tid] = n
                    for c in (5, 9, 10, 16):
                        if isinstance(row[c], (int, float)):
                            row[c] = (epoch + timedelta(days=row[c])).strftime('%Y-%m-%d')
                    if row[19]:
                        try:
                            events = json.loads(row[19])
                            if isinstance(events, list):
                                audit.extend(e for e in events if isinstance(e, dict))
                        except (ValueError, TypeError):
                            pass
            sheets[name] = GridSheet(rows)
        fingerprint = hashlib.sha256(json.dumps([s.rows for s in sheets.values()], default=str).encode()).hexdigest()
        data = self.decode(GridWorkbook(sheets), self.url, fingerprint)
        for task in data['tasks']:
            task['taskRevision'] = task_revision(task)
        data['history'] = sorted(data['history'] + audit, key=lambda e: str(e.get('time', '')), reverse=True)[:60]
        data['capabilities'] = {'create': True, 'delete': True, 'reorder': True, 'lists': True}
        self._rows = row_ids
        self._next_row = last_row + 1
        self._last_read = time.monotonic()
        with self._lock:
            self._data = data
            self._sheets = set(sheets)
            self._failing_since = None
            self._state.update(state='ready', autosave=True, busy=False,
                               lastReadAt=datetime.now().isoformat(timespec='seconds'),
                               test=self.local_test,
                               message='מצב בדיקה: מחובר ל-Excel עם קובץ מקומי.' if self.local_test else
                               'מחובר ל-Excel • AutoSave פעיל. אישור העלאה ל-SharePoint זמין ב-Excel.')
        return data

    # ---------------------------------------------------------------- writes
    def new_task_id(self):
        """Random IDs cannot clash when two PCs create tasks at the same moment."""
        while True:
            candidate = 'T-' + ''.join(secrets.choice(ID_ALPHABET) for _ in range(5))
            if candidate not in self._rows:
                return candidate

    def _cell_value(self, key, value):
        if key == 'progress':
            return float(value or 0) / 100
        if key in ('dueDate', 'startDate', 'completedDate') and value:
            epoch = datetime(1904, 1, 1) if self.workbook.Date1904 else datetime(1899, 12, 30)
            return (datetime.strptime(value, '%Y-%m-%d') - epoch).days
        if value is None:
            return ''
        if isinstance(value, str) and (len(value) > 32767 or any(ord(c) < 32 and c not in '\t\n\r' for c in value)):
            raise CloudConflict('הטקסט ארוך מדי או מכיל תווים שאינם נתמכים ב-Excel.')
        return value

    @staticmethod
    def _writable(sheet, row, col):
        cell = sheet.Cells(row, col)
        if col in (13, 14, 15) and cell.HasFormula:
            return False  # Calculated by the workbook itself.
        if cell.HasFormula or (sheet.ProtectContents and cell.Locked):
            raise CloudConflict('אחד השדות מכיל נוסחה או מוגן. לא בוצעו שינויים; בדוק את תבנית הקובץ.')
        return True

    def _field_patches(self, sheet, row, task, before):
        """Cells to write: changed fields for an existing row, filled fields for a new one."""
        fields = dict(task, dueMonth=(task.get('dueDate') or '')[:7])
        old = None if before is None else dict(before, dueMonth=(before.get('dueDate') or '')[:7])
        patches = []
        for col, key in enumerate(FIELDS, 1):
            if old is not None and (key == 'id' or fields.get(key) == old.get(key)):
                continue
            if old is None and fields.get(key) in (None, ''):
                continue
            value = self._cell_value(key, fields.get(key, ''))
            if self._writable(sheet, row, col):
                patches.append((col, value))
        return patches

    @staticmethod
    def _history_cell(sheet, row, event):
        # A per-task history cell avoids a shared append position across different tasks.
        cell = sheet.Cells(row, 20)
        if cell.HasFormula or (sheet.ProtectContents and cell.Locked):
            raise CloudConflict('עמודת ההיסטוריה T מוגנת או מכילה נוסחה.')
        try:
            return append_event(cell.Value2, event)
        except ValueError as exc:
            raise CloudConflict(str(exc))

    def _commit(self, sheet, row, patches, recovery):
        return self._commit_cells(sheet, [(row, col, val) for col, val in patches], dict(recovery, row=row))

    def _commit_cells(self, sheet, cells, recovery):
        # Retain intent before handing changes to Office. No automatic retry after a partial COM failure.
        self.recovery_dir.mkdir(parents=True, exist_ok=True)
        journal = self.recovery_dir / (uuid.uuid4().hex + '.json')
        recovery = dict(recovery, url=self.url, status='pending-in-excel', time=datetime.now().isoformat())
        journal.write_text(json.dumps(recovery, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        self._writing = True
        for row, col, val in cells:
            cell = sheet.Cells(row, col)
            if val is None:
                cell.ClearContents()
                continue
            if isinstance(val, str):
                cell.NumberFormat = '@'
            elif col == 7:
                cell.NumberFormat = '0%'
            elif col in (6, 10, 11):
                cell.NumberFormat = 'yyyy-mm-dd'
            cell.Value2 = val
        self.workbook.Save()
        recovery['status'] = 'saved-in-excel-cloud-not-confirmed'
        journal.write_text(json.dumps(recovery, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        self._writing = False
        try:
            return self._read()
        except CloudError:
            raise
        except Exception:
            logging.warning('Saved, but the follow-up read failed', exc_info=True)
            return None  # The background refresh will catch up.

    @staticmethod
    def _check_cell(sheet, row, col):
        cell = sheet.Cells(row, col)
        if cell.HasFormula or (sheet.ProtectContents and cell.Locked):
            raise CloudConflict('תא בקובץ מכיל נוסחה או מוגן, ולכן לא בוצעו שינויים.')

    def _current(self, data, task_id, expected_revision):
        existing = next((t for t in data['tasks'] if t['id'] == task_id), None)
        if not existing or existing['taskRevision'] != expected_revision:
            raise CloudConflict('המשימה השתנתה מאז פתיחת הטופס. הטיוטה נשמרה במסך; פתח מחדש את המשימה לאחר העתקת השינויים.')
        return existing

    def _update(self, task, expected_revision, history):
        self._writing = False
        existing = self._current(self._read(), task['id'], expected_revision)
        row = self._rows[task['id']]
        sheet = self.workbook.Worksheets(self.schema['task'])
        patches = self._field_patches(sheet, row, task, existing)
        if history:
            patches.append((20, self._history_cell(sheet, row, history)))
        fresh = self._commit(sheet, row, patches, dict(kind='update', task=task, previous=existing))
        return {'ok': True, 'data': fresh, 'cloudConfirmed': False}

    def _create(self, task, history, verify=True):
        """Write a new task into the first empty row after the last used one."""
        self._writing = False
        self._read()
        task = dict(task)
        if not task.get('id') or task['id'] in self._rows:
            task['id'] = self.new_task_id()
        row = self._next_row
        if row > MAX_ROWS:
            raise CloudConflict('הקובץ הגיע למגבלת השורות של החיבור.')
        sheet = self.workbook.Worksheets(self.schema['task'])
        values = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, 20)).Value2[0]
        if any(v not in (None, '') for v in values):
            raise CloudConflict('השורה הבאה בקובץ אינה ריקה. רענן ונסה שוב.')
        patches = self._field_patches(sheet, row, task, None)
        if history:
            patches.append((20, self._history_cell(sheet, row, dict(history, id=task['id']))))
        fresh = self._commit(sheet, row, patches, dict(kind='create', task=task))
        if verify:
            self._pending_creates.append(dict(task=task, history=history, since=time.monotonic(), attempts=1))
        return {'ok': True, 'data': fresh, 'task': task, 'cloudConfirmed': False}

    def _delete(self, task_id, expected_revision, history):
        """Empty the row but keep its ID, marked deleted, so no other row moves."""
        self._writing = False
        existing = self._current(self._read(), task_id, expected_revision)
        row = self._rows[task_id]
        sheet = self.workbook.Worksheets(self.schema['task'])
        values = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, 19)).Value2[0]
        patches = [(col, None) for col, value in enumerate(values, 1)
                   if col != 12 and value not in (None, '') and self._writable(sheet, row, col)]
        patches.append((20, self._history_cell(sheet, row, dict(history or {}, id=task_id, deleted=True))))
        fresh = self._commit(sheet, row, patches, dict(kind='delete', previous=existing))
        return {'ok': True, 'data': fresh, 'cloudConfirmed': False}

    def _comment(self, task_id, event):
        """Only the task's own history cell changes, so a comment never conflicts with field edits."""
        self._writing = False
        data = self._read()
        if not any(t['id'] == task_id for t in data['tasks']):
            raise CloudConflict('המשימה אינה קיימת עוד. רענן את הנתונים.')
        row = self._rows[task_id]
        sheet = self.workbook.Worksheets(self.schema['task'])
        patches = [(20, self._history_cell(sheet, row, dict(event, id=task_id)))]
        fresh = self._commit(sheet, row, patches, dict(kind='comment', id=task_id, event=event))
        return {'ok': True, 'data': fresh, 'cloudConfirmed': False}

    def _lists(self, lists, expected_revision):
        """Rewrite the team lists, refusing if another PC changed them since this PC loaded them."""
        self._writing = False
        data = self._read()
        if expected_revision and expected_revision != data.get('listsRevision'):
            raise CloudConflict('רשימות הצוות שונו בינתיים במחשב אחר. השינויים שלך נשארו בטופס; רענן והשווה לפני שמירה.')
        name = self.schema['lists']
        try:
            sheet = self.workbook.Worksheets(name)
        except Exception:
            sheet = self.workbook.Worksheets.Add(After=self.workbook.Worksheets(self.workbook.Worksheets.Count))
            sheet.Name = name
            for col, title in ((1, 'שם'), (2, 'צוות'), (3, 'אימייל'), (4, 'פעיל'), (6, 'קטגוריה'), (8, 'סטטוס'), (10, 'קוד')):
                sheet.Cells(5, col).Value2 = title
        end = max(6, int(sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1))
        grid = [list(r) for r in sheet.Range(sheet.Cells(1, 1), sheet.Cells(end, 10)).Value2]
        # Team, email and active stay with each person's name, as in local mode.
        details = {str(r[0]).strip(): r[1:4] for r in grid[5:] if r[0] not in (None, '')}
        wanted = {}
        for key, col in LIST_COLUMNS.items():
            for i, value in enumerate(lists.get(key) or []):
                wanted[(6 + i, col)] = value
                if key == 'people':
                    for offset, extra in enumerate(details.get(value, (None, None, 'כן')), 2):
                        wanted[(6 + i, offset)] = extra
        cells = []
        for row in range(6, max(end, 6 + max(len(v or []) for v in lists.values())) + 1):
            for col in (1, 2, 3, 4, 6, 8, 10):
                new = wanted.get((row, col))
                old = grid[row - 1][col - 1] if row <= len(grid) else None
                if (new in (None, '') and old in (None, '')) or new == old:
                    continue
                self._check_cell(sheet, row, col)
                cells.append((row, col, None if new in (None, '') else new))
        if not cells:
            return {'ok': True, 'data': data}
        fresh = self._commit_cells(sheet, cells, dict(kind='lists', lists=lists))
        return {'ok': True, 'data': fresh}

    def _reorder(self, ids):
        """Manual order through column U. Moving one card writes one number: halfway between its neighbours."""
        self._writing = False
        data = self._read()
        current = [t['id'] for t in data['tasks']]
        wanted = [i for i in dict.fromkeys(ids) if i in self._rows]
        wanted += [i for i in current if i not in wanted]
        if wanted == current:
            return {'ok': True, 'data': data}
        key = {t['id']: (t['order'] if t.get('order') is not None else self._rows[t['id']]) for t in data['tasks']}
        moved = [i for i in wanted if [x for x in wanted if x != i] == [x for x in current if x != i]]
        updates = {}
        if len(moved) >= 1:
            task_id = moved[0]
            at = wanted.index(task_id)
            before = key[wanted[at - 1]] if at > 0 else None
            after = key[wanted[at + 1]] if at + 1 < len(wanted) else None
            if before is None:
                updates[task_id] = after - 1
            elif after is None:
                updates[task_id] = before + 1
            elif after - before > 1e-6:
                updates[task_id] = (before + after) / 2
        if not updates:
            # Several cards moved, or the gap became too small: number every card again.
            updates = {task_id: float(n) for n, task_id in enumerate(wanted, 1) if key[task_id] != n}
        sheet = self.workbook.Worksheets(self.schema['task'])
        cells = []
        header = sheet.Cells(6, ORDER_COLUMN)
        if header.Value2 in (None, ''):
            cells.append((6, ORDER_COLUMN, 'סדר'))
        for task_id, value in updates.items():
            self._check_cell(sheet, self._rows[task_id], ORDER_COLUMN)
            cells.append((self._rows[task_id], ORDER_COLUMN, value))
        fresh = self._commit_cells(sheet, cells, dict(kind='reorder', order=wanted))
        return {'ok': True, 'data': fresh}

    def _verify_creates(self):
        """Re-add a created task if a simultaneous creation on another PC took its row."""
        now = time.monotonic()
        still_checking = []
        for pending in self._pending_creates:
            task = pending['task']
            if now - pending['since'] > VERIFY_SECONDS:
                continue
            if task['id'] in self._rows:
                still_checking.append(pending)  # Present, or deleted on purpose by someone.
                continue
            if pending['attempts'] >= 3:
                logging.error('Created task %s keeps disappearing; giving up', task['id'])
                continue
            logging.warning('Created task %s was lost in a merge; adding it again', task['id'])
            self._create(task, pending['history'], verify=False)
            pending.update(attempts=pending['attempts'] + 1, since=now)
            still_checking.append(pending)
        self._pending_creates = still_checking
