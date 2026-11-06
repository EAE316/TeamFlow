"""Desktop Excel connection. COM objects never leave the owning STA thread.

Cloud mode deliberately supports existing-task edits only. Structural operations
need a separate allocation protocol; Excel coauthoring is not a distributed lock.
No local-file writes, browser cookies, Graph credentials, or cloud-sync claims.
"""
import copy
import hashlib
import json
import queue
import threading
import time
import uuid
from concurrent.futures import Future, TimeoutError
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qs, unquote, quote

MAX_ROWS = 15000
FIELDS = ('title', 'owner', 'priority', 'status', 'category', 'dueDate',
          'progress', 'blocked', 'notes', 'startDate', 'completedDate', 'id',
          'daysLeft', 'health', 'dueMonth', 'createdBy', 'createdAt',
          'updatedBy', 'updatedAt')
CLOUD_LIMIT = 'בחיבור SharePoint הניסיוני ניתן לערוך משימות קיימות. יצירה, מחיקה, סידור שורות ושינוי רשימות אינם זמינים כדי למנוע התנגשויות בין מחשבים.'


class CloudError(Exception):
    pass


class CloudConflict(CloudError):
    pass


def normalize_url(value):
    if not isinstance(value, str) or len(value) > 4096:
        raise ValueError('יש להזין כתובת של קובץ Excel ב-SharePoint.')
    value = value.strip()
    if value.startswith('ms-excel:ofe|u|'):
        value = value[len('ms-excel:ofe|u|'):]
    parts = urlsplit(value)
    host = (parts.hostname or '').lower()
    if (parts.scheme.lower() != 'https' or parts.username or parts.password
            or parts.port not in (None, 443)
            or not any(host.endswith(s) for s in ('.sharepoint.com', '.sharepoint.us', '.sharepoint.de', '.sharepoint.cn'))
            or any(ord(c) < 32 for c in value)):
        raise ValueError('נדרשת כתובת HTTPS של SharePoint ארגוני.')
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
    def __init__(self, url, decode, schema, recovery_dir, on_ready=None):
        self.url = normalize_url(url)
        self.decode, self.schema = decode, schema
        self.recovery_dir = Path(recovery_dir)
        self.on_ready = on_ready
        self._lock = threading.RLock()
        self._queue = queue.Queue()
        self._stop = threading.Event()
        self._data = None
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
            if operation == 'update' and self._state.get('writePending'):
                raise CloudError('Excel עדיין מטפל בשמירה קודמת. יש להמתין לסיום.')
            if operation == 'update':
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
            excel.Visible = True
            excel.AutomationSecurity = 3  # Force-disable macros for this dedicated instance.
            workbook = excel.Workbooks.Open(self.url, UpdateLinks=0, ReadOnly=False, AddToMru=False)
            if normalize_url(str(workbook.FullName)).casefold() != self.url.casefold():
                raise CloudError('Excel פתח קובץ אחר מהכתובת המבוקשת. החיבור לא הופעל.')
            if workbook.ReadOnly:
                raise CloudError('הקובץ נפתח לקריאה בלבד. בדוק הרשאת עריכה ב-SharePoint.')
            if not workbook.AutoSaveOn:
                workbook.AutoSaveOn = True
            if not workbook.AutoSaveOn:
                raise CloudError('AutoSave אינו זמין. נדרשת גרסת Microsoft 365 התומכת בעריכה משותפת.')
            self.excel, self.workbook = excel, workbook
            self._read()
            if self.on_ready:
                self.on_ready(self)
            excel.WindowState = -4140  # Minimize only our dedicated Excel instance.
            while not self._stop.is_set() or not self._queue.empty():
                pythoncom.PumpWaitingMessages()
                try:
                    future, operation, args = self._queue.get(timeout=0.15)
                except queue.Empty:
                    if time.monotonic() - self._last_read >= 2:
                        self._read()
                    continue
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    if operation == 'show':
                        excel.Visible = True
                        excel.WindowState = -4143
                        result = {'ok': True}
                    elif operation == 'update':
                        result = self._update(*args)
                    else:
                        raise ValueError('Unknown Excel operation')
                    self._status(writePending=False)
                    future.set_result(result)
                except CloudConflict as exc:
                    self._status(writePending=False)
                    future.set_exception(exc)
                except Exception as exc:
                    future.set_exception(exc)
                    raise
        except Exception as exc:
            message = str(exc)
            if not isinstance(exc, (CloudError, ValueError)):
                message = 'לא ניתן להמשיך בחיבור ל-Excel. בדוק את חלון Excel, הכניסה לחשבון ו-AutoSave. הפרטים נשמרו ביומן.'
                import logging
                logging.exception('Excel connection failed')
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
            if excel is not None:
                try:
                    excel.Visible = True
                except Exception:
                    pass
            self.workbook = self.excel = None
            workbook = excel = None
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    def _read(self):
        workbook = self.workbook
        if workbook.ReadOnly or not workbook.AutoSaveOn:
            raise CloudError('העריכה נעצרה: Excel במצב קריאה בלבד או AutoSave כבוי. בדוק ב-Excel והתחבר מחדש.')
        sheets, row_ids, audit = {}, {}, []
        task_sheet = self.schema['task']
        for name, columns in ((task_sheet, 20), (self.schema['lists'], 10), (self.schema['history'], 6)):
            try:
                sheet = workbook.Worksheets(name)
            except Exception:
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
        data['capabilities'] = {'existingTasksOnly': True}
        self._rows = row_ids
        self._last_read = time.monotonic()
        with self._lock:
            self._data = data
            self._state.update(state='ready', autosave=True,
                               lastReadAt=datetime.now().isoformat(timespec='seconds'),
                               message='מחובר ל-Excel • AutoSave פעיל. אישור העלאה ל-SharePoint זמין ב-Excel.')
        return data

    def _update(self, task, expected_revision, history):
        data = self._read()
        existing = next((t for t in data['tasks'] if t['id'] == task['id']), None)
        if not existing or existing['taskRevision'] != expected_revision:
            raise CloudConflict('המשימה השתנתה מאז פתיחת הטופס. הטיוטה נשמרה במסך; פתח מחדש את המשימה לאחר העתקת השינויים.')
        row = self._rows[task['id']]
        sheet = self.workbook.Worksheets(self.schema['task'])
        fields = dict(task, dueMonth=(task.get('dueDate') or '')[:7])
        before = dict(existing, dueMonth=(existing.get('dueDate') or '')[:7])
        patches = []
        for col, key in enumerate(FIELDS, 1):
            if key == 'id' or fields.get(key) == before.get(key):
                continue
            val = fields.get(key, '')
            if key == 'progress':
                val = float(val or 0) / 100
            if key in ('dueDate', 'startDate', 'completedDate') and val:
                epoch = datetime(1904, 1, 1) if self.workbook.Date1904 else datetime(1899, 12, 30)
                val = (datetime.strptime(val, '%Y-%m-%d') - epoch).days
            if isinstance(val, str) and (len(val) > 32767 or any(ord(c) < 32 and c not in '\t\n\r' for c in val)):
                raise CloudConflict('הטקסט ארוך מדי או מכיל תווים שאינם נתמכים ב-Excel.')
            cell = sheet.Cells(row, col)
            if col in (13, 14, 15) and cell.HasFormula:
                continue
            if cell.HasFormula or (sheet.ProtectContents and cell.Locked):
                raise CloudConflict('אחד השדות מכיל נוסחה או מוגן. לא בוצעו שינויים; בדוק את תבנית הקובץ.')
            patches.append((col, val))
        # A per-task history cell avoids a shared append position across different tasks.
        if history:
            cell = sheet.Cells(row, 20)
            if cell.HasFormula or (sheet.ProtectContents and cell.Locked):
                raise CloudConflict('עמודת ההיסטוריה T מוגנת או מכילה נוסחה.')
            try:
                records = json.loads(cell.Value2 or '[]')
            except (TypeError, ValueError):
                raise CloudConflict('עמודה T משמשת לנתונים אחרים. נדרשת עמודה T פנויה להיסטוריית TeamFlow.')
            if not isinstance(records, list):
                raise CloudConflict('עמודה T אינה זמינה להיסטוריית TeamFlow.')
            event = dict(history, time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            records = (records + [event])[-15:]
            encoded = json.dumps(records, ensure_ascii=False)
            while len(encoded) > 30000 and len(records) > 1:
                records.pop(0)
                encoded = json.dumps(records, ensure_ascii=False)
            if len(encoded) > 32767:
                raise CloudConflict('רשומת ההיסטוריה גדולה מדי.')
            patches.append((20, encoded))
        # Retain intent before handing changes to Office. No automatic retry after a partial COM failure.
        self.recovery_dir.mkdir(parents=True, exist_ok=True)
        journal = self.recovery_dir / (uuid.uuid4().hex + '.json')
        recovery = dict(url=self.url, task=task, previous=existing,
                        status='pending-in-excel', time=datetime.now().isoformat())
        journal.write_text(json.dumps(recovery, ensure_ascii=False, indent=2), encoding='utf-8')
        for col, val in patches:
            cell = sheet.Cells(row, col)
            if isinstance(val, str):
                cell.NumberFormat = '@'
            elif col == 7:
                cell.NumberFormat = '0%'
            elif col in (6, 10, 11):
                cell.NumberFormat = 'yyyy-mm-dd'
            cell.Value2 = val
        self.workbook.Save()
        recovery['status'] = 'saved-in-excel-cloud-not-confirmed'
        journal.write_text(json.dumps(recovery, ensure_ascii=False, indent=2), encoding='utf-8')
        fresh = self._read()
        return {'ok': True, 'data': fresh, 'cloudConfirmed': False}
