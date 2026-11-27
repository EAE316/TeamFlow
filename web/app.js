/**
 * TeamFlow / ניהול משימות צוות
 * Vanilla JavaScript Frontend Application
 */

// Top-level state and variables
let state = null;
let currentView = 'myday';
let calendarMonth = null;
let lastPayload = '';
let editing = false;
let settingsDirty = false;
let metricFilter = '';
let draggedTaskId = '';
let lastMovedTask = null;
let tableSortKey = '';
let tableSortAsc = true;
const selectedTasks = new Set();
let timelineGroupBy = 'owner';
let currentDrawerSubtasks = [];
let currentDrawerCleanNotes = '';
let currentDrawerLinks = [];
let mutationEpoch = 0;
let mutationsPending = 0;
let refreshPending = false;
let savingTask = false;
let drawerRevision = '';
let drawerWorkbook = '';
let drawerTaskRevision = '';
let connectionPollPending = false;
let cloudConnectPending = false;
let lastConnectionUrl = '';
let lastAutomaticRefresh = 0;
let settingsRevision = '';
let settingsWorkbook = '';
let settingsListsRevision = '';

// Utility selectors
const q = (sel) => document.querySelector(sel);
const qa = (sel) => [...document.querySelectorAll(sel)];

// Constants
const ALL = 'הכול';
const DONE = 'בוצע';
const CANCELLED = 'בוטל';
const YES = 'כן';
const NO = 'לא';
const HEALTH_BLOCKED = 'חסום';
const HEALTH_OVERDUE = 'באיחור';
const HEALTH_SOON = 'קרוב ליעד';
const HEALTH_ON_TRACK = 'במסלול';
const HEALTH_CLOSED = 'סגור';

const priorityOrder = Object.assign(Object.create(null), { P0: 0, P1: 1, P2: 2, P3: 3 });
const priorityClass = Object.assign(Object.create(null), { P0: 'p0', P1: 'p1', P2: 'p2', P3: 'p3' });
const priorityColor = Object.assign(Object.create(null), { P0: 'var(--prio-0)', P1: 'var(--prio-1)', P2: 'var(--prio-2)', P3: 'var(--prio-3)' });

const viewMeta = {
  myday: ['היום שלי', 'המשימות שלך לפי מועד.'],
  calendar: ['לוח שנה', 'משימות לפי תאריך יעד.'],
  dashboard: ['דשבורד', 'מבט מהיר על עומס, עדיפויות ומשימות שדורשות טיפול.'],
  timeline: ['ציר זמן (גאנט)', 'מפת דרכים כרונולוגית של משימות הצוות לפי תאריכי ביצוע.'],
  board: ['לוח עבודה', 'כרטיסים לפי אחראי, מסודרים לפי עדיפות ותאריך יעד.'],
  tasks: ['כל המשימות', 'טבלה פשוטה לסינון ועריכה מהירה.'],
  settings: ['צוות ורשימות', 'ניהול אנשים, קטגוריות, סטטוסים ועדיפויות.']
};

/**
 * API call helper
 */
async function api(path, options = {}) {
  const isMutation = (options.method || 'GET').toUpperCase() === 'POST';
  if (isMutation && /^\/api\/(task|delete|done|bulk|reorder|lists)$/.test(path) && state) {
    const target = (state.tasks || []).find(t => t.id === options.body?.id);
    options = { ...options, body: { workbook: state.workbook, revision: state.revision, taskRevision: target?.taskRevision || '', ...options.body } };
  }
  const fetchOptions = {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Accept': 'application/json'
    },
    ...options
  };

  if (fetchOptions.body && typeof fetchOptions.body === 'object') {
    fetchOptions.body = JSON.stringify(fetchOptions.body);
  }

  if (isMutation) { mutationEpoch++; mutationsPending++; }
  try {
  const res = await fetch(path, fetchOptions);
  let data;
  try {
    data = await res.json();
  } catch (err) {
    if (!res.ok) throw new Error('שגיאת שרת');
    return {};
  }

  if (!res.ok || (data && data.ok === false)) {
    const errorMsg = (data && data.error) ? data.error : 'שגיאה בביצוע הפעולה';
    throw new Error(errorMsg);
  }
  return data;
  } finally {
    if (isMutation) { mutationEpoch++; mutationsPending--; }
  }
}

/**
 * Toast Notification
 */
let toastTimer = null;
function toast(message) {
  const el = q('#globalToast');
  if (!el) return;
  el.textContent = message;
  el.classList.add('show');

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.classList.remove('show');
  }, 1900);
}

/**
 * String Escape Helpers
 */
function escapeHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function initials(name) {
  if (!name) return '?';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.slice(0, 2);
}

function clampPercent(value) {
  return Math.max(0, Math.min(100, Math.round(Number(value) || 0)));
}

// Dates are stored as yyyy-mm-dd (sorting and the overdue checks rely on it) and shown as dd/mm/yyyy.
function fmtDate(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : String(value || '');
}
function fmtDateTime(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  return m ? `${m[3]}/${m[2]}/${m[1]}${m[4] ? ` ${m[4]}:${m[5]}` : ''}` : String(value || '');
}
// Typed dates: 30/11/2026, 30.11.26, 30-11-2026 or 30112026. Returns yyyy-mm-dd, '' for empty, null if invalid.
function parseDisplayDate(text) {
  const s = String(text || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/) || s.match(/^(\d{2})(\d{2})(\d{4})$/);
  if (!m) return null;
  const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  const iso = `${year}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
  return parseIsoDate(iso) ? iso : null;
}

// Each person keeps one colour everywhere: their place in the team list picks a slot of the
// validated palette (names outside the list get a stable slot from their spelling).
function personSlot(name) {
  const n = String(name || '').trim();
  if (!n) return '';
  const idx = (state?.people || []).indexOf(n);
  if (idx >= 0) return String(idx % 8);
  let h = 0;
  for (const ch of n) h = (h * 31 + ch.codePointAt(0)) >>> 0;
  return String(h % 8);
}

function avatarHtml(name, cls = 'user-avatar', withTitle = true) {
  const slot = personSlot(name);
  const title = withTitle ? ` title="${escapeAttr(name || 'ללא שיוך')}"` : '';
  return `<span class="${cls}"${slot ? ` data-slot="${slot}"` : ''}${title}>${escapeHtml(initials(name || '?'))}</span>`;
}

const ICONS = {
  calendar: '<rect x="3" y="4.5" width="14" height="12.5" rx="2"/><path d="M3 8.5h14M7 2.5v3M13 2.5v3"/>',
  check: '<path d="m5 10.5 3.2 3L15 6.5"/>',
  checklist: '<rect x="3" y="3" width="14" height="14" rx="3.5"/><path d="m7 10 2 2 4-4"/>',
  comment: '<path d="M4.5 4h11A1.5 1.5 0 0 1 17 5.5v7a1.5 1.5 0 0 1-1.5 1.5H9.5L6 16.8V14H4.5A1.5 1.5 0 0 1 3 12.5v-7A1.5 1.5 0 0 1 4.5 4z"/>',
  link: '<path d="M8.6 11.4a3 3 0 0 0 4.2 0l2.4-2.4A3 3 0 0 0 11 4.8l-.8.8"/><path d="M11.4 8.6a3 3 0 0 0-4.2 0L4.8 11A3 3 0 0 0 9 15.2l.8-.8"/>',
  pencil: '<path d="M13.2 3.8a1.8 1.8 0 0 1 2.6 2.6L7 15.2l-3.5.9.9-3.5z"/>',
  layers: '<path d="M10 3 2.5 7 10 11l7.5-4z"/><path d="m2.5 10.5 7.5 4 7.5-4"/>',
  open: '<circle cx="10" cy="10" r="7"/><circle cx="10" cy="10" r="2.2"/>',
  alarm: '<circle cx="10" cy="11" r="6"/><path d="M10 8v3l2 1.5M3.5 4.5l2-2M16.5 4.5l-2-2"/>',
  blocked: '<circle cx="10" cy="10" r="7"/><path d="M5 15 15 5"/>',
  done: '<circle cx="10" cy="10" r="7"/><path d="m7 10 2 2 4-4"/>',
};
function icon(name) {
  return `<svg class="ico" viewBox="0 0 20 20" aria-hidden="true">${ICONS[name] || ''}</svg>`;
}

// Due date as people say it: "מחר", "באיחור 3 ימים". Tone drives the chip colour; the text always names it.
function dueInfo(t) {
  if (!t.dueDate) return { text: 'ללא יעד', tone: 'none' };
  const d = parseIsoDate(t.dueDate);
  if (!d) return { text: t.dueDate, tone: 'none' };
  const days = Math.round((d - parseIsoDate(todayIso())) / 86400000);
  const closed = t.status === DONE || t.status === CANCELLED;
  if (days < 0 && !closed) {
    return { text: days === -1 ? 'באיחור יום' : days === -2 ? 'באיחור יומיים' : `באיחור ${-days} ימים`, tone: 'overdue' };
  }
  if (days === 0) return { text: 'היום', tone: closed ? '' : 'today' };
  if (days === 1) return { text: 'מחר', tone: closed ? '' : 'soon' };
  return { text: dueLabel(t.dueDate), tone: !closed && days > 0 && days <= 3 ? 'soon' : '' };
}

function progressRing(pct, size = 18) {
  const c = size / 2;
  const r = (size - 3) / 2;
  const len = 2 * Math.PI * r;
  const fill = pct > 0 ? `<circle class="pr-fill" cx="${c}" cy="${c}" r="${r}" stroke-dasharray="${len.toFixed(2)}" stroke-dashoffset="${(len * (1 - pct / 100)).toFixed(2)}"/>` : '';
  return `<svg class="progress-ring" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" aria-hidden="true"><circle class="pr-track" cx="${c}" cy="${c}" r="${r}"/>${fill}</svg>`;
}

// One hover tooltip shared by the dashboard charts.
let chartTipEl = null;
function showChartTip(anchor, html) {
  if (!chartTipEl) {
    chartTipEl = document.createElement('div');
    chartTipEl.className = 'chart-tip';
    chartTipEl.setAttribute('role', 'tooltip');
    document.body.appendChild(chartTipEl);
  }
  chartTipEl.innerHTML = html;
  chartTipEl.hidden = false;
  const r = anchor.getBoundingClientRect();
  const tip = chartTipEl.getBoundingClientRect();
  chartTipEl.style.left = `${Math.round(Math.max(8, Math.min(window.innerWidth - tip.width - 8, r.left + r.width / 2 - tip.width / 2)))}px`;
  chartTipEl.style.top = `${Math.round(Math.max(8, r.top - tip.height - 8))}px`;
}
function hideChartTip() {
  if (chartTipEl) chartTipEl.hidden = true;
}

// Current active user for audit tracking (auto-detected from Excel/Office or selected)
let currentUser = localStorage.getItem('teamflow_active_user') || '';

function getCurrentUser() {
  if (currentUser) return currentUser;
  if (state && state.systemUser) {
    currentUser = state.systemUser;
    localStorage.setItem('teamflow_active_user', currentUser);
    return currentUser;
  }
  if (state && state.people && state.people.length > 0) {
    currentUser = state.people[0];
    localStorage.setItem('teamflow_active_user', currentUser);
    return currentUser;
  }
  return 'משתמש Excel';
}

function setCurrentUser(name) {
  if (!name) return;
  currentUser = name.trim();
  localStorage.setItem('teamflow_active_user', currentUser);
  updateUserProfileDisplay();
  render();
  toast(`הזהות הפעילה עודכנה ל-${currentUser}`);
}

function updateUserProfileDisplay() {
  const user = getCurrentUser();
  const avatarEl = q('#currentUserAvatar');
  const nameEl = q('#currentUserName');
  if (avatarEl) {
    avatarEl.textContent = initials(user);
    const slot = personSlot(user);
    if (slot) avatarEl.dataset.slot = slot; else delete avatarEl.dataset.slot;
  }
  if (nameEl) {
    nameEl.textContent = user;
    nameEl.title = `משתמש פעיל: ${user}`;
  }
  qa('.user-dropdown-item').forEach(it => {
    it.classList.toggle('active', it.dataset.user === user);
  });
}

function renderUserDropdown() {
  const listEl = q('#userDropdownList');
  if (!listEl || !state) return;

  const current = getCurrentUser();
  const options = new Set();
  if (state.systemUser) options.add(state.systemUser);
  (state.people || []).forEach(p => options.add(p));
  if (current) options.add(current);

  listEl.innerHTML = [...options].map(userName => {
    const isAct = userName === current ? 'active' : '';
    const isOffice = userName === state.systemUser ? '<span class="user-dropdown-note"> · חשבון Office</span>' : '';
    return `
      <div class="user-dropdown-item ${isAct}" data-user="${escapeAttr(userName)}">
        ${avatarHtml(userName, 'user-avatar-sm', false)}
        <span class="user-dropdown-name">${escapeHtml(userName)}${isOffice}</span>
        ${userName === current ? '<span class="user-dropdown-check">✓</span>' : ''}
      </div>
    `;
  }).join('');

  listEl.querySelectorAll('.user-dropdown-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.stopPropagation();
      setCurrentUser(item.dataset.user);
      const menu = q('#userDropdownMenu');
      const badge = q('#userProfileSelector');
      if (menu) menu.style.display = 'none';
      if (badge) badge.classList.remove('open');
    });
  });
}

/**
 * Load initial data
 */
async function load() {
  const epoch = mutationEpoch;
  try {
    const data = await api('/api/data');
    if (epoch !== mutationEpoch) return;
    showDataError('');
    state = data;
    updateConnection(data.connection);
    lastPayload = dataFingerprint(data);
    checkReminders();
    if (!currentUser && data.systemUser) {
      currentUser = data.systemUser;
      localStorage.setItem('teamflow_active_user', currentUser);
    }
    updateUserProfileDisplay();
    setView(currentView);
  } catch (err) {
    showDataError(err.message);
    toast(err.message);
  }
}

/**
 * Refresh from disk without overriding user editing
 */
async function refreshFromDisk(silent = true) {
  if (mutationsPending || refreshPending || cloudConnectPending) return;
  const epoch = mutationEpoch;
  refreshPending = true;
  try {
    const data = await api('/api/data');
    if (mutationsPending || epoch !== mutationEpoch) return;
    showDataError('');
    updateConnection(data.connection);
    const newPayload = dataFingerprint(data);
    if (newPayload !== lastPayload) {
      state = data;
      lastPayload = newPayload;
      render();
      checkReminders();
      const openId = editing ? q('#formTaskId').value : '';
      if (openId) renderActivity(state.tasks.find(t => t.id === openId));
      if (!silent) toast('הנתונים עודכנו מהאקסל');
    } else {
      if (!silent) toast('אין שינויים חדשים');
    }
  } catch (err) {
    showDataError(err.message);
    if (!silent) toast(err.message);
  } finally {
    refreshPending = false;
  }
}

function showDataError(message) {
  const banner = q('#dataError');
  if (!banner) return;
  banner.hidden = !message;
  q('#dataErrorText').textContent = message ? `לא ניתן לרענן את סביבת העבודה. ${message}` : '';
}

/**
 * Helper to populate select dropdowns
 */
function fillSelect(select, values, includeAll = false) {
  if (!select) return;
  const current = select.value;
  select.innerHTML = '';
  if (includeAll) {
    const optAll = document.createElement('option');
    optAll.value = ALL;
    optAll.textContent = ALL;
    select.appendChild(optAll);
  }
  for (const v of values) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v || 'ללא שיוך';
    select.appendChild(opt);
  }
  if (current && (values.includes(current) || current === ALL)) {
    select.value = current;
  }
}

/**
 * Returns tasks matching global search across all fields
 */
function visibleTasks() {
  if (!state || !state.tasks) return [];
  const searchInput = q('#globalSearchInput');
  const term = (searchInput ? searchInput.value : '').trim().toLowerCase();

  const matchTag = q('#searchMatchCount');
  const clearBtn = q('#btnClearSearch');

  if (!term) {
    if (matchTag) matchTag.style.display = 'none';
    if (clearBtn) clearBtn.style.display = 'none';
    return state.tasks;
  }

  const filtered = state.tasks.filter(t => {
    return (
      (t.title && String(t.title).toLowerCase().includes(term)) ||
      (t.notes && String(t.notes).toLowerCase().includes(term)) ||
      (t.owner && String(t.owner).toLowerCase().includes(term)) ||
      (t.category && String(t.category).toLowerCase().includes(term)) ||
      (t.status && String(t.status).toLowerCase().includes(term)) ||
      (t.priority && String(t.priority).toLowerCase().includes(term)) ||
      (t.health && String(t.health).toLowerCase().includes(term)) ||
      (t.id && String(t.id).toLowerCase().includes(term))
    );
  });

  if (matchTag) {
    matchTag.style.display = 'inline-block';
    matchTag.textContent = `${filtered.length} נמצאו`;
  }
  if (clearBtn) {
    clearBtn.style.display = 'inline-block';
  }

  return filtered;
}

/**
 * Returns tasks where status is not done and not cancelled
 */
function openTasks(tasks = visibleTasks()) {
  if (!tasks) return [];
  return tasks.filter(t => t.status !== DONE && t.status !== CANCELLED);
}

/**
 * Returns tasks filtered for the tasks table view
 */
function filteredTableTasks() {
  let list = visibleTasks();

  const filterOwner = q('#filterOwner');
  const filterStatus = q('#filterStatus');
  const filterPriority = q('#filterPriority');

  if (filterOwner && filterOwner.value && filterOwner.value !== ALL) {
    list = list.filter(t => t.owner === filterOwner.value);
  }
  if (filterStatus && filterStatus.value && filterStatus.value !== ALL) {
    list = list.filter(t => t.status === filterStatus.value);
  }
  if (filterPriority && filterPriority.value && filterPriority.value !== ALL) {
    list = list.filter(t => t.priority === filterPriority.value);
  }

  // Metric filter
  if (metricFilter === 'my-tasks') {
    const user = getCurrentUser();
    list = list.filter(t => t.owner === user && t.status !== DONE && t.status !== CANCELLED);
  } else if (metricFilter === 'open') {
    list = list.filter(t => t.status !== DONE && t.status !== CANCELLED);
  } else if (metricFilter === 'overdue') {
    list = list.filter(t => t.health === HEALTH_OVERDUE);
  } else if (metricFilter === 'blocked') {
    list = list.filter(t => t.health === HEALTH_BLOCKED);
  } else if (metricFilter === 'done') {
    list = list.filter(t => t.status === DONE);
  }

  // Column header sort
  if (tableSortKey) {
    list = [...list].sort((a, b) => {
      let valA = a[tableSortKey] ?? '';
      let valB = b[tableSortKey] ?? '';
      let cmp = 0;

      if (tableSortKey === 'priority') {
        const pA = priorityOrder[a.priority] !== undefined ? priorityOrder[a.priority] : 99;
        const pB = priorityOrder[b.priority] !== undefined ? priorityOrder[b.priority] : 99;
        cmp = pA - pB;
      } else if (tableSortKey === 'due') {
        if (a.dueDate && b.dueDate) cmp = a.dueDate.localeCompare(b.dueDate);
        else cmp = a.dueDate ? -1 : (b.dueDate ? 1 : 0);
      } else if (tableSortKey === 'status') {
        const statuses = state && state.statuses ? state.statuses : [];
        const sA = statuses.indexOf(a.status);
        const sB = statuses.indexOf(b.status);
        cmp = (sA !== -1 && sB !== -1) ? sA - sB : String(valA).localeCompare(String(valB), 'he');
      } else if (tableSortKey === 'health') {
        const healthOrder = { [HEALTH_OVERDUE]: 0, [HEALTH_BLOCKED]: 1, [HEALTH_SOON]: 2, [HEALTH_ON_TRACK]: 3, [HEALTH_CLOSED]: 4 };
        const hA = healthOrder[a.health] !== undefined ? healthOrder[a.health] : 9;
        const hB = healthOrder[b.health] !== undefined ? healthOrder[b.health] : 9;
        cmp = hA - hB;
      } else {
        cmp = String(valA).localeCompare(String(valB), 'he');
      }

      return tableSortAsc ? cmp : -cmp;
    });
  }

  return list;
}

/**
 * Sort function for board cards
 */
function taskSort(a, b) {
  const pA = priorityOrder[a.priority] !== undefined ? priorityOrder[a.priority] : 99;
  const pB = priorityOrder[b.priority] !== undefined ? priorityOrder[b.priority] : 99;
  if (pA !== pB) return pA - pB;

  // Due date
  if (a.dueDate && b.dueDate) {
    if (a.dueDate !== b.dueDate) return a.dueDate.localeCompare(b.dueDate);
  } else if (a.dueDate && !b.dueDate) {
    return -1;
  } else if (!a.dueDate && b.dueDate) {
    return 1;
  }

  return (a.title || '').localeCompare(b.title || '');
}

/**
 * Main Render Dispatcher
 */
function render() {
  if (!state) return;
  const existingIds = new Set((state.tasks || []).map(t => t.id));
  for (const id of selectedTasks) if (!existingIds.has(id)) selectedTasks.delete(id);

  const cloudMode = state.connection?.mode === 'excel';

  // Update workbook path display
  const pathEl = q('#workbookPathDisplay');
  if (pathEl) {
    let name = (state.workbook || '').split(/[\\\/]/).pop();
    if (cloudMode) { try { name = decodeURIComponent(name); } catch { /* keep the raw name */ } }
    pathEl.textContent = name || 'לא נבחר קובץ';
    pathEl.title = state.workbook || '';
  }

  // Update dropdown selects in drawer and filters
  if (!editing) {
  fillSelect(q('#formTaskOwner'), ['', ...(state.people || [])]);
  fillSelect(q('#formTaskPriority'), state.priorities || []);
  fillSelect(q('#formTaskStatus'), state.statuses || []);
  fillSelect(q('#formTaskCategory'), ['', ...(state.categories || [])]);
  }

  fillSelect(q('#filterOwner'), [...new Set([...(state.people || []), ...state.tasks.map(t => t.owner).filter(Boolean)])], true);
  fillSelect(q('#filterStatus'), [...new Set([...(state.statuses || []), ...state.tasks.map(t => t.status).filter(Boolean)])], true);
  fillSelect(q('#filterPriority'), [...new Set([...(state.priorities || []), ...state.tasks.map(t => t.priority).filter(Boolean)])], true);

  // Update Bulk Dropdowns
  renderBulkDropdownMenus();

  // Populate settings textareas if not dirty and not focused
  const settingsInputs = [q('#settingsPeople'), q('#settingsCategories'), q('#settingsStatuses'), q('#settingsPriorities')].filter(Boolean);
  const isAnySettingsFocused = settingsInputs.some(el => document.activeElement === el);

  if (!editing && !settingsDirty && !isAnySettingsFocused) {
    const setPeople = q('#settingsPeople');
    const setCat = q('#settingsCategories');
    const setStat = q('#settingsStatuses');
    const setPrio = q('#settingsPriorities');

    if (setPeople) setPeople.value = (state.people || []).join('\n');
    if (setCat) setCat.value = (state.categories || []).join('\n');
    if (setStat) setStat.value = (state.statuses || []).join('\n');
    if (setPrio) setPrio.value = (state.priorities || []).join('\n');
  }

  // Update User Identity Display & Dropdown
  updateUserProfileDisplay();
  renderUserDropdown();

  // Render view-specific content
  if (currentView === 'dashboard') {
    renderExecutiveSummary();
    renderMetrics();
    renderPriorityChart();
    renderWorkloadChart();
    renderAlerts();
    renderTrend();
  } else if (currentView === 'timeline') {
    renderTimeline();
  } else if (currentView === 'board') {
    renderBoard();
  } else if (currentView === 'tasks') {
    renderTable();
    renderHistory();
  } else if (currentView === 'myday') {
    renderMyDay();
  } else if (currentView === 'calendar') {
    renderCalendar();
  }
  const myDayCount = q('#navCountMyDay');
  if (myDayCount) {
    const buckets = myDayBuckets();
    const urgent = buckets.overdue.length + buckets.today.length;
    myDayCount.textContent = urgent ? String(urgent) : '';
  }

  // Update dynamic sidebar counters
  updateSidebarCounters();

  renderBulkBar();
}

/**
 * Update dynamic badges in the sidebar
 */
function updateSidebarCounters() {
  if (!state || !state.tasks) return;
  const tasks = state.tasks;
  const myUser = getCurrentUser();
  const myTasksCount = tasks.filter(t => t.owner === myUser && t.status !== DONE && t.status !== CANCELLED).length;
  const overdueCount = tasks.filter(t => t.health === HEALTH_OVERDUE).length;
  const blockedCount = tasks.filter(t => t.health === HEALTH_BLOCKED).length;
  const p0Count = tasks.filter(t => t.priority === 'P0').length;
  const doneCount = tasks.filter(t => t.status === DONE).length;

  const elMyTasks = q('#sidebarCountMyTasks');
  const elOverdue = q('#sidebarCountOverdue');
  const elBlocked = q('#sidebarCountBlocked');
  const elP0 = q('#sidebarCountP0');
  const elDone = q('#sidebarCountDone');

  if (elMyTasks) elMyTasks.textContent = myTasksCount;
  if (elOverdue) elOverdue.textContent = overdueCount;
  if (elBlocked) elBlocked.textContent = blockedCount;
  if (elP0) elP0.textContent = p0Count;
  if (elDone) elDone.textContent = doneCount;
}

/**
 * Render Executive Summary in Hero Banner & Smart Insights
 */
function renderExecutiveSummary() {
  if (!state || !state.tasks) return;
  const tasks = state.tasks;
  const total = tasks.length;
  const doneCount = tasks.filter(t => t.status === DONE).length;
  const rate = total > 0 ? Math.round((doneCount / total) * 100) : 0;

  const openT = openTasks(tasks);
  const activePeople = new Set(openT.map(t => t.owner).filter(Boolean)).size;

  const rateEl = q('#heroStatRate');
  const peopleEl = q('#heroStatActivePeople');
  if (rateEl) rateEl.textContent = `${rate}%`;
  if (peopleEl) peopleEl.textContent = activePeople;
  const avg = openT.length ? Math.round(openT.reduce((sum, t) => sum + clampPercent(t.progress), 0) / openT.length) : 0;
  q('#heroStatOpen').textContent = openT.length;
  q('#heroStatAvg').textContent = `${avg}%`;
  q('#heroDate').textContent = new Date().toLocaleDateString('he-IL', { weekday: 'long', day: 'numeric', month: 'long' });
  q('#heroRingFill').style.strokeDashoffset = String(314.16 * (1 - rate / 100));
  q('#heroRingSub').textContent = `הושלמו (${doneCount}/${total})`;
  q('#heroRing').setAttribute('aria-label', `${rate}% מהמשימות הושלמו, ${doneCount} מתוך ${total}`);

  const greeting = q('#heroGreeting');
  if (greeting) {
    const hour = new Date().getHours();
    const part = hour < 12 ? 'בוקר טוב' : hour < 17 ? 'צהריים טובים' : 'ערב טוב';
    const first = String(getCurrentUser()).trim().split(/\s+/)[0];
    greeting.textContent = first ? `${part}, ${first}` : part;
  }

  // One plain-language line about what needs attention.
  const insightEl = q('#smartInsightText');
  if (!insightEl) return;

  const overdue = tasks.filter(t => t.health === HEALTH_OVERDUE);
  const blocked = tasks.filter(t => t.health === HEALTH_BLOCKED);
  const p0Open = overdue.filter(t => t.priority === 'P0');

  if (overdue.length > 0 && p0Open.length > 0) {
    insightEl.innerHTML = `<strong>${overdue.length} משימות באיחור</strong>, מתוכן ${p0Open.length} בעדיפות P0. כדאי להתחיל בהן.`;
  } else if (overdue.length > 0) {
    insightEl.innerHTML = `<strong>${overdue.length} משימות באיחור.</strong> לחיצה על "באיחור" למטה מציגה אותן.`;
  } else if (blocked.length > 0) {
    insightEl.innerHTML = `<strong>${blocked.length} משימות חסומות</strong> ממתינות להחלטה או לעזרה.`;
  } else if (total > 0 && rate >= 80) {
    insightEl.innerHTML = `<strong>${rate}% מהמשימות הושלמו.</strong> הצוות קרוב לסיום.`;
  } else if (openT.length > 0) {
    insightEl.textContent = `${openT.length} משימות פתוחות אצל ${activePeople} אנשים, בלי איחורים או חסימות.`;
  } else {
    insightEl.textContent = 'אין משימות פתוחות כרגע.';
  }
}

/**
 * Render 5 clickable Metric Cards
 */
function renderMetrics() {
  const container = q('#metricsGrid');
  if (!container || !state) return;

  const tasks = visibleTasks();
  const total = tasks.length;
  const openCount = openTasks(tasks).length;
  const overdueCount = tasks.filter(t => t.health === HEALTH_OVERDUE).length;
  const blockedCount = tasks.filter(t => t.health === HEALTH_BLOCKED).length;
  const doneCount = tasks.filter(t => t.status === DONE).length;

  const metrics = [
    { key: 'all', label: 'כל המשימות', count: total, hint: 'בקובץ המשותף', tone: '', icon: 'layers' },
    { key: 'open', label: 'פתוחות', count: openCount, hint: 'עדיין בטיפול', tone: 'accent', icon: 'open' },
    { key: 'overdue', label: 'באיחור', count: overdueCount, hint: 'עברו את תאריך היעד', tone: 'critical', icon: 'alarm' },
    { key: 'blocked', label: 'חסומות', count: blockedCount, hint: 'דורשות עזרה או החלטה', tone: 'serious', icon: 'blocked' },
    { key: 'done', label: 'בוצעו', count: doneCount, hint: 'משימות שהסתיימו', tone: 'good', icon: 'done' },
  ];

  container.innerHTML = metrics.map(m => `
    <button type="button" class="metric-card" data-tone="${m.tone}" data-metric="${m.key}">
      <span class="metric-icon">${icon(m.icon)}</span>
      <span class="metric-value">${m.count}</span>
      <span class="metric-label">${escapeHtml(m.label)}</span>
      <span class="metric-hint">${escapeHtml(m.hint)}</span>
    </button>
  `).join('');

  qa('.metric-card').forEach(card => {
    card.addEventListener('click', () => {
      applyMetricFilter(card.dataset.metric);
    });
  });
}

/**
 * Render Priority Breakdown Chart
 */
function renderPriorityChart() {
  const container = q('#priorityChartContainer');
  if (!container || !state) return;

  const tasks = visibleTasks();
  const total = tasks.length || 1;
  const priorities = [...new Set([...(state.priorities || []), ...tasks.map(t => t.priority)])];

  const rowsHtml = priorities.map(pr => {
    const count = tasks.filter(t => t.priority === pr).length;
    const pct = Math.round((count / total) * 100);
    const color = priorityColor[pr] || 'var(--bar-default)';

    return `
      <div class="priority-row">
        <div class="priority-header">
          <div class="priority-badge-label">
            <span class="priority-badge ${priorityClass[pr] || ''}">${escapeHtml(pr)}</span>
            <span>${count} משימות</span>
          </div>
          <span class="priority-pct">${pct}%</span>
        </div>
        <div class="priority-bar-track">
          <div class="priority-bar-fill" style="width: ${pct}%; background: ${color}"></div>
        </div>
      </div>
    `;
  }).join('');

  container.innerHTML = `<div class="priority-list">${rowsHtml}</div>`;
}

/**
 * Render Workload Chart with Priority Stacked Bar and Category Breakdown
 */
function renderWorkloadChart() {
  const container = q('#workloadChartContainer');
  if (!container || !state) return;

  const openT = openTasks(visibleTasks());
  const people = [...new Set([...(state.people || []), ...openT.map(t => t.owner)])];
  const lists = people.map(p => openT.filter(t => t.owner === p));
  const max = Math.max(1, ...lists.map(list => list.length));

  const rows = people.map((p, i) => {
    const list = lists[i];
    const name = p || 'ללא שיוך';
    const pCounts = Object.assign(Object.create(null), { P0: 0, P1: 0, P2: 0, P3: 0 });
    const cats = Object.create(null);
    for (const t of list) {
      pCounts[pCounts[t.priority] !== undefined ? t.priority : 'P2']++;
      const c = t.category || 'כללי';
      cats[c] = (cats[c] || 0) + 1;
    }
    const catText = Object.entries(cats).sort((a, b) => b[1] - a[1]).map(([c, n]) => `${c} ${n}`).join(', ');
    const tip = list.length ? `${name}: ${list.length} פתוחות${catText ? ` · ${catText}` : ''}` : `${name}: אין משימות פתוחות`;
    const segs = ['P0', 'P1', 'P2', 'P3'].filter(k => pCounts[k]).map(k =>
      `<div class="bar-seg seg-${k.toLowerCase()}" style="flex: ${pCounts[k]}" title="${k}: ${pCounts[k]}"></div>`).join('');
    return `
      <div class="workload-row${list.length ? '' : ' is-empty'}" data-owner="${escapeAttr(p)}" title="${escapeAttr(tip)}"${p ? ' style="cursor: pointer"' : ''}>
        <div class="workload-user">${avatarHtml(p, 'user-avatar-sm', false)}<span class="workload-name">${escapeHtml(name)}</span></div>
        <div class="workload-track">${list.length ? `<div class="workload-stacked-bar" style="width: ${(list.length / max * 100).toFixed(1)}%">${segs}</div>` : ''}</div>
        <span class="workload-count">${list.length}</span>
      </div>`;
  }).join('');

  container.innerHTML = `<div class="workload-rows">${rows}</div>`;
  container.querySelectorAll('.workload-row[data-owner]').forEach(row => {
    if (!row.dataset.owner) return;
    row.addEventListener('click', () => {
      metricFilter = '';
      setView('tasks');
      const fo = q('#filterOwner');
      if ([...fo.options].some(o => o.value === row.dataset.owner)) fo.value = row.dataset.owner;
      renderTable();
    });
  });
}

/**
 * Render Alerts Panel
 */
function renderAlerts() {
  const container = q('#alertsContainer');
  if (!container || !state) return;

  const rank = { [HEALTH_OVERDUE]: 0, [HEALTH_BLOCKED]: 1, [HEALTH_SOON]: 2 };
  const alerts = visibleTasks()
    .filter(t => t.health === HEALTH_OVERDUE || t.health === HEALTH_BLOCKED || t.health === HEALTH_SOON)
    .sort((a, b) => rank[a.health] - rank[b.health] || (a.dueDate || '9').localeCompare(b.dueDate || '9'));

  if (alerts.length === 0) {
    container.innerHTML = '<div class="empty-state">✓ אין משימות שדורשות תשומת לב כרגע</div>';
    return;
  }

  const tagFor = { [HEALTH_OVERDUE]: 'alert-tag-overdue', [HEALTH_BLOCKED]: 'alert-tag-blocked', [HEALTH_SOON]: 'alert-tag-soon' };
  const itemsHtml = alerts.slice(0, 7).map(t => `
      <div class="alert-item" data-id="${escapeAttr(t.id)}">
        ${avatarHtml(t.owner, 'user-avatar-sm', false)}
        <div class="alert-main">
          <div class="alert-title">${escapeHtml(t.title)}</div>
          <div class="alert-meta">${escapeHtml(t.owner || 'ללא שיוך')} · ${escapeHtml(t.dueDate ? dueInfo(t).text : 'ללא תאריך יעד')}</div>
        </div>
        <span class="alert-tag ${tagFor[t.health]}">${escapeHtml(t.health)}</span>
      </div>`).join('');
  const more = alerts.length > 7 ? `<div class="alerts-more">ועוד ${alerts.length - 7} משימות</div>` : '';

  container.innerHTML = `<div class="alerts-list">${itemsHtml}</div>${more}`;
  container.querySelectorAll('.alert-item').forEach(el => {
    el.addEventListener('click', () => openTask(el.dataset.id));
  });
}

/**
 * Completed per day over the last 14 days: thin columns, the tallest labelled, hover for each value.
 */
function renderTrend() {
  const container = q('#trendChartContainer');
  if (!container || !state) return;
  hideChartTip();
  const days = [];
  for (let i = 13; i >= 0; i--) days.push(isoPlus(-i));
  const counts = Object.create(null);
  for (const t of state.tasks || []) {
    if (t.status === DONE && t.completedDate) {
      const d = String(t.completedDate).slice(0, 10);
      counts[d] = (counts[d] || 0) + 1;
    }
  }
  const values = days.map(d => counts[d] || 0);
  const total = values.reduce((a, b) => a + b, 0);
  const week = values.slice(-7).reduce((a, b) => a + b, 0);
  const max = Math.max(...values);
  const scale = Math.max(max, 4);
  const maxIdx = max ? values.lastIndexOf(max) : -1;
  q('#trendSubtitle').textContent = total
    ? `${week} בשבוע האחרון · ${total} ב-14 הימים האחרונים`
    : 'משימות שהושלמו בכל יום, 14 הימים האחרונים';
  const letters = ['א׳', 'ב׳', 'ג׳', 'ד׳', 'ה׳', 'ו׳', 'ש׳'];
  const dayName = iso => `יום ${HEBREW_WEEKDAYS[parseIsoDate(iso).getDay()]} ${fmtDate(iso)}`;
  container.innerHTML = `
    <div class="trend-chart" role="img" aria-label="${escapeAttr(`${total} משימות הושלמו ב-14 הימים האחרונים, ${week} מהן בשבוע האחרון`)}">
      <div class="trend-plot">${days.map((d, i) => {
        const h = (values[i] / scale * 100).toFixed(1);
        return `
        <div class="trend-col${i === 13 ? ' is-today' : ''}" data-date="${d}" data-count="${values[i]}">
          ${i === maxIdx ? `<span class="trend-val" style="bottom: ${h}%">${values[i]}</span>` : ''}
          <div class="trend-bar" style="height: ${h}%"></div>
        </div>`;
      }).join('')}</div>
      <div class="trend-labels" aria-hidden="true">${days.map((d, i) => `<span class="${i === 13 ? 'is-today' : ''}">${letters[parseIsoDate(d).getDay()]}</span>`).join('')}</div>
      ${total ? '' : '<div class="trend-empty">עוד לא הושלמו משימות בשבועיים האחרונים.</div>'}
    </div>`;
  container.querySelectorAll('.trend-col').forEach(col => {
    col.addEventListener('mouseenter', () => showChartTip(col, `<strong>${col.dataset.count}</strong> הושלמו · ${escapeHtml(dayName(col.dataset.date))}`));
    col.addEventListener('mouseleave', hideChartTip);
  });
}

/**
 * Generate HTML for Board Card
 */
function taskCardHtml(t, commentCounts = null) {
  const pClass = priorityClass[t.priority] || 'p2';
  const selected = selectedTasks.has(t.id);
  const healthClass = (t.health || '').replace(/\s+/g, '-');
  const prog = clampPercent(t.progress);
  const due = dueInfo(t);
  const parsed = parseNotesAndSubtasks(t.notes || '');
  const subDone = parsed.subtasks.filter(s => s.done).length;
  const comments = commentCounts ? (commentCounts[t.id] || 0) : 0;
  const showHealth = t.health === HEALTH_OVERDUE || t.health === HEALTH_BLOCKED || t.health === HEALTH_SOON;
  const stats = [
    parsed.subtasks.length ? `<span class="card-stat" title="משימות משנה שהושלמו">${icon('checklist')}${subDone}/${parsed.subtasks.length}</span>` : '',
    comments ? `<span class="card-stat" title="תגובות">${icon('comment')}${comments}</span>` : '',
    parsed.links.length ? `<span class="card-stat" title="קישורים">${icon('link')}${parsed.links.length}</span>` : '',
  ].join('');

  return `
    <div class="task-card${selected ? ' is-selected' : ''}" draggable="true" data-id="${escapeAttr(t.id)}" data-owner="${escapeAttr(t.owner)}" data-priority="${escapeAttr(t.priority)}">
      <div class="card-top-row">
        <div class="card-top-left">
          <input type="checkbox" class="card-checkbox" data-id="${escapeAttr(t.id)}" aria-label="בחירת ${escapeAttr(t.title)}" ${selected ? 'checked' : ''}>
          <span class="priority-badge ${pClass}">${escapeHtml(t.priority)}</span>
        </div>
        ${showHealth ? `<span class="health-pill ${escapeAttr(healthClass)}">${escapeHtml(t.health)}</span>` : ''}
      </div>
      <div class="card-title" data-id="${escapeAttr(t.id)}">${escapeHtml(t.title)}</div>
      <div class="card-meta">
        <span class="chip due-chip due-${due.tone || 'plain'}" title="${escapeAttr(t.dueDate ? `תאריך יעד: ${fmtDate(t.dueDate)}` : 'ללא תאריך יעד')}">${icon('calendar')}${escapeHtml(due.text)}</span>
        ${t.category ? `<span class="chip">${escapeHtml(t.category)}</span>` : ''}
        ${t.status ? `<span class="chip">${escapeHtml(t.status)}</span>` : ''}
      </div>
      <div class="card-footer">
        <div class="card-stats"><span class="card-stat" title="אחוז ביצוע">${progressRing(prog)}${prog}%</span>${stats}</div>
        <div class="card-end">
          <button type="button" class="btn-card-done" data-id="${escapeAttr(t.id)}" title="סימון כבוצע">${icon('check')}סיום</button>
          ${avatarHtml(t.owner, 'avatar-xs')}
        </div>
      </div>
    </div>
  `;
}

/**
 * Render Board View
 */
function renderBoard() {
  const container = q('#boardColumnsContainer');
  if (!container || !state) return;

  const people = state.people || [];
  const openT = openTasks(visibleTasks());

  const sortSelect = q('#boardSortSelect');
  const sortMode = sortSelect ? sortSelect.value : 'manual';
  const searchInput = q('#globalSearchInput');
  const isSearching = Boolean(searchInput && searchInput.value.trim());

  // Determine all columns: people in settings + extra owners who still have open tasks
  const peopleSet = new Set(people);
  const extraOwners = [];
  openT.forEach(t => {
    const o = (t.owner || '').trim();
    if (o && !peopleSet.has(o) && !extraOwners.includes(o)) {
      extraOwners.push(o);
    }
  });

  const allColumns = [
    ...people.map(p => ({ owner: p, isExtra: false, isUnassigned: false })),
    ...extraOwners.map(p => ({ owner: p, isExtra: true, isUnassigned: false })),
  ];

  const unassignedTasks = openT.filter(t => !(t.owner || '').trim());
  if (unassignedTasks.length > 0) {
    allColumns.push({ owner: '', isExtra: false, isUnassigned: true, label: 'ללא שיוך' });
  }

  // Comment counts for the cards, counted once per render.
  const commentCounts = Object.create(null);
  for (const t of openT) {
    const n = (t.activity || []).filter(e => e.comment).length;
    if (n) commentCounts[t.id] = n;
  }
  const countedFromTask = new Set(Object.keys(commentCounts));
  for (const h of state.history || []) {
    if (h.comment && h.id && !countedFromTask.has(h.id)) commentCounts[h.id] = (commentCounts[h.id] || 0) + 1;
  }

  // Group by owner
  const columnsHtml = allColumns.map(col => {
    const person = col.owner;
    let personTasks = col.isUnassigned
      ? unassignedTasks
      : openT.filter(t => t.owner === person);

    if (sortMode === 'priority') {
      personTasks.sort((a, b) => {
        const pA = priorityOrder[a.priority] !== undefined ? priorityOrder[a.priority] : 99;
        const pB = priorityOrder[b.priority] !== undefined ? priorityOrder[b.priority] : 99;
        return pA - pB;
      });
    } else if (sortMode === 'due') {
      personTasks.sort((a, b) => {
        if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
        return a.dueDate ? -1 : (b.dueDate ? 1 : 0);
      });
    } else if (sortMode === 'title') {
      personTasks.sort((a, b) => (a.title || '').localeCompare(b.title || '', 'he'));
    }
    // If 'manual': keep exact state.tasks order!

    const emptyText = isSearching ? 'אין תוצאות תואמות' : 'אין משימות פתוחות. אפשר לגרור לכאן כרטיס.';
    const cardsHtml = personTasks.length > 0
      ? personTasks.map(t => taskCardHtml(t, commentCounts)).join('')
      : `<div class="col-empty">${emptyText}</div>`;

    const displayName = col.isUnassigned ? 'ללא שיוך' : person;
    const badgeExtra = col.isExtra ? '<span class="col-extra-tag">(מחוץ לרשימה)</span>' : '';

    return `
      <div class="board-column" data-owner="${escapeAttr(person)}"${personSlot(person) ? ` data-slot="${personSlot(person)}"` : ''}>
        <div class="col-header">
          <div class="col-user-info">
            ${avatarHtml(col.isUnassigned ? '' : person, 'user-avatar', false)}
            <span class="col-title">${escapeHtml(displayName)} ${badgeExtra}</span>
          </div>
          <div class="col-header-actions">
            ${col.isUnassigned ? '' : `<button class="btn-col-add" data-owner="${escapeAttr(person)}" title="הוספת משימה ל-${escapeAttr(person)}">+</button>`}
            <span class="col-badge">${personTasks.length}</span>
          </div>
        </div>
        <div class="col-cards" data-owner="${escapeAttr(person)}">
          ${cardsHtml}
        </div>
      </div>
    `;
  }).join('');

  container.innerHTML = columnsHtml;

  // Attach card event listeners (checkbox, quick done, click to edit, drag & drop)
  container.querySelectorAll('.btn-col-add').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openNewTaskFor(btn.dataset.owner);
    });
  });

  container.querySelectorAll('.card-checkbox').forEach(cb => {
    cb.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleSelection(cb.dataset.id, cb.checked);
    });
  });

  container.querySelectorAll('.btn-card-done').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      btn.closest('.task-card')?.classList.add('completing');
      setTimeout(() => markDone(btn.dataset.id), 220);
    });
  });

  // The whole card opens the task; the checkbox and buttons keep their own action.
  container.querySelectorAll('.task-card').forEach(card => {
    card.addEventListener('click', (e) => {
      if (e.target.closest('input, button')) return;
      openTask(card.dataset.id);
    });
  });

  if (lastMovedTask && Date.now() - lastMovedTask.at < 5000) {
    container.querySelector(`.task-card[data-id="${CSS.escape(lastMovedTask.id)}"]`)?.classList.add('just-moved');
  }

  // Reorder and drop handler
  async function handleDrop(sourceId, targetId, targetOwner) {
    if (!sourceId) return;
    const sourceTask = state.tasks.find(t => t.id === sourceId);
    if (!sourceTask) return;

    let ownerChanged = false;
    if (targetOwner !== undefined && sourceTask.owner !== targetOwner) {
      ownerChanged = true;
    }

    let currentIds = state.tasks.map(t => t.id);
    const srcIdx = currentIds.indexOf(sourceId);
    if (srcIdx !== -1) {
      currentIds.splice(srcIdx, 1);
      if (targetId && targetId !== sourceId) {
        const tgtIdx = currentIds.indexOf(targetId);
        if (tgtIdx !== -1) {
          currentIds.splice(tgtIdx, 0, sourceId);
        } else {
          currentIds.push(sourceId);
        }
      } else {
        currentIds.push(sourceId);
      }
    }

    if (sortSelect) {
      sortSelect.value = 'manual';
    }

    if (ownerChanged) {
      try {
        await api('/api/bulk', { method: 'POST', body: { ids: [sourceId], action: 'owner', value: targetOwner, user: getCurrentUser() } });
        await load();
        toast(`המשימה הועברה ל-${targetOwner}`);
      } catch (err) {
        toast(err.message);
        return;
      }
    }

    lastMovedTask = { id: sourceId, at: Date.now() };
    await persistOrder(currentIds);
    lastMovedTask = null;
  }

  // Drag and drop handlers on cards
  container.querySelectorAll('.task-card').forEach(card => {
    card.addEventListener('dragstart', (e) => {
      draggedTaskId = card.dataset.id;
      e.dataTransfer.setData('text/plain', card.dataset.id);
      card.classList.add('dragging');
    });

    card.addEventListener('dragend', () => {
      draggedTaskId = '';
      card.classList.remove('dragging');
      qa('.task-card').forEach(c => c.classList.remove('drag-over'));
      qa('.col-cards').forEach(c => c.classList.remove('drag-col-over'));
    });

    card.addEventListener('dragover', (e) => {
      e.preventDefault();
      card.classList.add('drag-over');
    });

    card.addEventListener('dragleave', () => {
      card.classList.remove('drag-over');
    });

    card.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      card.classList.remove('drag-over');
      const targetId = card.dataset.id;
      const targetOwner = card.dataset.owner;
      handleDrop(draggedTaskId, targetId, targetOwner);
    });
  });

  // Drag and drop handlers on column background (for empty column or dropping below cards)
  container.querySelectorAll('.col-cards').forEach(col => {
    col.addEventListener('dragover', (e) => {
      e.preventDefault();
      col.classList.add('drag-col-over');
    });

    col.addEventListener('dragleave', () => {
      col.classList.remove('drag-col-over');
    });

    col.addEventListener('drop', (e) => {
      e.preventDefault();
      col.classList.remove('drag-col-over');
      const targetOwner = col.dataset.owner;
      handleDrop(draggedTaskId, null, targetOwner);
    });
  });
}

/**
 * Reorder helper for drag and drop
 */
function orderedIdsAfterDrag(sourceId, targetId) {
  const currentIds = state.tasks.map(t => t.id);
  const srcIdx = currentIds.indexOf(sourceId);
  const tgtIdx = currentIds.indexOf(targetId);
  if (srcIdx === -1 || tgtIdx === -1) return currentIds;

  const result = [...currentIds];
  const [removed] = result.splice(srcIdx, 1);
  result.splice(tgtIdx, 0, removed);
  return result;
}

/**
 * Render Tasks Table View
 */
function renderTable() {
  const tbody = q('#tasksTableBody');
  const emptyState = q('#tableEmptyState');
  if (!tbody || !state) return;

  const tasks = filteredTableTasks();

  if (tasks.length === 0) {
    tbody.innerHTML = '';
    if (emptyState) emptyState.style.display = 'block';
    return;
  }

  if (emptyState) emptyState.style.display = 'none';

  tbody.innerHTML = tasks.map(t => {
    const isChecked = selectedTasks.has(t.id) ? 'checked' : '';
    const pClass = priorityClass[t.priority] || 'p2';
    const due = dueInfo(t);
    const healthClass = (t.health || '').replace(/\s+/g, '-');

    let rowClass = selectedTasks.has(t.id) ? 'is-selected ' : '';
    if (t.health === HEALTH_OVERDUE) rowClass += 'row-overdue';
    else if (t.health === HEALTH_BLOCKED) rowClass += 'row-blocked';

    return `
      <tr class="${rowClass}" data-id="${escapeAttr(t.id)}">
        <td class="td-cb">
          <input type="checkbox" class="row-checkbox" data-id="${escapeAttr(t.id)}" ${isChecked}>
        </td>
        <td class="td-title-cell" data-id="${escapeAttr(t.id)}">
          ${escapeHtml(t.title)}
        </td>
        <td><span class="owner-cell">${avatarHtml(t.owner, 'avatar-xs', false)}${escapeHtml(t.owner || 'ללא שיוך')}</span></td>
        <td><span class="priority-badge ${pClass}">${escapeHtml(t.priority)}</span></td>
        <td>${escapeHtml(t.status)}</td>
        <td>${escapeHtml(t.category)}</td>
        <td><span class="chip due-chip due-${due.tone || 'plain'}" title="${escapeAttr(t.dueDate ? fmtDate(t.dueDate) : 'ללא תאריך יעד')}">${escapeHtml(due.text)}</span></td>
        <td><span class="health-pill ${escapeAttr(healthClass)}">${escapeHtml(t.health || HEALTH_ON_TRACK)}</span></td>
        <td>${t.updatedBy ? `<span class="user-pill" title="${escapeAttr(t.updatedAt ? formatWhen(t.updatedAt) : '')}">${escapeHtml(t.updatedBy)}</span>` : '<span class="muted">-</span>'}</td>
        <td>
          <button class="btn-table-edit" data-id="${escapeAttr(t.id)}" title="עריכה" aria-label="עריכת ${escapeAttr(t.title)}">${icon('pencil')}</button>
        </td>
      </tr>
    `;
  }).join('');

  // Attach event listeners
  tbody.querySelectorAll('.row-checkbox').forEach(cb => {
    cb.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleSelection(cb.dataset.id, cb.checked);
    });
  });

  tbody.querySelectorAll('.td-title-cell, .btn-table-edit').forEach(el => {
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      openTask(el.dataset.id);
    });
  });

  tbody.querySelectorAll('tr').forEach(tr => {
    tr.addEventListener('click', (e) => {
      if (e.target.tagName.toLowerCase() === 'input' || e.target.tagName.toLowerCase() === 'button') return;
      openTask(tr.dataset.id);
    });
  });

  // Master checkbox state
  const masterCb = q('#thMasterCheckbox');
  if (masterCb) {
    const allChecked = tasks.length > 0 && tasks.every(t => selectedTasks.has(t.id));
    masterCb.checked = allChecked;
  }

  // Update sort indicators on table headers
  qa('#tasksTable thead th.th-sortable').forEach(th => {
    const key = th.dataset.sort;
    const iconSpan = th.querySelector('.sort-icon');
    if (iconSpan) {
      if (key === tableSortKey) {
        iconSpan.textContent = tableSortAsc ? '▲' : '▼';
        th.classList.add('sorted');
      } else {
        iconSpan.textContent = '';
        th.classList.remove('sorted');
      }
    }
  });
}

/**
 * Render History Table
 */
function renderHistory() {
  const tbody = q('#historyTableBody');
  const emptyState = q('#historyEmptyState');
  if (!tbody || !state) return;

  const history = state.history || [];
  if (history.length === 0) {
    tbody.innerHTML = '';
    if (emptyState) emptyState.style.display = 'block';
    return;
  }

  if (emptyState) emptyState.style.display = 'none';

  tbody.innerHTML = history.slice(0, 30).map(h => `
    <tr>
      <td class="cell-time">${escapeHtml(fmtDateTime(h.time))}</td>
      <td><span class="user-pill">${avatarHtml(h.user || '', 'avatar-xs', false)}${escapeHtml(h.user || 'Excel')}</span></td>
      <td><strong>${escapeHtml(h.action)}</strong></td>
      <td>${escapeHtml(h.id)}</td>
      <td>${escapeHtml(h.title)}</td>
      <td>${escapeHtml(h.details)}</td>
    </tr>
  `).join('');
}

/**
 * Bulk Bar Rendering
 */
function renderBulkBar() {
  const bar = q('#bulkBar');
  const countEl = q('#bulkCount');
  if (!bar || !countEl) return;

  if (selectedTasks.size > 0) {
    countEl.textContent = selectedTasks.size;
    bar.style.display = 'flex';
  } else {
    bar.style.display = 'none';
  }
}

function renderBulkDropdownMenus() {
  if (!state) return;

  const statusMenu = q('#bulkStatusMenu');
  if (statusMenu) {
    statusMenu.innerHTML = (state.statuses || []).map(st => `
      <div class="bulk-menu-item" data-action="status" data-value="${escapeAttr(st)}">${escapeHtml(st)}</div>
    `).join('');
  }

  const ownerMenu = q('#bulkOwnerMenu');
  if (ownerMenu) {
    ownerMenu.innerHTML = (state.people || []).map(p => `
      <div class="bulk-menu-item" data-action="owner" data-value="${escapeAttr(p)}">${escapeHtml(p)}</div>
    `).join('');
  }

  const priorityMenu = q('#bulkPriorityMenu');
  if (priorityMenu) {
    priorityMenu.innerHTML = (state.priorities || []).map(pr => `
      <div class="bulk-menu-item" data-action="priority" data-value="${escapeAttr(pr)}">${escapeHtml(pr)}</div>
    `).join('');
  }

  qa('.bulk-menu-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.stopPropagation();
      const action = item.dataset.action;
      const value = item.dataset.value;
      qa('.bulk-menu').forEach(m => m.classList.remove('show'));
      bulkAction(action, value);
    });
  });
}

/**
 * Subtasks Checklist Helpers (Parsed safely to/from task.notes)
 */
// Links live in the notes as lines starting with 🔗, so they need no extra Excel column.
const LINK_LINE = /^\s*🔗\s*(.*?)\s*((?:https?:\/\/|mailto:)\S+)\s*$/u;

function parseNotesAndSubtasks(notesText) {
  if (!notesText) return { subtasks: [], links: [], cleanNotes: '' };
  const lines = String(notesText).split('\n');
  const subtasks = [];
  const links = [];
  const otherLines = [];

  for (const line of lines) {
    const match = line.match(/^\s*(?:[-*]\s*)?\[([ xX])\]\s*(.*)$/);
    const link = match ? null : line.match(LINK_LINE);
    if (match) {
      subtasks.push({
        done: match[1].toLowerCase() === 'x',
        text: match[2].trim()
      });
    } else if (link) {
      links.push({ title: link[1].trim(), url: link[2] });
    } else {
      otherLines.push(line);
    }
  }
  return { subtasks, links, cleanNotes: otherLines.join('\n').trim() };
}

function serializeNotesAndSubtasks(cleanNotes, subtasks, links = []) {
  const parts = [];
  for (const st of subtasks || []) {
    if (st.text) parts.push(`- [${st.done ? 'x' : ' '}] ${st.text}`);
  }
  for (const link of links || []) {
    parts.push(`🔗 ${link.title ? link.title + ' ' : ''}${link.url}`);
  }
  if (cleanNotes && cleanNotes.trim()) {
    if (parts.length > 0) parts.push('');
    parts.push(cleanNotes.trim());
  }
  return parts.join('\n');
}

function safeUrl(url) {
  return /^(https?:\/\/|mailto:)/i.test(String(url || '')) ? url : '';
}

function shortUrl(url) {
  try {
    const parsed = new URL(url);
    if (parsed.protocol === 'mailto:') return parsed.pathname;
    const last = parsed.pathname.split('/').filter(Boolean).pop();
    let name = last || '';
    try { name = decodeURIComponent(name); } catch { /* keep encoded */ }
    return name ? `${parsed.hostname} › ${name}` : parsed.hostname;
  } catch {
    return url;
  }
}

function renderDrawerLinks() {
  const list = q('#linksList');
  if (!list) return;
  if (!currentDrawerLinks.length) {
    list.innerHTML = '<div class="muted-empty">אין קישורים. אפשר לצרף מסמך SharePoint, אתר או כתובת מייל.</div>';
    return;
  }
  list.innerHTML = currentDrawerLinks.map((link, idx) => `
    <div class="link-item">
      <span class="link-icon" aria-hidden="true">🔗</span>
      <a href="${escapeAttr(safeUrl(link.url))}" target="_blank" rel="noopener noreferrer" class="link-anchor" title="${escapeAttr(link.url)}">${escapeHtml(link.title || shortUrl(link.url))}</a>
      <button type="button" class="link-del-btn" data-idx="${idx}" title="הסרת קישור" aria-label="הסרת קישור">&times;</button>
    </div>
  `).join('');
  list.querySelectorAll('.link-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentDrawerLinks.splice(parseInt(btn.dataset.idx, 10), 1);
      renderDrawerLinks();
    });
  });
}

function addDrawerLink() {
  const urlInput = q('#newLinkUrl');
  const titleInput = q('#newLinkTitle');
  let url = (urlInput.value || '').trim();
  if (!url) return;
  if (/^[\w.-]+@[\w.-]+\.\w+$/.test(url)) url = 'mailto:' + url;
  else if (/^[a-z][a-z0-9+.-]*:/i.test(url) && !/^(https?:\/\/|mailto:)/i.test(url)) url = '';  // other schemes
  else if (!/^(https?:\/\/|mailto:)/i.test(url)) url = 'https://' + url;
  if (/\s/.test(url) || !safeUrl(url)) {
    toast('כתובת הקישור אינה תקינה');
    return;
  }
  currentDrawerLinks.push({ title: (titleInput.value || '').trim().replace(/\s+/g, ' '), url });
  urlInput.value = '';
  titleInput.value = '';
  renderDrawerLinks();
}

function formatWhen(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (!m) return value || '';
  const when = new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  const minutes = Math.round((Date.now() - when.getTime()) / 60000);
  if (m[4] && minutes >= 0 && minutes < 1) return 'עכשיו';
  if (m[4] && minutes >= 0 && minutes < 60) return `לפני ${minutes} דק׳`;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = new Date(when);
  day.setHours(0, 0, 0, 0);
  const days = Math.round((today - day) / 86400000);
  const time = m[4] ? ` ${m[4]}:${m[5]}` : '';
  if (days === 0) return `היום${time}`;
  if (days === 1) return `אתמול${time}`;
  return `${m[3]}/${m[2]}/${m[1]}${time}`;
}

function renderActivity(task) {
  const group = q('#activityGroup');
  const list = q('#activityList');
  if (!group || !list) return;
  const tab = q('#tabActivity');
  if (!task) {
    group.hidden = true;
    tab.hidden = true;
    q('#tabCountActivity').textContent = '';
    if (q('.drawer-tab.active')?.dataset.tab === 'activity') setDrawerTab('details');
    return;
  }
  group.hidden = false;
  tab.hidden = false;
  // SharePoint keeps events with the task; local mode also has the history sheet. Merge both.
  const seen = new Set();
  const events = [];
  for (const e of [...(task.activity || []), ...(state.history || []).filter(h => h.id === task.id)]) {
    const key = [e.time, e.user, e.action, e.comment || e.details || ''].join('|');
    if (seen.has(key)) continue;
    seen.add(key);
    events.push(e);
  }
  events.sort((a, b) => String(b.time || '').localeCompare(String(a.time || '')));
  const commentTotal = events.filter(e => e.comment).length;
  q('#tabCountActivity').textContent = commentTotal ? String(commentTotal) : '';
  list.innerHTML = events.length ? events.slice(0, 40).map(e => e.comment ? `
    <div class="activity-item is-comment">
      ${avatarHtml(e.user || '', 'activity-avatar', false)}
      <div class="activity-body">
        <div class="activity-meta"><strong>${escapeHtml(e.user || '')}</strong><span>${escapeHtml(formatWhen(e.time))}</span></div>
        <div class="activity-comment">${escapeHtml(e.comment)}</div>
      </div>
    </div>` : `
    <div class="activity-item">
      <span class="activity-dot" aria-hidden="true"></span>
      <div class="activity-body">
        <div class="activity-meta"><strong>${escapeHtml(e.user || 'Excel')}</strong><span>${escapeHtml(e.action || '')}</span><span>${escapeHtml(formatWhen(e.time))}</span></div>
        ${e.details ? `<div class="activity-details">${escapeHtml(e.details)}</div>` : ''}
      </div>
    </div>`).join('') : '<div class="muted-empty">אין פעילות עדיין.</div>';
}

async function addComment() {
  const id = q('#formTaskId').value;
  const box = q('#newCommentText');
  const text = (box.value || '').trim();
  if (!id || !text) return;
  const button = q('#btnAddComment');
  button.disabled = true;
  try {
    await api('/api/comment', { method: 'POST', body: { id, text, user: getCurrentUser(), workbook: state.workbook } });
    box.value = '';
    await load();
    drawerRevision = state.revision;  // Only our comment changed the file.
    renderActivity(state.tasks.find(t => t.id === id));
    toast('התגובה נשמרה');
  } catch (err) {
    toast(err.message);
  } finally {
    button.disabled = false;
  }
}

function initDrawerSubtasks(notesText) {
  const parsed = parseNotesAndSubtasks(notesText);
  currentDrawerCleanNotes = parsed.cleanNotes;
  currentDrawerSubtasks = parsed.subtasks;
  currentDrawerLinks = parsed.links;
  renderDrawerSubtasks();
  renderDrawerLinks();
}

function updateChecklistProgress() {
  const total = currentDrawerSubtasks.length;
  const pct = total ? Math.round(currentDrawerSubtasks.filter(s => s.done).length / total * 100) : 0;
  q('#formTaskProgress').value = pct;
  q('#formProgressValText').textContent = `${pct}%`;
  const select = q('#formTaskStatus');
  if (pct === 100) {
    if (![...select.options].some(o => o.value === DONE)) select.add(new Option(DONE, DONE));
    select.value = DONE;
  } else if (select.value === DONE) {
    const next = (state.statuses || []).find(s => s !== DONE && s !== CANCELLED) || 'בתהליך';
    if (![...select.options].some(o => o.value === next)) select.add(new Option(next, next));
    select.value = next;
  }
}

function renderDrawerSubtasks() {
  const listEl = q('#subtasksList');
  const counterEl = q('#subtasksCounter');
  if (!listEl) return;

  const total = currentDrawerSubtasks.length;
  const done = currentDrawerSubtasks.filter(s => s.done).length;
  const pct = total ? Math.round((done / total) * 100) : 0;
  q('#tabCountChecklist').textContent = total ? `${done}/${total}` : '';
  q('#subtasksProgressFill').style.width = `${pct}%`;
  if (total === 0) {
    listEl.innerHTML = '<div class="muted-empty">אין משימות משנה. מפרקים את המשימה לצעדים קטנים, והאחוז מתעדכן לבד.</div>';
    if (counterEl) counterEl.textContent = '0/0 הושלמו (0%)';
    return;
  }
  if (counterEl) counterEl.textContent = `${done}/${total} הושלמו (${pct}%)`;

  listEl.innerHTML = currentDrawerSubtasks.map((st, idx) => `
    <div class="subtask-item ${st.done ? 'completed' : ''}" data-idx="${idx}">
      <input type="checkbox" class="subtask-cb" data-idx="${idx}" ${st.done ? 'checked' : ''}>
      <span class="subtask-text">${escapeHtml(st.text)}</span>
      <button type="button" class="subtask-del-btn" data-idx="${idx}" title="מחיקת סעיף">&times;</button>
    </div>
  `).join('');

  listEl.querySelectorAll('.subtask-cb').forEach(cb => {
    cb.addEventListener('change', () => {
      const idx = parseInt(cb.dataset.idx, 10);
      if (currentDrawerSubtasks[idx]) {
        currentDrawerSubtasks[idx].done = cb.checked;
        updateChecklistProgress();
        renderDrawerSubtasks();
      }
    });
  });

  listEl.querySelectorAll('.subtask-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      currentDrawerSubtasks.splice(idx, 1);
      updateChecklistProgress();
      renderDrawerSubtasks();
    });
  });
}

/**
 * Drawer Operations
 */
function setDueField(iso) {
  q('#formTaskDueDate').value = iso || '';
  const text = q('#formTaskDueDateText');
  text.value = iso ? fmtDate(iso) : '';
  text.classList.remove('is-invalid');
}

// Reads the typed date into the stored field. False (and a message) when it is not a real date.
function syncDueFromText() {
  const text = q('#formTaskDueDateText');
  const iso = parseDisplayDate(text.value);
  if (iso === null) {
    text.classList.add('is-invalid');
    toast('תאריך היעד אינו תקין. יש לכתוב יום/חודש/שנה, למשל 30/11/2026');
    return false;
  }
  setDueField(iso);
  return true;
}

function setDrawerTab(name) {
  qa('.drawer-tab').forEach(tab => {
    const on = tab.dataset.tab === name;
    tab.classList.toggle('active', on);
    tab.setAttribute('aria-selected', String(on));
  });
  qa('.drawer-panel').forEach(panel => panel.classList.toggle('active', panel.dataset.panel === name));
}

function openNewTask() {
  if (cloudConnectPending) { toast('יש להמתין לסיום החיבור ל-Excel'); return; }

  if (!state) return;
  fillSelect(q('#formTaskOwner'), ['', ...(state.people || [])]);
  q('#formTaskOwner').value = state.people[0] || '';
  fillSelect(q('#formTaskPriority'), state.priorities || []);
  fillSelect(q('#formTaskStatus'), state.statuses || []);
  fillSelect(q('#formTaskCategory'), ['', ...(state.categories || [])]);
  q('#formTaskCategory').value = state.categories[0] || '';
  q('#formTaskPriority').value = (state.priorities || []).includes('P2') ? 'P2' : state.priorities[0];
  q('#formTaskStatus').value = state.statuses[0] || '';
  q('#formTaskId').value = '';
  q('#formTaskTitle').value = '';
  setDueField('');
  q('#formTaskBlocked').value = NO;
  q('#formTaskProgress').value = 0;
  q('#formProgressValText').textContent = '0%';
  q('#formTaskNotes').value = '';

  initDrawerSubtasks('');
  renderActivity(null);

  const auditBox = q('#taskAuditBox');
  if (auditBox) auditBox.style.display = 'none';

  q('#drawerTitle').textContent = 'משימה חדשה';
  q('#drawerKicker').textContent = '';
  const btnDelete = q('#formBtnDelete');
  if (btnDelete) btnDelete.style.visibility = 'hidden';

  setDrawerTab('details');
  openDrawer();
  setTimeout(() => { if (editing && !q('#formTaskId').value) q('#formTaskTitle').focus(); }, 80);
}

function openNewTaskFor(owner) {
  if (cloudConnectPending) { toast('יש להמתין לסיום החיבור ל-Excel'); return; }
  openNewTask();
  if (owner && q('#formTaskOwner')) {
    if (![...q('#formTaskOwner').options].some(o => o.value === owner)) {
      q('#formTaskOwner').add(new Option(owner, owner));
    }
    q('#formTaskOwner').value = owner;
  }
}

function openTask(id) {
  if (cloudConnectPending) { toast('יש להמתין לסיום החיבור ל-Excel'); return; }
  if (!state || !state.tasks) return;
  const task = state.tasks.find(t => t.id === id || t.title === id);
  if (!task) return;

  q('#formTaskId').value = task.id;
  q('#formTaskTitle').value = task.title || '';

  const ownerSel = q('#formTaskOwner');
  if (ownerSel && task.owner && ![...ownerSel.options].some(o => o.value === task.owner)) {
    const opt = document.createElement('option');
    opt.value = task.owner;
    opt.textContent = `${task.owner} (לא ברשימה)`;
    ownerSel.appendChild(opt);
  }
  if (ownerSel) ownerSel.value = task.owner || '';

  q('#formTaskPriority').value = task.priority || 'P2';
  q('#formTaskStatus').value = task.status || '';
  for (const [selector, value] of [['#formTaskPriority', task.priority], ['#formTaskStatus', task.status]]) {
    const select = q(selector);
    if (value && ![...select.options].some(o => o.value === value)) {
      select.add(new Option(value, value));
    }
    select.value = value || '';
  }

  const catSel = q('#formTaskCategory');
  if (catSel && task.category && ![...catSel.options].some(o => o.value === task.category)) {
    const opt = document.createElement('option');
    opt.value = task.category;
    opt.textContent = `${task.category} (לא ברשימה)`;
    catSel.appendChild(opt);
  }
  if (catSel) catSel.value = task.category || '';
  setDueField(task.dueDate || '');
  q('#formTaskBlocked').value = task.blocked === YES ? YES : NO;
  q('#formTaskProgress').value = task.progress || 0;
  q('#formProgressValText').textContent = `${task.progress || 0}%`;
  initDrawerSubtasks(task.notes || '');
  q('#formTaskNotes').value = currentDrawerCleanNotes;
  renderActivity(task);

  const auditBox = q('#taskAuditBox');
  if (auditBox) {
    if (task.createdBy || task.updatedBy) {
      auditBox.style.display = 'flex';
      const cBy = q('#drawerCreatedBy');
      const cAt = q('#drawerCreatedAt');
      const uBy = q('#drawerUpdatedBy');
      const uAt = q('#drawerUpdatedAt');
      if (cBy) cBy.textContent = task.createdBy || 'Excel';
      if (cAt) cAt.textContent = task.createdAt ? `(${fmtDateTime(task.createdAt)})` : '';
      if (uBy) uBy.textContent = task.updatedBy || task.createdBy || '-';
      if (uAt) uAt.textContent = task.updatedAt ? `(${fmtDateTime(task.updatedAt)})` : '';
    } else {
      auditBox.style.display = 'none';
    }
  }

  q('#drawerTitle').textContent = 'עריכת משימה';
  q('#drawerKicker').textContent = task.id;
  const btnDelete = q('#formBtnDelete');
  if (btnDelete) btnDelete.style.visibility = 'visible';

  setDrawerTab('details');
  openDrawer();
}

function openDrawer() {
  editing = true;
  drawerRevision = state.revision;
  drawerWorkbook = state.workbook;
  drawerTaskRevision = (state.tasks || []).find(t => t.id === q('#formTaskId')?.value)?.taskRevision || '';
  const drawer = q('#taskDrawer');
  const backdrop = q('#drawerBackdrop');
  if (drawer) {
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden', 'false');
  }
  if (backdrop) backdrop.classList.add('show');
}

function closeDrawer() {
  editing = false;
  const drawer = q('#taskDrawer');
  const backdrop = q('#drawerBackdrop');
  if (drawer) {
    drawer.classList.remove('open');
    drawer.setAttribute('aria-hidden', 'true');
  }
  if (backdrop) backdrop.classList.remove('show');
}

async function saveTask(event) {
  if (event) event.preventDefault();
  if (savingTask) return;
  const title = (q('#formTaskTitle').value || '').trim();
  if (!title) {
    toast('חסר שם משימה');
    return;
  }

  if (!syncDueFromText()) {
    setDrawerTab('details');
    q('#formTaskDueDateText').focus();
    return;
  }

  const payload = {
    id: q('#formTaskId').value || '',
    title: title,
    owner: q('#formTaskOwner').value,
    priority: q('#formTaskPriority').value,
    status: q('#formTaskStatus').value,
    category: q('#formTaskCategory').value,
    dueDate: q('#formTaskDueDate').value,
    blocked: q('#formTaskBlocked').value,
    progress: q('#formTaskProgress').value,
    notes: serializeNotesAndSubtasks(q('#formTaskNotes').value, currentDrawerSubtasks, currentDrawerLinks),
    user: getCurrentUser(),
    revision: drawerRevision,
    workbook: drawerWorkbook,
    taskRevision: drawerTaskRevision,
  };

  savingTask = true;
  try {
    await api('/api/task', { method: 'POST', body: payload });
    closeDrawer();
    toast('המשימה נשמרה באקסל');
    await load();
  } catch (err) {
    toast(err.message);
  } finally {
    savingTask = false;
  }
}

async function deleteCurrentTask() {
  const id = q('#formTaskId').value;
  if (!id) return;
  if (!confirm('למחוק את המשימה?')) return;

  try {
    await api('/api/delete', { method: 'POST', body: { id, user: getCurrentUser(), revision: drawerRevision, workbook: drawerWorkbook } });
    selectedTasks.delete(id);
    closeDrawer();
    toast('המשימה נמחקה');
    await load();
  } catch (err) {
    toast(err.message);
  }
}

async function markDone(id) {
  if (!id) return;
  try {
    await api('/api/done', { method: 'POST', body: { id, user: getCurrentUser() } });
    fireConfetti();
    toast('המשימה סומנה כבוצעה! 🎉');
    await load();
  } catch (err) {
    toast(err.message);
    render();
  }
}

async function bulkAction(action, value = '') {
  if (selectedTasks.size === 0) {
    toast('לא נבחרו משימות');
    return;
  }

  const count = selectedTasks.size;
  if (action === 'delete') {
    if (!confirm(`למחוק ${count} משימות?`)) return;
  }

  try {
    const res = await api('/api/bulk', {
      method: 'POST',
      body: {
        ids: [...selectedTasks],
        action,
        value,
        user: getCurrentUser(),
      }
    });
    if (action === 'done') fireConfetti();
    selectedTasks.clear();
    const skipped = (res && res.failed) ? res.failed.length : 0;
    toast(skipped
      ? `${res.count} משימות עודכנו. ${skipped} השתנו בינתיים או נמחקו ולא עודכנו.`
      : `${count} משימות עודכנו באקסל`);
    await load();
  } catch (err) {
    toast(err.message);
  }
}

async function persistOrder(orderedIds) {
  try {
    await api('/api/reorder', { method: 'POST', body: { ids: orderedIds, user: getCurrentUser() } });
    toast('הסדר נשמר');
    await load();
  } catch (err) {
    toast(err.message);
  }
}

async function saveLists() {
  const saveBtn = q('#btnSaveSettings');
  if (saveBtn && saveBtn.disabled) return;
  const inputs = qa('#settingsForm textarea');
  inputs.forEach(input => input.disabled = true);
  const originalText = saveBtn ? saveBtn.textContent : 'שמירת שינויים באקסל';
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = 'שומר באקסל... ⏳';
  }

  const people = q('#settingsPeople').value;
  const categories = q('#settingsCategories').value;
  const statuses = q('#settingsStatuses').value;
  const priorities = q('#settingsPriorities').value;

  try {
    const res = await api('/api/lists', {
      method: 'POST',
      body: { people, categories, statuses, priorities, user: getCurrentUser(),
        revision: settingsDirty ? settingsRevision : state.revision,
        workbook: settingsDirty ? settingsWorkbook : state.workbook,
        listsRevision: settingsDirty ? settingsListsRevision : (state.listsRevision || '') }
    });
    settingsDirty = false;
    toast('ההגדרות והרשימות נשמרו באקסל בהצלחה! ✓');
    if (res && res.data) {
      state = res.data;
      lastPayload = JSON.stringify(res.data);
      render();
    } else {
      await load();
    }
  } catch (err) {
    toast(err.message || 'שגיאה בשמירת הרשימות');
  } finally {
    inputs.forEach(input => input.disabled = false);
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.textContent = originalText;
    }
  }
}

function setView(view) {
  currentView = view;

  // Nav buttons active state
  qa('.nav-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === view);
  });

  // Views display active state
  qa('.view').forEach(v => {
    v.classList.remove('active');
  });

  const targetView = q(`#${view}View`);
  if (targetView) targetView.classList.add('active');

  // Update Topbar Heading
  const meta = viewMeta[view] || [view, ''];
  const titleEl = q('#viewTitle');
  const subEl = q('#viewSubtitle');
  if (titleEl) titleEl.textContent = meta[0];
  if (subEl) subEl.textContent = meta[1];

  render();
}

function applyMetricFilter(filter) {
  for (const selector of ['#filterOwner', '#filterStatus', '#filterPriority']) {
    q(selector).value = ALL;
  }
  metricFilter = filter;
  setView('tasks');
}

function toggleSelection(id, checked) {
  if (checked) {
    selectedTasks.add(id);
  } else {
    selectedTasks.delete(id);
  }
  renderBulkBar();

  // Update row checkbox if in table view
  const rowCb = q(`.row-checkbox[data-id="${CSS.escape(id)}"]`);
  if (rowCb) rowCb.checked = checked;

  // Update card checkbox if in board view
  const cardCb = q(`.card-checkbox[data-id="${CSS.escape(id)}"]`);
  if (cardCb) cardCb.checked = checked;
  cardCb?.closest('.task-card')?.classList.toggle('is-selected', checked);
  rowCb?.closest('tr')?.classList.toggle('is-selected', checked);
}

function clearSelection() {
  selectedTasks.clear();
  renderBulkBar();
  qa('.row-checkbox').forEach(cb => cb.checked = false);
  qa('.card-checkbox').forEach(cb => cb.checked = false);
  qa('.is-selected').forEach(el => el.classList.remove('is-selected'));
  const master = q('#thMasterCheckbox');
  if (master) master.checked = false;
}

/**
 * Celebration Confetti Burst
 */
function fireConfetti() {
  const canvas = q('#confettiCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  if (!ctx) return;
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;

  const particles = [];
  const colors = ['#0284c7', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#3b82f6', '#14b8a6'];
  for (let i = 0; i < 70; i++) {
    particles.push({
      x: canvas.width / 2,
      y: canvas.height * 0.4,
      vx: (Math.random() - 0.5) * 16,
      vy: (Math.random() - 0.8) * 14,
      size: Math.random() * 8 + 4,
      color: colors[Math.floor(Math.random() * colors.length)],
      rotation: Math.random() * 360,
      rotSpeed: (Math.random() - 0.5) * 12,
      life: 1,
      decay: Math.random() * 0.02 + 0.015,
    });
  }

  function step() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    let alive = false;
    for (const p of particles) {
      p.x += p.vx;
      p.y += p.vy;
      p.vy += 0.38;
      p.rotation += p.rotSpeed;
      p.life -= p.decay;
      if (p.life > 0) {
        alive = true;
        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate((p.rotation * Math.PI) / 180);
        ctx.globalAlpha = Math.max(0, p.life);
        ctx.fillStyle = p.color;
        ctx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size * 0.6);
        ctx.restore();
      }
    }
    if (alive) {
      requestAnimationFrame(step);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
  }
  step();
}

/**
 * Command Palette (Ctrl+K)
 */
let paletteSelectedIndex = 0;
let paletteItems = [];

function openPalette() {
  const backdrop = q('#paletteBackdrop');
  const input = q('#paletteInput');
  if (!backdrop || !input) return;
  backdrop.style.display = 'flex';
  input.value = '';
  paletteSelectedIndex = 0;
  renderPaletteResults('');
  setTimeout(() => input.focus(), 50);
}

function closePalette() {
  const backdrop = q('#paletteBackdrop');
  if (backdrop) backdrop.style.display = 'none';
}

function renderPaletteResults(query) {
  const container = q('#paletteResults');
  if (!container) return;
  const qStr = (query || '').trim().toLowerCase();

  const standardActions = [
    { type: 'action', title: 'משימה חדשה', desc: 'פתיחת טופס ליצירת משימה חדשה', icon: '➕', run: () => openNewTask() },
    { type: 'action', title: 'מעבר ללוח הבקרה', desc: 'מצב הצוות, עומס ומשימות שדורשות טיפול', icon: '📊', run: () => setView('dashboard') },
    { type: 'action', title: 'מעבר לציר הזמן', desc: 'משימות לפי תאריכי התחלה ויעד', icon: '📅', run: () => setView('timeline') },
    { type: 'action', title: 'מעבר ללוח המשימות', desc: 'עמודה לכל אחד מהצוות', icon: '📋', run: () => setView('board') },
    { type: 'action', title: 'מעבר לטבלת משימות', desc: 'טבלה מפורטת עם כל הנתונים וההיסטוריה', icon: '📑', run: () => setView('tasks') },
    { type: 'action', title: 'המשימות שלי', desc: 'סינון מהיר של משימות המשויכות אלי', icon: '👤', run: () => applyMetricFilter('my-tasks') },
    { type: 'action', title: 'שיתוף סטטוס ב-WhatsApp', desc: 'ניסוח סיכום מוכן לשליחה מהירה', icon: '💬', run: () => openShareModal() },
    { type: 'action', title: 'פתח קובץ באקסל ישירות', desc: 'הפעלת Microsoft Excel עם הקובץ הפעיל', icon: '↗️', run: () => api('/api/open-excel', { method: 'POST' }).then(() => toast('נפתח ב-Excel ↗')).catch(e => toast(e.message)) },
    { type: 'action', title: 'היום שלי', desc: 'המשימות שלי לפי מועד', icon: '☀️', run: () => setView('myday') },
    { type: 'action', title: 'לוח שנה', desc: 'משימות לפי תאריך יעד בתצוגה חודשית', icon: '🗓️', run: () => setView('calendar') },
    { type: 'action', title: 'דוח שבועי', desc: 'סיכום להדפסה, PDF או Excel', icon: '📈', run: () => openReport() },
    { type: 'action', title: 'מעבר להגדרות', desc: 'ניהול רשימות עובדים, קטגוריות וסטטוסים', icon: '⚙️', run: () => setView('settings') },
    { type: 'action', title: 'רענון נתונים מהאקסל', desc: 'סנכרון מיידי מהקובץ בדיסק', icon: '🔄', run: () => refreshFromDisk(false) },
    { type: 'action', title: 'סינון: משימות באיחור', desc: 'צפייה רק במשימות שעבר תאריך היעד שלהן', icon: '🚨', run: () => applyMetricFilter('overdue') },
    { type: 'action', title: 'סינון: משימות חסומות', desc: 'צפייה במשימות המסומנות כחסומות', icon: '⛔', run: () => applyMetricFilter('blocked') },
  ];

  let filteredActions = standardActions.filter(a => !qStr || a.title.toLowerCase().includes(qStr) || a.desc.toLowerCase().includes(qStr));

  let matchedTasks = [];
  if (state && state.tasks && qStr) {
    matchedTasks = state.tasks
      .filter(t => (t.title || '').toLowerCase().includes(qStr) || (t.owner || '').toLowerCase().includes(qStr) || (t.category || '').toLowerCase().includes(qStr) || (t.notes || '').toLowerCase().includes(qStr))
      .slice(0, 8)
      .map(t => ({
        type: 'task',
        title: t.title,
        desc: `${t.owner} • ${t.priority} • ${t.status} • יעד: ${t.dueDate ? fmtDate(t.dueDate) : 'ללא'}`,
        icon: '📌',
        run: () => openTask(t.id),
      }));
  }

  paletteItems = [...filteredActions, ...matchedTasks];
  if (paletteSelectedIndex >= paletteItems.length) paletteSelectedIndex = 0;

  if (paletteItems.length === 0) {
    container.innerHTML = '<div class="palette-empty">לא נמצאו תוצאות לפקודה זו</div>';
    return;
  }

  container.innerHTML = paletteItems.map((item, idx) => {
    const isSel = idx === paletteSelectedIndex ? 'selected' : '';
    return `
      <div class="palette-item ${isSel}" data-index="${idx}">
        <div class="palette-item-icon">${item.icon}</div>
        <div class="palette-item-content">
          <div class="palette-item-title">${escapeHtml(item.title)}</div>
          <div class="palette-item-desc">${escapeHtml(item.desc)}</div>
        </div>
      </div>
    `;
  }).join('');

  container.querySelectorAll('.palette-item').forEach(el => {
    el.addEventListener('click', () => {
      const idx = parseInt(el.dataset.index, 10);
      if (paletteItems[idx]) {
        closePalette();
        paletteItems[idx].run();
      }
    });
  });
}

/**
 * Shortcuts Modal (?)
 */
function openShortcuts() {
  const el = q('#shortcutsBackdrop');
  if (el) el.style.display = 'flex';
}

function closeShortcuts() {
  const el = q('#shortcutsBackdrop');
  if (el) el.style.display = 'none';
}

/**
 * Date and Periodic Reports Analytics
 */
function parseIsoDate(str) {
  if (!str) return null;
  const match = String(str).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  const parsed = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return parsed.getFullYear() === Number(match[1]) && parsed.getMonth() === Number(match[2]) - 1
    && parsed.getDate() === Number(match[3]) ? parsed : null;
}


/**
 * Timeline / Gantt Roadmap View
 */
function renderTimeline() {
  const container = q('#timelineScrollWrap');
  if (!container || !state) return;

  const tasks = visibleTasks();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dayCount = 14;
  const startDay = new Date(today);
  startDay.setDate(today.getDate() - 3);

  const days = [];
  const hebrewDays = ['א\'', 'ב\'', 'ג\'', 'ד\'', 'ה\'', 'ו\'', 'ש\''];
  for (let i = 0; i < dayCount; i++) {
    const cur = new Date(startDay);
    cur.setDate(startDay.getDate() + i);
    const isToday = cur.getTime() === today.getTime();
    days.push({
      date: cur,
      dayName: hebrewDays[cur.getDay()],
      dayNum: cur.getDate(),
      monthNum: cur.getMonth() + 1,
      isToday,
      iso: `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, '0')}-${String(cur.getDate()).padStart(2, '0')}`
    });
  }

  const headerHtml = `
    <div class="timeline-header-days">
      <div class="tl-corner">משימה</div>
      ${days.map(d => `
        <div class="timeline-day-col ${d.isToday ? 'today-col' : ''}">
          <div>${d.dayName}</div>
          <div class="tl-date">${String(d.dayNum).padStart(2, '0')}/${String(d.monthNum).padStart(2, '0')}</div>
        </div>
      `).join('')}
    </div>
  `;

  const groups = Object.create(null);
  if (timelineGroupBy === 'category') {
    (state.categories || []).forEach(cat => groups[cat] = []);
    tasks.forEach(t => {
      const k = t.category || 'ללא קטגוריה';
      if (!groups[k]) groups[k] = [];
      groups[k].push(t);
    });
  } else if (timelineGroupBy === 'status') {
    (state.statuses || []).forEach(st => groups[st] = []);
    tasks.forEach(t => {
      const k = t.status || 'אחר';
      if (!groups[k]) groups[k] = [];
      groups[k].push(t);
    });
  } else {
    (state.people || []).forEach(p => groups[p] = []);
    tasks.forEach(t => {
      const k = t.owner || 'לא משויך';
      if (!groups[k]) groups[k] = [];
      groups[k].push(t);
    });
  }

  const groupRowsHtml = Object.entries(groups).map(([grpName, grpTasks]) => {
    if (!grpTasks || grpTasks.length === 0) return '';

    const tasksHtml = grpTasks.map(t => {
      const tDue = parseIsoDate(t.dueDate);
      const tStart = parseIsoDate(t.startDate) || (tDue ? new Date(tDue.getTime() - 86400000 * 2) : today);

      let startIdx = 0;
      let spanDays = 2;
      let inRange = false;

      if (tDue) {
        const msPerDay = 86400000;
        const diffStart = Math.round((Math.min(tStart.getTime(), tDue.getTime()) - startDay.getTime()) / msPerDay);
        const diffEnd = Math.round((tDue.getTime() - startDay.getTime()) / msPerDay);
        inRange = diffEnd >= 0 && diffStart < dayCount;
        startIdx = Math.max(0, Math.min(dayCount - 1, diffStart));
        const endIdx = Math.max(0, Math.min(dayCount - 1, diffEnd));
        spanDays = Math.max(1, (endIdx - startIdx) + 1);
      }

      const leftPct = (startIdx / dayCount) * 100;
      const widthPct = Math.min(100 - leftPct, (spanDays / dayCount) * 100);

      const hClass = (t.status === DONE) ? 'ontrack' : (t.health === HEALTH_OVERDUE ? 'overdue' : (t.health === HEALTH_BLOCKED ? 'blocked' : (t.health === HEALTH_SOON ? 'soon' : 'ontrack')));

      return `
        <div class="timeline-task-row">
          <div class="timeline-task-name" title="${escapeAttr(t.title)}" data-id="${escapeAttr(t.id)}">
            <span class="priority-badge sm ${priorityClass[t.priority] || 'p2'}">${escapeHtml(t.priority)}</span>
            ${escapeHtml(t.title)}
          </div>
          <div class="timeline-bar-slot">
            ${inRange ? `<div class="timeline-bar ${hClass}" data-id="${escapeAttr(t.id)}" style="right:${leftPct}%;width:${widthPct}%;">
              <span class="tl-bar-label">${clampPercent(t.progress)}% · ${escapeHtml(t.title)}</span>
            </div>` : `<span class="timeline-out-of-range">${tDue ? 'מחוץ לטווח התאריכים המוצג' : 'ללא תאריך יעד'}</span>`}
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="timeline-group-row">
        <div class="timeline-group-header">
          ${timelineGroupBy === 'owner' ? avatarHtml(grpName === 'לא משויך' ? '' : grpName, 'user-avatar-sm', false) : `<span class="user-avatar-sm">${escapeHtml(initials(grpName))}</span>`}
          <span>${escapeHtml(grpName)}</span>
          <span class="timeline-group-count">${grpTasks.length}</span>
        </div>
        ${tasksHtml}
      </div>
    `;
  }).filter(Boolean).join('');

  container.innerHTML = `
    <div class="timeline-grid">
      ${headerHtml}
      ${groupRowsHtml || '<div class="empty-state">אין משימות להצגה בטווח ציר הזמן</div>'}
    </div>
  `;

  container.querySelectorAll('.timeline-task-name, .timeline-bar').forEach(el => {
    el.addEventListener('click', () => {
      const id = el.dataset.id;
      if (id) openTask(id);
    });
  });
}

/**
 * WhatsApp & Email Share Digest Modal
 */
function generateShareDigest() {
  const allTasks = state && state.tasks ? state.tasks : [];
  const done = allTasks.filter(t => t.status === DONE);
  const open = allTasks.filter(t => t.status !== DONE && t.status !== CANCELLED);
  const overdue = open.filter(t => t.health === HEALTH_OVERDUE);
  const blocked = open.filter(t => t.health === HEALTH_BLOCKED);
  const urgent = open.filter(t => t.priority === 'P0' || t.priority === 'P1');
  const user = getCurrentUser();

  const total = allTasks.length;
  const pct = total > 0 ? Math.round((done.length / total) * 100) : 0;

  const now = new Date();
  const dateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

  let text = `📊 *סטטוס משימות צוות ופרויקטים - TeamFlow*\n`;
  text += `📅 תאריך: *${dateStr}*\n`;
  text += `👤 מדווח ע"י: *${user}*\n\n`;
  text += `📈 *מדדי ביצוע מרכזיים:*\n`;
  text += `• סה"כ משימות רשומות: ${total}\n`;
  text += `• הושלמו בהצלחה: ${done.length} (${pct}%)\n`;
  text += `• משימות בטיפול פעיל: ${open.length}\n`;
  if (overdue.length > 0) text += `• ⚠️ באיחור מלו"ז: ${overdue.length}\n`;
  if (blocked.length > 0) text += `• 🛑 חסומות: ${blocked.length}\n`;
  text += `\n`;

  if (urgent.length > 0) {
    text += `🔥 *משימות דחופות בטיפול (P0 / P1):*\n`;
    for (const t of urgent.slice(0, 5)) {
      text += `• [${t.priority}] ${t.title} (${t.owner || 'לא משויך'}${t.dueDate ? ' | יעד: ' + fmtDate(t.dueDate) : ''})\n`;
    }
    text += `\n`;
  }

  text += `📌 *הופק אוטומטית מתוך TeamFlow / קובץ אקסל*`;
  return text;
}

function openShareModal() {
  const text = generateShareDigest();
  const textarea = q('#shareDigestText');
  if (textarea) textarea.value = text;

  const waBtn = q('#btnOpenWhatsApp');
  if (waBtn) {
    waBtn.href = `https://api.whatsapp.com/send?text=${encodeURIComponent(text)}`;
  }

  const mailBtn = q('#btnOpenMailto');
  if (mailBtn) {
    const subject = encodeURIComponent('דוח סטטוס משימות צוות - TeamFlow');
    mailBtn.href = `mailto:?subject=${subject}&body=${encodeURIComponent(text)}`;
  }

  const modal = q('#shareModalBackdrop');
  if (modal) modal.style.display = 'flex';
}

function closeShareModal() {
  const modal = q('#shareModalBackdrop');
  if (modal) modal.style.display = 'none';
}

/**
 * Workbook Management Modal
 */
async function openWorkbookModal() {
  const modal = q('#wbModalBackdrop');
  if (!modal) return;
  modal.style.display = 'flex';
  const errBox = q('#wbModalError');
  errBox.style.display = 'none';
  errBox.textContent = '';

  const isCloud = state?.connection?.mode === 'excel';
  q('#sourceLocal').classList.toggle('active', Boolean(state) && !isCloud);
  q('#sourceCloud').classList.toggle('active', isCloud);
  const nameEl = q('#wbModalFileName');
  const workbook = state?.workbook || '';
  const fileName = workbook.split(/[\\\/]/).pop();
  let shown = fileName;
  if (isCloud) { try { shown = decodeURIComponent(fileName); } catch { /* keep the raw name */ } }
  nameEl.textContent = shown || 'לא נבחר קובץ';
  nameEl.title = workbook;
}

function closeWorkbookModal() {
  const modal = q('#wbModalBackdrop');
  if (modal) modal.style.display = 'none';
}

function adoptWorkbook(data) {
  cloudConnectPending = false;
  q('#btnConnectCloud').disabled = false;
  updateConnection(data.connection || {mode:'local', state:'ready'});
  closeDrawer();
  showDataError('');
  mutationEpoch++;
  selectedTasks.clear();
  settingsDirty = false;
  metricFilter = '';
  state = data;
  lastPayload = dataFingerprint(data);
  render();
}

/**
 * Event Listeners Initialization
 */
// ---------------------------------------------------------------- dates
function isoDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function todayIso() { return isoDate(new Date()); }
function isoPlus(days, from = new Date()) {
  const d = new Date(from.getFullYear(), from.getMonth(), from.getDate() + days);
  return isoDate(d);
}
const HEBREW_WEEKDAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];
function dueLabel(iso) {
  const d = parseIsoDate(iso);
  if (!d) return 'ללא תאריך';
  const today = parseIsoDate(todayIso());
  const days = Math.round((d - today) / 86400000);
  if (days === 0) return 'היום';
  if (days === 1) return 'מחר';
  if (days === -1) return 'אתמול';
  if (days < 0) return `לפני ${-days} ימים`;
  if (days < 7) return `יום ${HEBREW_WEEKDAYS[d.getDay()]}`;
  return fmtDate(iso);
}

// ---------------------------------------------------------------- my day
function myDayBuckets() {
  const user = getCurrentUser();
  const mine = openTasks(state?.tasks || []).filter(t => t.owner === user);
  const today = todayIso();
  const weekEnd = isoPlus(7);
  const byDue = (a, b) => (a.dueDate || '').localeCompare(b.dueDate || '') || taskSort(a, b);
  return {
    mine,
    overdue: mine.filter(t => t.dueDate && t.dueDate < today).sort(byDue),
    today: mine.filter(t => t.dueDate === today).sort(taskSort),
    week: mine.filter(t => t.dueDate && t.dueDate > today && t.dueDate <= weekEnd).sort(byDue),
    later: mine.filter(t => !t.dueDate || t.dueDate > weekEnd).sort(byDue),
  };
}

function myDayRow(t) {
  const due = dueInfo(t);
  const dueHtml = due.tone === 'overdue' ? `<span class="is-late">${escapeHtml(due.text)}</span>` : escapeHtml(dueLabel(t.dueDate));
  const meta = [t.category ? escapeHtml(t.category) : '', dueHtml, t.status ? escapeHtml(t.status) : ''].filter(Boolean).join(' · ');
  return `
    <div class="myday-row ${t.blocked === YES ? 'is-blocked' : ''}" data-id="${escapeAttr(t.id)}">
      <button type="button" class="myday-check" data-id="${escapeAttr(t.id)}" title="סימון כבוצע" aria-label="סימון כבוצע: ${escapeAttr(t.title)}"></button>
      <div class="myday-main">
        <div class="myday-title">${escapeHtml(t.title)}</div>
        <div class="myday-meta">${meta}${t.blocked === YES ? ' · <span class="myday-blocked">חסומה</span>' : ''}</div>
      </div>
      <span class="priority-badge ${priorityClass[t.priority] || 'p2'}">${escapeHtml(t.priority)}</span>
    </div>`;
}

function renderMyDay() {
  const content = q('#mydayContent');
  if (!content || !state) return;
  const user = getCurrentUser();
  const b = myDayBuckets();
  const now = new Date();
  const dateText = now.toLocaleDateString('he-IL', { weekday: 'long', day: 'numeric', month: 'long' });
  const parts = [];
  if (b.overdue.length) parts.push(`${b.overdue.length} באיחור`);
  if (b.today.length) parts.push(`${b.today.length} להיום`);
  if (b.week.length) parts.push(`${b.week.length} השבוע`);
  q('#mydaySummary').textContent = `${dateText} · ${parts.length ? parts.join(', ') : 'אין משימות דחופות'}`;
  const remindersOn = localStorage.getItem('teamflow_reminders') === 'on';
  q('#btnRemindersLabel').textContent = remindersOn ? 'תזכורות פעילות' : 'הפעלת תזכורות';
  q('#btnReminders').classList.toggle('is-on', remindersOn);

  if (!b.mine.length) {
    const known = (state.tasks || []).some(t => t.owner === user);
    content.innerHTML = `
      <div class="myday-empty">
        <svg class="myday-empty-art" viewBox="0 0 120 120" aria-hidden="true">
          <path class="art-rays" d="M90 8v6M90 50v6M66 32h6M108 32h6M73 15l4 4M103 45l4 4M107 15l-4 4M77 45l-4 4"/>
          <circle class="art-sun" cx="90" cy="32" r="11"/>
          <rect class="art-card" x="14" y="42" width="70" height="62" rx="13"/>
          <path class="art-check" d="M27 62l5 5 9-10"/>
          <path class="art-line" d="M50 63h22"/>
          <path class="art-check" d="M27 84l5 5 9-10"/>
          <path class="art-line" d="M50 85h15"/>
        </svg>
        <h3>${known ? 'הכול הושלם להיום' : `אין משימות על שם ${escapeHtml(user)}`}</h3>
        <p>${known ? 'אין לך משימות פתוחות. אפשר לקחת משימה מהלוח או ליצור חדשה.' : 'אם השם שלך מופיע בקובץ אחרת, בחר אותו מתפריט המשתמש למעלה.'}</p>
        <div class="myday-empty-actions">
          ${known ? '<button type="button" class="btn btn-primary" data-empty-action="new">משימה חדשה</button>'
                  : '<button type="button" class="btn btn-primary" data-empty-action="user">בחירת משתמש</button>'}
          <button type="button" class="btn btn-outline" data-empty-action="board">ללוח המשימות</button>
        </div>
      </div>`;
    content.querySelectorAll('[data-empty-action]').forEach(btn => btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const action = btn.dataset.emptyAction;
      if (action === 'new') openNewTask();
      else if (action === 'board') setView('board');
      else {
        q('#userDropdownMenu').style.display = 'block';
        q('#userProfileSelector').classList.add('open');
      }
    }));
    return;
  }
  const section = (title, list, tone) => list.length ? `
    <section class="myday-section">
      <h2 class="myday-heading" data-tone="${tone}">${title}<span>${list.length}</span></h2>
      <div class="myday-list">${list.map(myDayRow).join('')}</div>
    </section>` : '';
  content.innerHTML = section('באיחור', b.overdue, 'critical') + section('להיום', b.today, 'accent')
    + section('בשבוע הקרוב', b.week, '') + section('בהמשך או ללא תאריך', b.later, '');
  content.querySelectorAll('.myday-row').forEach(row => row.addEventListener('click', () => openTask(row.dataset.id)));
  content.querySelectorAll('.myday-check').forEach(btn => btn.addEventListener('click', (e) => {
    e.stopPropagation();
    btn.classList.add('checked');
    setTimeout(() => btn.closest('.myday-row')?.classList.add('completing'), 160);
    setTimeout(() => markDone(btn.dataset.id), 380);
  }));
}

// ---------------------------------------------------------------- reminders
function remindersAllowed() {
  return localStorage.getItem('teamflow_reminders') === 'on' && 'Notification' in window && Notification.permission === 'granted';
}

async function toggleReminders() {
  if (!('Notification' in window)) { toast('הדפדפן אינו תומך בהתראות'); return; }
  if (localStorage.getItem('teamflow_reminders') === 'on') {
    localStorage.setItem('teamflow_reminders', 'off');
    toast('התזכורות כובו');
    renderMyDay();
    return;
  }
  const permission = Notification.permission === 'granted' ? 'granted' : await Notification.requestPermission();
  if (permission !== 'granted') { toast('יש לאשר התראות בדפדפן כדי לקבל תזכורות'); return; }
  localStorage.setItem('teamflow_reminders', 'on');
  toast('תזכורות פעילות. התראה תופיע כשמשימה שלך מגיעה ליעד או באיחור.');
  renderMyDay();
  checkReminders();
}

// One Windows notification per task per day, while TeamFlow is open.
function checkReminders() {
  if (!state || !remindersAllowed()) return;
  const today = todayIso();
  const key = 'teamflow_reminded_' + today;
  let shown;
  try { shown = new Set(JSON.parse(localStorage.getItem(key) || '[]')); } catch { shown = new Set(); }
  const b = myDayBuckets();
  const fresh = [...b.overdue, ...b.today].filter(t => !shown.has(t.id));
  if (!fresh.length) return;
  for (const k of Object.keys(localStorage)) {
    if (k.startsWith('teamflow_reminded_') && k !== key) localStorage.removeItem(k);
  }
  const notify = (title, body, taskId) => {
    const n = new Notification(title, { body, icon: 'assets/logo.svg', tag: taskId || 'teamflow-summary', lang: 'he', dir: 'rtl' });
    n.onclick = () => { window.focus(); setView('myday'); if (taskId) openTask(taskId); n.close(); };
  };
  if (fresh.length <= 3) {
    for (const t of fresh) notify(t.title, `${t.dueDate < today ? 'באיחור' : 'יעד היום'} · ${t.priority}`, t.id);
  } else {
    notify(`${fresh.length} משימות שלך דורשות טיפול`, fresh.slice(0, 3).map(t => t.title).join(' · '), '');
  }
  fresh.forEach(t => shown.add(t.id));
  localStorage.setItem(key, JSON.stringify([...shown]));
}

// ---------------------------------------------------------------- calendar
function renderCalendar() {
  const grid = q('#calendarGrid');
  if (!grid || !state) return;
  if (!calendarMonth) {
    const now = new Date();
    calendarMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  }
  q('#calendarTitle').textContent = calendarMonth.toLocaleDateString('he-IL', { month: 'long', year: 'numeric' });
  q('#calendarWeekdays').innerHTML = HEBREW_WEEKDAYS.map(d => `<div>${d}</div>`).join('');
  const tasks = visibleTasks();
  const byDay = Object.create(null);
  for (const t of tasks) {
    if (!t.dueDate) continue;
    (byDay[t.dueDate] = byDay[t.dueDate] || []).push(t);
  }
  const undated = tasks.filter(t => !t.dueDate && t.status !== DONE && t.status !== CANCELLED).length;
  q('#calendarNote').textContent = 'גוררים משימה ליום אחר כדי לשנות את תאריך היעד.'
    + (undated ? ` ${undated} משימות פתוחות ללא תאריך יעד אינן מוצגות.` : '');
  const start = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth(), 1 - calendarMonth.getDay());
  const monthEnd = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() + 1, 0);
  const weeks = Math.ceil((monthEnd.getDate() + calendarMonth.getDay()) / 7);
  const today = todayIso();
  const cells = [];
  for (let i = 0; i < weeks * 7; i++) {
    const d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
    const iso = isoDate(d);
    const items = (byDay[iso] || []).sort(taskSort);
    const classes = ['cal-cell', d.getMonth() !== calendarMonth.getMonth() ? 'other-month' : '', iso === today ? 'is-today' : '', d.getDay() === 6 ? 'is-weekend' : ''].filter(Boolean).join(' ');
    cells.push(`
      <div class="${classes}" data-date="${iso}">
        <div class="cal-date">${d.getDate()}</div>
        <div class="cal-items">${items.map(t => {
          const tone = t.status === DONE ? 'is-done' : t.health === HEALTH_OVERDUE ? 'is-overdue' : t.health === HEALTH_BLOCKED ? 'is-blocked' : '';
          return `<button type="button" class="cal-chip ${tone}" draggable="true" data-id="${escapeAttr(t.id)}" title="${escapeAttr(`${t.title} · ${t.owner || 'ללא שיוך'} · ${t.priority} · ${t.status}`)}">${avatarHtml(t.owner, 'avatar-xs', false)}<span class="cal-chip-text">${escapeHtml(t.title)}</span></button>`;
        }).join('')}</div>
      </div>`);
  }
  grid.innerHTML = cells.join('');
  grid.querySelectorAll('.cal-chip').forEach(chip => {
    chip.addEventListener('click', () => openTask(chip.dataset.id));
    chip.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', chip.dataset.id); chip.classList.add('dragging'); });
    chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
  });
  grid.querySelectorAll('.cal-cell').forEach(cell => {
    cell.addEventListener('dragover', (e) => { e.preventDefault(); cell.classList.add('drop-target'); });
    cell.addEventListener('dragleave', () => cell.classList.remove('drop-target'));
    cell.addEventListener('drop', (e) => {
      e.preventDefault();
      cell.classList.remove('drop-target');
      moveDueDate(e.dataTransfer.getData('text/plain'), cell.dataset.date);
    });
  });
}

async function moveDueDate(id, dueDate) {
  const t = (state.tasks || []).find(task => task.id === id);
  if (!t || t.dueDate === dueDate) return;
  try {
    await api('/api/task', { method: 'POST', body: {
      id: t.id, title: t.title, owner: t.owner, priority: t.priority, status: t.status, category: t.category,
      dueDate, blocked: t.blocked, progress: t.progress, notes: t.notes, user: getCurrentUser() } });
    toast(`תאריך היעד עודכן ל-${dueLabel(dueDate)}`);
    await load();
  } catch (err) {
    toast(err.message);
  }
}

function shiftCalendar(months) {
  if (months === 0) {
    const now = new Date();
    calendarMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  } else {
    calendarMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() + months, 1);
  }
  renderCalendar();
}

// ---------------------------------------------------------------- weekly report
function reportRange(kind) {
  const now = new Date();
  const sunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
  if (kind === 'last-7') return { from: isoPlus(-6), to: todayIso() };
  if (kind === 'last-week') return { from: isoPlus(-7, sunday), to: isoPlus(-1, sunday) };
  if (kind === 'this-month') return { from: isoDate(new Date(now.getFullYear(), now.getMonth(), 1)), to: todayIso() };
  return { from: isoDate(sunday), to: todayIso() };
}

// Same rules as build_report in app.py, so the screen and the Excel file agree.
function buildReport(range) {
  const tasks = state?.tasks || [];
  const today = todayIso();
  const weekEnd = isoPlus(7);
  const inRange = v => Boolean(v) && range.from <= String(v).slice(0, 10) && String(v).slice(0, 10) <= range.to;
  const open = tasks.filter(t => t.status !== DONE && t.status !== CANCELLED);
  const completed = tasks.filter(t => t.status === DONE && inRange(t.completedDate));
  const created = tasks.filter(t => inRange(t.createdAt));
  const overdue = open.filter(t => t.health === HEALTH_OVERDUE).sort((a, b) => (a.dueDate || '').localeCompare(b.dueDate || ''));
  const upcoming = open.filter(t => t.dueDate && t.dueDate > today && t.dueDate <= weekEnd).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const owners = [...new Set([...(state?.people || []), ...tasks.map(t => t.owner || '')])];
  const people = owners.map(owner => {
    const mine = tasks.filter(t => (t.owner || '') === owner);
    return { owner: owner || 'ללא שיוך', open: mine.filter(t => open.includes(t)).length,
      done: mine.filter(t => completed.includes(t)).length, overdue: mine.filter(t => overdue.includes(t)).length };
  }).filter(p => p.open || p.done);
  return { ...range, completed, created, overdue, upcoming, people, openCount: open.length };
}

function reportHtml(r) {
  const fmt = iso => (parseIsoDate(iso) ? fmtDate(iso) : '');
  const list = (title, items, empty, extra) => `
    <section class="report-section">
      <h4>${title} <span>${items.length}</span></h4>
      ${items.length ? `<ul class="report-list">${items.slice(0, 40).map(t => `
        <li><span class="priority-badge sm ${priorityClass[t.priority] || 'p2'}">${escapeHtml(t.priority)}</span>
        <span class="report-task">${escapeHtml(t.title)}</span>
        <span class="report-meta">${escapeHtml(t.owner || 'ללא שיוך')} · ${escapeHtml(extra(t))}</span></li>`).join('')}</ul>` : `<p class="muted-empty">${empty}</p>`}
    </section>`;
  return `
    <div class="report-head">
      <h2>דוח צוות</h2>
      <p>${fmt(r.from)} – ${fmt(r.to)} · הופק ע"י ${escapeHtml(getCurrentUser())}</p>
    </div>
    <div class="report-figures">
      <div><strong>${r.completed.length}</strong><span>הושלמו בתקופה</span></div>
      <div><strong>${r.created.length}</strong><span>נוצרו בתקופה</span></div>
      <div><strong>${r.openCount}</strong><span>פתוחות כעת</span></div>
      <div><strong>${r.overdue.length}</strong><span>באיחור כעת</span></div>
    </div>
    ${r.people.length ? `<section class="report-section"><h4>לפי אחראי</h4>
      <table class="report-table"><thead><tr><th>אחראי</th><th>פתוחות</th><th>הושלמו</th><th>באיחור</th></tr></thead>
      <tbody>${r.people.map(p => `<tr><td>${escapeHtml(p.owner)}</td><td>${p.open}</td><td>${p.done}</td><td>${p.overdue}</td></tr>`).join('')}</tbody></table></section>` : ''}
    ${list('הושלמו', r.completed, 'לא הושלמו משימות בתקופה זו.', t => `הושלמה ${fmt(t.completedDate)}`)}
    ${list('באיחור', r.overdue, 'אין משימות באיחור.', t => `יעד ${fmt(t.dueDate)}`)}
    ${list('יעד בשבוע הקרוב', r.upcoming, 'אין יעדים בשבוע הקרוב.', t => `יעד ${fmt(t.dueDate)}`)}`;
}

function renderReport() {
  q('#reportBody').innerHTML = reportHtml(buildReport(reportRange(q('#reportPeriod').value)));
}

function openReport() {
  if (!state) return;
  renderReport();
  q('#reportBackdrop').style.display = 'flex';
}

function closeReport() {
  q('#reportBackdrop').style.display = 'none';
}

function printReport() {
  const area = q('#printArea');
  area.innerHTML = q('#reportBody').innerHTML;
  window.print();
}

async function downloadReport() {
  const range = reportRange(q('#reportPeriod').value);
  const button = q('#btnReportExcel');
  button.disabled = true;
  try {
    const res = await fetch(`/api/report.xlsx?from=${range.from}&to=${range.to}`);
    if (!res.ok) {
      let message = 'לא ניתן להפיק את הקובץ';
      try { message = (await res.json()).error || message; } catch { /* not JSON */ }
      throw new Error(message);
    }
    const url = URL.createObjectURL(await res.blob());
    const link = document.createElement('a');
    link.href = url;
    link.download = `TeamFlow-report-${range.to}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    toast('הדוח הורד כקובץ Excel');
  } catch (err) {
    toast(err.message);
  } finally {
    button.disabled = false;
  }
}

function toggleTheme() {
  const root = document.documentElement;
  const dark = root.dataset.theme === 'dark';
  root.dataset.theme = dark ? 'light' : 'dark';
  localStorage.setItem('teamflow_theme', root.dataset.theme);
}

function initEvents() {
  q('#btnThemeToggle').addEventListener('click', toggleTheme);
  q('#btnOpenPalette').addEventListener('click', (e) => { e.stopPropagation(); openPalette(); });
  q('#formTaskDueDateText').addEventListener('change', syncDueFromText);
  q('#formTaskDueDateText').addEventListener('input', () => q('#formTaskDueDateText').classList.remove('is-invalid'));
  q('#formTaskDueDate').addEventListener('change', () => setDueField(q('#formTaskDueDate').value));
  q('#btnPickDueDate').addEventListener('click', () => {
    const native = q('#formTaskDueDate');
    try { native.showPicker(); } catch { native.focus(); native.click(); }
  });
  qa('.drawer-tab').forEach(tab => tab.addEventListener('click', () => {
    setDrawerTab(tab.dataset.tab);
    if (tab.dataset.tab === 'checklist') q('#newSubtaskInput').focus();
    if (tab.dataset.tab === 'activity') q('#newCommentText').focus();
  }));
  q('#btnReminders').addEventListener('click', toggleReminders);
  q('#calPrev').addEventListener('click', () => shiftCalendar(-1));
  q('#calNext').addEventListener('click', () => shiftCalendar(1));
  q('#calToday').addEventListener('click', () => shiftCalendar(0));
  q('#btnWeeklyReport').addEventListener('click', openReport);
  q('#reportCloseBtn').addEventListener('click', closeReport);
  q('#reportBackdrop').addEventListener('click', (e) => { if (e.target === q('#reportBackdrop')) closeReport(); });
  q('#reportPeriod').addEventListener('change', renderReport);
  q('#btnReportPrint').addEventListener('click', printReport);
  q('#btnReportExcel').addEventListener('click', downloadReport);
  q('#dataErrorChooseWorkbook').addEventListener('click', openWorkbookModal);
  // Navigation view switcher buttons
  qa('.nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      metricFilter = '';
      setView(btn.dataset.view);
    });
  });

  // Topbar Actions
  q('#btnNewTask').addEventListener('click', openNewTask);
  q('#btnRefresh').addEventListener('click', () => refreshFromDisk(false));

  // Quick Share Actions
  const btnShareTop = q('#btnShareDigest');
  if (btnShareTop) btnShareTop.addEventListener('click', openShareModal);

  // Open active workbook directly in Excel
  const btnOpenExcel = q('#btnOpenExcel');
  if (btnOpenExcel) {
    btnOpenExcel.addEventListener('click', async () => {
      try {
        await api('/api/open-excel', { method: 'POST' });
        toast('קובץ האקסל נפתח ב-Excel ↗');
      } catch (err) {
        toast(err.message || 'שגיאה בפתיחת האקסל');
      }
    });
  }

  // Timeline Grouping Buttons
  qa('.time-group-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      qa('.time-group-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      timelineGroupBy = btn.dataset.group || 'owner';
      renderTimeline();
    });
  });

  // Share Modal Controls
  const shareClose = q('#shareCloseBtn');
  if (shareClose) shareClose.addEventListener('click', closeShareModal);
  const shareBackdrop = q('#shareModalBackdrop');
  if (shareBackdrop) {
    shareBackdrop.addEventListener('click', (e) => {
      if (e.target === shareBackdrop) closeShareModal();
    });
  }
  const btnCopyShare = q('#btnCopyShareText');
  if (btnCopyShare) {
    btnCopyShare.addEventListener('click', async () => {
      const txt = q('#shareDigestText');
      if (txt) {
        try {
          await navigator.clipboard.writeText(txt.value);
          toast('הסיכום הועתק ללוח! 📋');
        } catch {
          txt.select();
          document.execCommand('copy');
          toast('הסיכום הועתק ללוח!');
        }
      }
    });
  }

  // Subtasks Checklist input and add button
  const btnAddSubtask = q('#btnAddSubtask');
  const newSubtaskInput = q('#newSubtaskInput');
  const handleAddSubtask = () => {
    if (!newSubtaskInput) return;
    const val = (newSubtaskInput.value || '').trim();
    if (!val) return;
    currentDrawerSubtasks.push({ done: false, text: val });
    updateChecklistProgress();
    newSubtaskInput.value = '';
    renderDrawerSubtasks();
  };
  if (btnAddSubtask) btnAddSubtask.addEventListener('click', handleAddSubtask);
  if (newSubtaskInput) {
    newSubtaskInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleAddSubtask();
      }
    });
  }

  // Links and comments in the task drawer
  q('#btnAddLink').addEventListener('click', addDrawerLink);
  for (const selector of ['#newLinkUrl', '#newLinkTitle']) {
    q(selector).addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); addDrawerLink(); }
    });
  }
  q('#btnAddComment').addEventListener('click', addComment);
  q('#newCommentText').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); addComment(); }
  });

  // The notes box holds only the notes; the checklist and links are joined to them when saving.
  const formNotes = q('#formTaskNotes');
  if (formNotes) {
    formNotes.addEventListener('input', () => {
      currentDrawerCleanNotes = formNotes.value;
    });
  }

  // User Profile Selector
  const userSelector = q('#userProfileSelector');
  const userMenu = q('#userDropdownMenu');
  if (userSelector && userMenu) {
    userSelector.addEventListener('click', (e) => {
      e.stopPropagation();
      const isVisible = userMenu.style.display === 'block';
      userMenu.style.display = isVisible ? 'none' : 'block';
      userSelector.classList.toggle('open', !isVisible);
    });
  }

  const btnSetCustom = q('#btnSetCustomUser');
  const customInput = q('#customUserNameInput');
  const handleCustomUserSubmit = () => {
    if (!customInput) return;
    const val = (customInput.value || '').trim();
    if (val) {
      setCurrentUser(val);
      customInput.value = '';
      if (userMenu) userMenu.style.display = 'none';
      if (userSelector) userSelector.classList.remove('open');
    }
  };
  if (btnSetCustom) btnSetCustom.addEventListener('click', handleCustomUserSubmit);
  if (customInput) {
    customInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleCustomUserSubmit();
      }
    });
  }

  // Click outside to close user dropdown
  document.addEventListener('click', (e) => {
    if (userMenu && !userMenu.contains(e.target) && userSelector && !userSelector.contains(e.target)) {
      userMenu.style.display = 'none';
      userSelector.classList.remove('open');
    }
  });

  const btnExit = q('#btnExitApp');
  if (btnExit) {
    btnExit.addEventListener('click', () => {
      if (confirm('האם אתה בטוח שברצונך לסגור את האפליקציה ולכבות את השרת?')) {
        if (navigator.sendBeacon) {
          navigator.sendBeacon('/api/exit');
        } else {
          api('/api/exit', { method: 'POST' }).catch(() => {});
        }
        document.body.innerHTML = `
          <div class="exit-screen">
            <div class="exit-card">
              <div class="exit-check">✓</div>
              <h2>האפליקציה נסגרה</h2>
              <p>השרת כובה והשינויים שמורים בקובץ. אפשר לסגור את החלון.</p>
              <button class="btn btn-primary" onclick="window.close()">סגירת החלון</button>
            </div>
          </div>
        `;
        try { window.close(); } catch (e) {}
      }
    });
  }

  const searchInput = q('#globalSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      render();
    });
  }

  const clearSearchBtn = q('#btnClearSearch');
  if (clearSearchBtn) {
    clearSearchBtn.addEventListener('click', () => {
      if (searchInput) {
        searchInput.value = '';
        render();
        searchInput.focus();
      }
    });
  }

  // Keyboard shortcut Ctrl+K, Escape, ?, numbers, N, R, P
  document.addEventListener('keydown', (e) => {
    // If command palette is open
    const paletteBackdrop = q('#paletteBackdrop');
    const isPaletteOpen = paletteBackdrop && paletteBackdrop.style.display === 'flex';

    if (isPaletteOpen) {
      if (e.key === 'Escape') {
        e.preventDefault();
        closePalette();
        return;
      }
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        if (paletteItems.length > 0) {
          paletteSelectedIndex = (paletteSelectedIndex + 1) % paletteItems.length;
          renderPaletteResults(q('#paletteInput').value);
          const sel = q(`.palette-item[data-index="${paletteSelectedIndex}"]`);
          if (sel) sel.scrollIntoView({ block: 'nearest' });
        }
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        if (paletteItems.length > 0) {
          paletteSelectedIndex = (paletteSelectedIndex - 1 + paletteItems.length) % paletteItems.length;
          renderPaletteResults(q('#paletteInput').value);
          const sel = q(`.palette-item[data-index="${paletteSelectedIndex}"]`);
          if (sel) sel.scrollIntoView({ block: 'nearest' });
        }
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        if (paletteItems[paletteSelectedIndex]) {
          closePalette();
          paletteItems[paletteSelectedIndex].run();
        }
        return;
      }
      return;
    }

    // Ctrl+K opens Command Palette
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      openPalette();
      return;
    }

    // Don't trigger single-key shortcuts if typing in form inputs
    const tag = (e.target && e.target.tagName ? e.target.tagName.toLowerCase() : '');
    const isTyping = tag === 'input' || tag === 'textarea' || tag === 'select' || (e.target && e.target.isContentEditable);
    if (isTyping) {
      if (e.key === 'Escape') {
        if (editing) closeDrawer();
        else if (e.target === searchInput) {
          searchInput.value = '';
          render();
          searchInput.blur();
        }
      }
      return;
    }

    if (e.key === '?') {
      e.preventDefault();
      openShortcuts();
      return;
    }
    if (e.key === 'Escape') {
      if (editing) closeDrawer();
      closeShortcuts();
      closePalette();
      closeShareModal();
      closeReport();
      return;
    }
    if (e.key.toLowerCase() === 'n') {
      e.preventDefault();
      openNewTask();
      return;
    }
    if (e.key.toLowerCase() === 'r') {
      e.preventDefault();
      refreshFromDisk(false);
      return;
    }
    const order = ['myday', 'dashboard', 'timeline', 'calendar', 'board', 'tasks', 'settings'];
    if (/^[1-7]$/.test(e.key)) { metricFilter = ''; setView(order[Number(e.key) - 1]); }
  });

  // Palette live search
  const pInput = q('#paletteInput');
  if (pInput) {
    pInput.addEventListener('input', () => {
      paletteSelectedIndex = 0;
      renderPaletteResults(pInput.value);
    });
  }

  // Palette backdrop close on outside click
  const pBackdrop = q('#paletteBackdrop');
  if (pBackdrop) {
    pBackdrop.addEventListener('click', (e) => {
      if (e.target === pBackdrop) closePalette();
    });
  }

  // Shortcuts modal close
  const sCloseBtn = q('#shortcutsCloseBtn');
  if (sCloseBtn) sCloseBtn.addEventListener('click', closeShortcuts);
  const sBackdrop = q('#shortcutsBackdrop');
  if (sBackdrop) {
    sBackdrop.addEventListener('click', (e) => {
      if (e.target === sBackdrop) closeShortcuts();
    });
  }

  // Sidebar Quick Filter Shortcuts
  qa('.sidebar-filter-item').forEach(btn => {
    btn.addEventListener('click', () => {
      const filterType = btn.dataset.filter;
      qa('.sidebar-filter-item').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      if (filterType === 'my-tasks') {
        applyMetricFilter('my-tasks');
      } else if (filterType === 'overdue') {
        applyMetricFilter('overdue');
      } else if (filterType === 'blocked') {
        applyMetricFilter('blocked');
      } else if (filterType === 'p0') {
        metricFilter = '';
        setView('tasks');
        const fp = q('#filterPriority');
        if (fp) fp.value = 'P0';
        renderTable();
      } else if (filterType === 'done') {
        applyMetricFilter('done');
      }
    });
  });

  // Workbook Modal & Management
  const btnSelectWb = q('#btnSelectWorkbook');
  if (btnSelectWb) btnSelectWb.addEventListener('click', openWorkbookModal);

  const btnCloseWb = q('#wbModalCloseBtn');
  if (btnCloseWb) btnCloseWb.addEventListener('click', closeWorkbookModal);

  const wbBackdrop = q('#wbModalBackdrop');
  if (wbBackdrop) {
    wbBackdrop.addEventListener('click', (e) => {
      if (e.target === wbBackdrop) closeWorkbookModal();
    });
  }

  const wbErrBox = q('#wbModalError');

  // Windows Native Picker trigger
  const btnWinPicker = q('#btnTriggerWinPicker');
  if (btnWinPicker) {
    btnWinPicker.addEventListener('click', async () => {
      btnWinPicker.disabled = true;
      btnWinPicker.textContent = 'ממתין לבחירה... ⏳';
      if (wbErrBox) wbErrBox.style.display = 'none';

      try {
        const res = await api('/api/browse-workbook', { method: 'POST' });
        if (res && res.cancelled) {
          // Dialog was closed by user without selection
        } else if (res && res.state === 'connecting') {
          cloudConnectPending = true;
          q('#btnConnectCloud').disabled = true;
          q('#cloudConnectFeedback').textContent = 'מצב בדיקה: Excel פותח את הקובץ המקומי…';
          updateConnection(res);
        } else if (res && res.data) {
          adoptWorkbook(res.data);
          closeWorkbookModal();
          toast('קובץ האקסל הוחלף ונטען בהצלחה! ✓');
        }
      } catch (err) {
        if (wbErrBox) {
          wbErrBox.textContent = err.message || 'שגיאה בפתיחת סייר הקבצים';
          wbErrBox.style.display = 'block';
        }
      } finally {
        btnWinPicker.disabled = false;
        btnWinPicker.textContent = 'בחירת קובץ…';
      }
    });
  }

  // Settings dirty tracking
  ['#settingsPeople', '#settingsCategories', '#settingsStatuses', '#settingsPriorities'].forEach(sel => {
    const el = q(sel);
    if (el) {
      el.addEventListener('input', () => {
        if (!settingsDirty && state) {
          settingsRevision = state.revision;
          settingsWorkbook = state.workbook;
          settingsListsRevision = state.listsRevision || '';
        }
        settingsDirty = true;
      });
    }
  });

  // Table Filters
  ['#filterOwner', '#filterStatus', '#filterPriority'].forEach(sel => {
    const el = q(sel);
    if (el) el.addEventListener('change', () => {
      metricFilter = '';
      renderTable();
    });
  });

  q('#btnClearFilters').addEventListener('click', () => {
    metricFilter = '';
    tableSortKey = '';
    tableSortAsc = true;
    const fo = q('#filterOwner');
    const fs = q('#filterStatus');
    const fp = q('#filterPriority');
    const si = q('#globalSearchInput');
    if (fo) fo.value = ALL;
    if (fs) fs.value = ALL;
    if (fp) fp.value = ALL;
    if (si) si.value = '';
    renderTable();
  });

  // Table Column Header Click to Sort
  qa('#tasksTable thead th.th-sortable').forEach(th => {
    th.addEventListener('click', () => {
      const sortKey = th.dataset.sort;
      if (tableSortKey === sortKey) {
        tableSortAsc = !tableSortAsc;
      } else {
        tableSortKey = sortKey;
        tableSortAsc = true;
      }
      renderTable();
    });
  });

  // Board Sort Select
  const boardSortSelect = q('#boardSortSelect');
  if (boardSortSelect) {
    boardSortSelect.addEventListener('change', () => {
      renderBoard();
    });
  }

  // Master Checkbox
  const masterCb = q('#thMasterCheckbox');
  if (masterCb) {
    masterCb.addEventListener('click', () => {
      const tasks = filteredTableTasks();
      const shouldCheck = masterCb.checked;
      for (const t of tasks) {
        if (shouldCheck) selectedTasks.add(t.id);
        else selectedTasks.delete(t.id);
      }
      renderTable();
      renderBulkBar();
    });
  }

  // Bulk Actions
  q('#bulkMarkDone').addEventListener('click', () => bulkAction('done'));
  q('#bulkDelete').addEventListener('click', () => bulkAction('delete'));
  q('#bulkClearSelection').addEventListener('click', clearSelection);

  // Bulk dropdown toggles
  ['#bulkStatusBtn', '#bulkOwnerBtn', '#bulkPriorityBtn'].forEach(btnSel => {
    const btn = q(btnSel);
    if (!btn) return;
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const menu = btn.nextElementSibling;
      const wasOpen = menu.classList.contains('show');
      qa('.bulk-menu').forEach(m => m.classList.remove('show'));
      if (!wasOpen) menu.classList.add('show');
    });
  });

  document.addEventListener('click', () => {
    qa('.bulk-menu').forEach(m => m.classList.remove('show'));
  });

  // Drawer Form
  q('#taskForm').addEventListener('submit', saveTask);
  q('#drawerCloseBtn').addEventListener('click', closeDrawer);
  q('#formBtnCancel').addEventListener('click', closeDrawer);
  q('#drawerBackdrop').addEventListener('click', closeDrawer);
  q('#formBtnDelete').addEventListener('click', deleteCurrentTask);

  // Drawer progress slider live value
  const progInput = q('#formTaskProgress');
  const progText = q('#formProgressValText');
  if (progInput && progText) {
    progInput.addEventListener('input', () => {
      progText.textContent = `${progInput.value}%`;
    });
  }

  // Settings Form
  q('#settingsForm').addEventListener('submit', (e) => {
    e.preventDefault();
    saveLists();
  });

  // Heartbeat ping every 5 seconds
  setInterval(() => {
    api('/api/ping').catch(() => {});
  }, 5000);

  // Send ping and refresh when tab becomes visible or gains focus
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      api('/api/ping').catch(() => {});
      refreshFromDisk(true);
    }
  });
  window.addEventListener('focus', () => {
    api('/api/ping').catch(() => {});
    refreshFromDisk(true);
  });

  // Refresh the main view without overwriting an open task draft.
  setInterval(() => {
    const interval = state?.connection?.mode === 'excel' ? 2000 : 7000;
    if (Date.now() - lastAutomaticRefresh >= interval) {
      lastAutomaticRefresh = Date.now();
      refreshFromDisk(true);
    }
  }, 1000);
}


function dataFingerprint(data) {
  const {connection, ...content} = data;
  return JSON.stringify(content);
}

function updateConnection(info) {
  if (!info) return;
  const el = q('#connectionIndicator');
  if (!el) return;
  const isCloud = info.mode === 'excel';
  el.dataset.state = info.state || 'ready';
  q('#connectionLabel').textContent = info.state === 'connecting' ? 'מתחבר דרך Excel…'
    : info.state === 'error' ? 'החיבור דורש טיפול'
    : info.busy ? 'Excel עסוק · ממתין'
    : isCloud ? 'מחובר ל-Excel · AutoSave פעיל' : 'עבודה מקומית';
  q('#connectionDetail').textContent = isCloud
    ? (info.lastReadAt ? 'קריאה אחרונה מ-Excel: ' + new Date(info.lastReadAt).toLocaleTimeString('he-IL') : info.message || '')
    : 'הנתונים נשמרים בקובץ שנבחר במחשב';
  q('#btnReconnectCloud').hidden = !isCloud || info.state === 'connecting';
  if (info.url) {
    lastConnectionUrl = info.url;
    if (document.activeElement !== q('#cloudWorkbookUrl')) q('#cloudWorkbookUrl').value = info.url;
  }
}

async function pollConnection() {
  if (connectionPollPending) return;
  connectionPollPending = true;
  try {
    const info = await api('/api/connection');
    updateConnection(info);
    const feedback = q('#cloudConnectFeedback');
    if (info.attempt) {
      // The new connection failed; the previous data source is still in use.
      feedback.textContent = info.attempt.message || '';
      feedback.dataset.state = 'error';
      cloudConnectPending = false;
      q('#btnConnectCloud').disabled = false;
    } else if (info.mode === 'excel') {
      feedback.textContent = info.message || '';
      feedback.dataset.state = info.state;
      if (info.state === 'error') {
        cloudConnectPending = false;
        q('#btnConnectCloud').disabled = false;
      } else if (info.state === 'ready' && cloudConnectPending) {
        cloudConnectPending = false;
        q('#btnConnectCloud').disabled = false;
        await load();
        feedback.textContent = 'החיבור פעיל. אפשר לסגור חלון זה ולערוך משימה קיימת.';
      }
    }
  } catch (err) {
    q('#connectionIndicator').dataset.state = 'error';
    q('#connectionLabel').textContent = 'לא ניתן לקרוא את מצב החיבור';
  } finally {
    connectionPollPending = false;
  }
}

async function connectCloud() {
  if (editing || settingsDirty) { toast('יש לסיים או לבטל את העריכה הנוכחית לפני החלפת חיבור'); return; }
  const url = (q('#cloudWorkbookUrl').value || '').trim();
  const feedback = q('#cloudConnectFeedback');
  if (!url) { feedback.textContent = 'יש להדביק כתובת של קובץ Excel ב-SharePoint'; return; }
  q('#btnConnectCloud').disabled = true;
  cloudConnectPending = true;
  feedback.textContent = 'פותח את Excel. יש להשלים כניסה לחשבון העבודה אם תתבקש.';
  try {
    const info = await api('/api/cloud/connect', {method: 'POST', body: {url}});
    updateConnection(info);
  } catch (err) {
    cloudConnectPending = false;
    q('#btnConnectCloud').disabled = false;
    feedback.textContent = err.message;
    feedback.dataset.state = 'error';
  }
}

function initCloudConnection() {
  q('#btnConnectCloud').addEventListener('click', connectCloud);
  q('#cloudWorkbookUrl').addEventListener('keydown', event => {
    if (event.key === 'Enter') { event.preventDefault(); connectCloud(); }
  });
  q('#btnReconnectCloud').addEventListener('click', () => {
    q('#cloudWorkbookUrl').value = lastConnectionUrl;
    connectCloud();
  });
  pollConnection();
  setInterval(pollConnection, 2500);
}

// Initial Boot
document.addEventListener('DOMContentLoaded', async () => {
  initEvents();
  initCloudConnection();
  const validViews = ['myday', 'dashboard', 'timeline', 'calendar', 'board', 'tasks', 'settings'];
  const initialHash = window.location.hash.replace('#', '');
  if (validViews.includes(initialHash)) {
    currentView = initialHash;
  }
  const urlParams = new URLSearchParams(window.location.search);
  const viewParam = urlParams.get('view');
  if (viewParam && validViews.includes(viewParam)) {
    currentView = viewParam;
  }
  const qParam = urlParams.get('q');
  if (qParam) {
    const si = q('#globalSearchInput');
    if (si) si.value = qParam;
  }
  window.addEventListener('hashchange', () => {
    const h = window.location.hash.replace('#', '');
    if (validViews.includes(h)) {
      setView(h);
    }
  });
  // Show the running server's version, so the badge cannot fall behind a release.
  api('/api/health').then(info => {
    const badge = q('.brand-badge');
    if (badge && info.version) badge.textContent = info.version.split('.').slice(0, 2).join('.');
  }).catch(() => {});
  await load();
  const taskParam = urlParams.get('task');
  if (taskParam) {
    openTask(taskParam);
  }
  const shareParam = urlParams.get('share');
  if (shareParam) {
    openShareModal();
  }
  const wbModalParam = urlParams.get('wbmodal');
  if (wbModalParam) {
    openWorkbookModal();
  }
});
