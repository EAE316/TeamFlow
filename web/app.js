/**
 * TeamFlow / ניהול משימות צוות
 * Vanilla JavaScript Frontend Application
 */

// Top-level state and variables
let state = null;
let currentView = 'dashboard';
let lastPayload = '';
let editing = false;
let settingsDirty = false;
let metricFilter = '';
let draggedTaskId = '';
let tableSortKey = '';
let tableSortAsc = true;
const selectedTasks = new Set();
let timelineGroupBy = 'owner';
let currentDrawerSubtasks = [];
let currentDrawerCleanNotes = '';
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
const priorityColor = Object.assign(Object.create(null), { P0: '#d92d20', P1: '#f97316', P2: '#d98400', P3: '#1f6feb' });

const viewMeta = {
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
  if (avatarEl) avatarEl.textContent = initials(user);
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
    const isOffice = userName === state.systemUser ? ' (חשבון Office מזוהה)' : '';
    return `
      <div class="user-dropdown-item ${isAct}" data-user="${escapeAttr(userName)}">
        <span>👤 ${escapeHtml(userName)}${isOffice}</span>
        ${userName === current ? '<span>✓</span>' : ''}
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
  for (const selector of ['#btnNewTask', '#formBtnDelete', '#bulkMarkDone', '#bulkStatusBtn', '#bulkOwnerBtn', '#bulkPriorityBtn', '#bulkDelete', '#settingsForm button[type="submit"]']) {
    const button = q(selector);
    if (button) {
      button.disabled = cloudMode;
      button.title = cloudMode ? 'בחיבור הניסיוני ניתן לערוך משימות קיימות בלבד' : '';
    }
  }
  // Update workbook path display
  const pathEl = q('#workbookPathDisplay');
  if (pathEl) {
    pathEl.textContent = state.workbook || '';
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
    renderProgressSummary();
  } else if (currentView === 'timeline') {
    renderTimeline();
  } else if (currentView === 'board') {
    renderBoard();
  } else if (currentView === 'tasks') {
    renderTable();
    renderHistory();
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

  // Smart Insights generator
  const insightEl = q('#smartInsightText');
  if (!insightEl) return;

  const overdue = tasks.filter(t => t.health === HEALTH_OVERDUE);
  const blocked = tasks.filter(t => t.health === HEALTH_BLOCKED);
  const p0Open = overdue.filter(t => t.priority === 'P0');

  if (overdue.length > 0 && p0Open.length > 0) {
    insightEl.innerHTML = `⚠️ <strong>תשומת לב ניהולית:</strong> ישנן ${overdue.length} משימות שחרגו מהיעד, מתוכן <strong>${p0Open.length} דחופות (P0)</strong>. מומלץ לתעדף סגירתן בהקדם.`;
  } else if (overdue.length > 0) {
    insightEl.innerHTML = `⚠️ ישנן <strong>${overdue.length} משימות באיחור</strong>. לחץ על כרטיס "באיחור" בראש העמוד כדי לסנן ולטפל בהן.`;
  } else if (blocked.length > 0) {
    insightEl.innerHTML = `🛑 זוהו <strong>${blocked.length} משימות חסומות</strong> הממתינות להסרת מכשולים או אישורים.`;
  } else if (rate >= 80) {
    insightEl.innerHTML = `🌟 <strong>ביצועים מעולים!</strong> ${rate}% מכלל המשימות הושלמו. הצוות מתקרב ליעדי הפרויקט.`;
  } else if (activePeople > 0) {
    insightEl.innerHTML = `🚀 כל המשימות הפתוחות (${openT.length}) מתקדמות במסלול תקין על פני ${activePeople} אנשי צוות פעילים.`;
  } else {
    insightEl.textContent = 'כל המשימות הושלמו או שאין משימות פעילות כרגע.';
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
    { key: 'all', label: 'סה"כ', count: total, hint: 'כל המשימות בקובץ', color: '#2563eb' },
    { key: 'open', label: 'פתוחות', count: openCount, hint: 'עדיין בטיפול', color: '#f59e0b' },
    { key: 'overdue', label: 'באיחור', count: overdueCount, hint: 'עברו את תאריך היעד', color: '#ef4444' },
    { key: 'blocked', label: 'חסומות', count: blockedCount, hint: 'דורשות עזרה או החלטה', color: '#8b5cf6' },
    { key: 'done', label: 'בוצעו', count: doneCount, hint: 'משימות שהסתיימו', color: '#10b981' },
  ];

  container.innerHTML = metrics.map(m => `
    <div class="metric-card" style="--card-color: ${m.color}" data-metric="${m.key}">
      <div class="metric-label">${escapeHtml(m.label)}</div>
      <div class="metric-value">${m.count}</div>
      <div class="metric-hint">${escapeHtml(m.hint)}</div>
    </div>
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
    const color = priorityColor[pr] || '#2563eb';

    return `
      <div class="priority-row">
        <div class="priority-header">
          <div class="priority-badge-label">
            <span class="priority-badge ${priorityClass[pr] || ''}">${escapeHtml(pr)}</span>
            <span>${count} משימות</span>
          </div>
          <span style="font-weight:700; color:${color}">${pct}%</span>
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

  const itemsHtml = people.map(p => {
    const userTasks = openT.filter(t => t.owner === p);
    const count = userTasks.length;
    const init = initials(p);

    if (count === 0) {
      return `
        <div class="workload-card">
          <div class="workload-card-header">
            <div class="workload-user">
              <div class="user-avatar">${escapeHtml(init)}</div>
              <span class="workload-name">${escapeHtml(p || 'ללא שיוך')}</span>
            </div>
            <span class="workload-count-badge zero">0 משימות</span>
          </div>
          <div class="workload-empty-sub">אין משימות פתוחות כרגע</div>
        </div>
      `;
    }

    // Priority breakdown
    const pCounts = Object.assign(Object.create(null), { P0: 0, P1: 0, P2: 0, P3: 0 });
    userTasks.forEach(t => {
      const prio = t.priority || 'P2';
      if (pCounts[prio] !== undefined) pCounts[prio]++;
      else pCounts['P2']++;
    });

    const p0Pct = (pCounts.P0 / count) * 100;
    const p1Pct = (pCounts.P1 / count) * 100;
    const p2Pct = (pCounts.P2 / count) * 100;
    const p3Pct = (pCounts.P3 / count) * 100;

    // Category breakdown
    const catCounts = Object.create(null);
    userTasks.forEach(t => {
      const c = t.category || 'כללי';
      catCounts[c] = (catCounts[c] || 0) + 1;
    });

    const catPills = Object.entries(catCounts)
      .sort((a, b) => b[1] - a[1])
      .map(([catName, catCount]) => `
        <span class="workload-cat-pill" title="${escapeAttr(catName)}: ${catCount} משימות">
          <span class="cat-pill-dot"></span>
          ${escapeHtml(catName)} <strong>${catCount}</strong>
        </span>
      `).join('');

    return `
      <div class="workload-card">
        <div class="workload-card-header">
          <div class="workload-user">
            <div class="user-avatar">${escapeHtml(init)}</div>
            <span class="workload-name">${escapeHtml(p || 'ללא שיוך')}</span>
          </div>
          <span class="workload-count-badge">${count} משימות</span>
        </div>

        <!-- Stacked Priority Bar -->
        <div class="workload-bar-wrapper">
          <div class="workload-stacked-bar">
            ${p0Pct > 0 ? `<div class="bar-seg seg-p0" style="width: ${p0Pct}%;" title="P0 (דחוף): ${pCounts.P0}"></div>` : ''}
            ${p1Pct > 0 ? `<div class="bar-seg seg-p1" style="width: ${p1Pct}%;" title="P1 (גבוה): ${pCounts.P1}"></div>` : ''}
            ${p2Pct > 0 ? `<div class="bar-seg seg-p2" style="width: ${p2Pct}%;" title="P2 (בינוני): ${pCounts.P2}"></div>` : ''}
            ${p3Pct > 0 ? `<div class="bar-seg seg-p3" style="width: ${p3Pct}%;" title="P3 (נמוך): ${pCounts.P3}"></div>` : ''}
          </div>
          <div class="workload-bar-legend">
            ${pCounts.P0 > 0 ? `<span class="legend-chip chip-p0">P0: ${pCounts.P0}</span>` : ''}
            ${pCounts.P1 > 0 ? `<span class="legend-chip chip-p1">P1: ${pCounts.P1}</span>` : ''}
            ${pCounts.P2 > 0 ? `<span class="legend-chip chip-p2">P2: ${pCounts.P2}</span>` : ''}
            ${pCounts.P3 > 0 ? `<span class="legend-chip chip-p3">P3: ${pCounts.P3}</span>` : ''}
          </div>
        </div>

        <!-- Category Breakdown Pills -->
        <div class="workload-cats">
          ${catPills}
        </div>
      </div>
    `;
  }).join('');

  container.innerHTML = `<div class="workload-cards-grid">${itemsHtml}</div>`;
}

/**
 * Render Alerts Panel
 */
function renderAlerts() {
  const container = q('#alertsContainer');
  if (!container || !state) return;

  const tasks = visibleTasks();
  const alerts = tasks.filter(t => {
    return t.health === HEALTH_OVERDUE || t.health === HEALTH_BLOCKED || t.health === HEALTH_SOON;
  });

  if (alerts.length === 0) {
    container.innerHTML = `<div class="empty-state">אין התראות כרגע</div>`;
    return;
  }

  const itemsHtml = alerts.slice(0, 8).map(t => {
    let itemClass = '';
    let tagClass = '';
    let tagText = t.health;

    if (t.health === HEALTH_BLOCKED) {
      itemClass = 'alert-blocked';
      tagClass = 'alert-tag-blocked';
    } else if (t.health === HEALTH_OVERDUE) {
      itemClass = '';
      tagClass = 'alert-tag-overdue';
    } else if (t.health === HEALTH_SOON) {
      itemClass = 'alert-soon';
      tagClass = 'alert-tag-soon';
    }

    const dueText = t.dueDate ? `יעד: ${t.dueDate}` : 'ללא תאריך יעד';

    return `
      <div class="alert-item ${itemClass}" data-id="${escapeAttr(t.id)}">
        <div>
          <div class="alert-title">${escapeHtml(t.title)}</div>
          <div class="alert-meta">${escapeHtml(t.owner)} &bull; ${escapeHtml(dueText)}</div>
        </div>
        <span class="alert-tag ${tagClass}">${escapeHtml(tagText)}</span>
      </div>
    `;
  }).join('');

  container.innerHTML = `<div class="alerts-list">${itemsHtml}</div>`;

  container.querySelectorAll('.alert-item').forEach(el => {
    el.addEventListener('click', () => openTask(el.dataset.id));
  });
}

/**
 * Render Progress Summary
 */
function renderProgressSummary() {
  const container = q('#progressSummaryContainer');
  if (!container || !state) return;

  const tasks = state.tasks || [];
  const total = tasks.length || 1;
  const completed = tasks.filter(t => t.status === DONE).length;
  const inProgress = tasks.filter(t => t.status === 'בתהליך').length;
  const waiting = tasks.filter(t => t.status === 'בהמתנה' || t.status === 'בבדיקה').length;
  const notStarted = tasks.filter(t => t.status === 'לא התחיל').length;

  let sumProgress = 0;
  for (const t of tasks) {
    sumProgress += Number(t.progress) || 0;
  }
  const avgProgress = Math.round(sumProgress / total);

  container.innerHTML = `
    <div class="progress-summary">
      <div class="progress-overall-bar-container">
        <div class="progress-overall-stats">
          <span>ביצוע ממוצע כולל</span>
          <span style="color: var(--blue)">${avgProgress}%</span>
        </div>
        <div class="progress-overall-track">
          <div class="progress-overall-fill" style="width: ${avgProgress}%"></div>
        </div>
      </div>
      <div class="status-counts-pills">
        <div class="status-count-card">
          <div class="status-count-val" style="color: var(--green)">${completed}</div>
          <div class="status-count-lbl">הסתיימו</div>
        </div>
        <div class="status-count-card">
          <div class="status-count-val" style="color: var(--blue)">${inProgress}</div>
          <div class="status-count-lbl">בתהליך</div>
        </div>
        <div class="status-count-card">
          <div class="status-count-val" style="color: var(--amber)">${waiting + notStarted}</div>
          <div class="status-count-lbl">ממתינות / לא התחילו</div>
        </div>
      </div>
    </div>
  `;
}

/**
 * Generate HTML for Board Card
 */
function taskCardHtml(t) {
  const pClass = priorityClass[t.priority] || 'p2';
  const isChecked = selectedTasks.has(t.id) ? 'checked' : '';
  const dueDisplay = t.dueDate ? `יעד: ${t.dueDate}` : 'יעד: ללא';
  const healthClass = (t.health || '').replace(/\s+/g, '-');
  const prog = clampPercent(t.progress);

  return `
    <div class="task-card" draggable="true" data-id="${escapeAttr(t.id)}" data-owner="${escapeAttr(t.owner)}" data-priority="${escapeAttr(t.priority)}">
      <div class="card-top-row">
        <div class="card-top-left">
          <input type="checkbox" class="card-checkbox" data-id="${escapeAttr(t.id)}" ${isChecked}>
          <span class="priority-badge ${pClass}">${escapeHtml(t.priority)}</span>
        </div>
        <span class="health-pill ${escapeAttr(healthClass)}">${escapeHtml(t.health || HEALTH_ON_TRACK)}</span>
      </div>
      <div class="card-title" data-id="${escapeAttr(t.id)}">${escapeHtml(t.title)}</div>
      <div class="card-chips">
        ${t.category ? `<span class="chip">${escapeHtml(t.category)}</span>` : ''}
        <span class="chip">${escapeHtml(t.status)}</span>
        <span class="chip">${escapeHtml(dueDisplay)}</span>
      </div>
      <div class="card-progress-bar">
        <div class="card-progress-fill" style="width: ${prog}%"></div>
      </div>
      <div class="card-bottom-row">
        <span>${prog}% ביצוע</span>
        ${t.updatedBy ? `<span class="card-updated-hint" title="עודכן ע&quot;י ${escapeAttr(t.updatedBy)} ב-${escapeAttr(t.updatedAt || '')}">עודכן: ${escapeHtml(t.updatedBy)}</span>` : ''}
        <button class="btn-card-done" data-id="${escapeAttr(t.id)}">סיום</button>
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

    const emptyText = isSearching ? 'אין תוצאות תואמות' : 'אין משימות פתוחות';
    const cardsHtml = personTasks.length > 0
      ? personTasks.map(taskCardHtml).join('')
      : `<div class="empty-state">${emptyText}</div>`;

    const displayName = col.isUnassigned ? 'ללא שיוך' : person;
    const badgeExtra = col.isExtra ? '<span style="font-size:0.7rem;color:#f97316;margin-right:4px;">(מחוץ לרשימה)</span>' : '';

    return `
      <div class="board-column" data-owner="${escapeAttr(person)}">
        <div class="col-header">
          <div class="col-user-info">
            <div class="user-avatar">${escapeHtml(initials(displayName))}</div>
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
      markDone(btn.dataset.id);
    });
  });

  container.querySelectorAll('.card-title').forEach(el => {
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      openTask(el.dataset.id);
    });
  });

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

    await persistOrder(currentIds);
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
    const dueDisplay = t.dueDate || 'ללא';
    const healthClass = (t.health || '').replace(/\s+/g, '-');

    let rowClass = '';
    if (t.health === HEALTH_OVERDUE) rowClass = 'row-overdue';
    else if (t.health === HEALTH_BLOCKED) rowClass = 'row-blocked';

    return `
      <tr class="${rowClass}" data-id="${escapeAttr(t.id)}">
        <td class="td-cb">
          <input type="checkbox" class="row-checkbox" data-id="${escapeAttr(t.id)}" ${isChecked}>
        </td>
        <td class="td-title-cell" data-id="${escapeAttr(t.id)}">
          ${escapeHtml(t.title)}
        </td>
        <td>${escapeHtml(t.owner)}</td>
        <td><span class="priority-badge ${pClass}">${escapeHtml(t.priority)}</span></td>
        <td>${escapeHtml(t.status)}</td>
        <td>${escapeHtml(t.category)}</td>
        <td>${escapeHtml(dueDisplay)}</td>
        <td><span class="health-pill ${escapeAttr(healthClass)}">${escapeHtml(t.health || HEALTH_ON_TRACK)}</span></td>
        <td>${t.updatedBy ? `<span class="user-pill" title="זמן: ${escapeAttr(t.updatedAt || '')}">${escapeHtml(t.updatedBy)}</span>` : '<span style="color:var(--muted-light)">-</span>'}</td>
        <td>
          <button class="btn-table-edit" data-id="${escapeAttr(t.id)}">עריכה</button>
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
      <td style="white-space: nowrap; direction: ltr; text-align: right;">${escapeHtml(h.time)}</td>
      <td><span class="user-pill">${escapeHtml(h.user || 'Excel')}</span></td>
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
function parseNotesAndSubtasks(notesText) {
  if (!notesText) return { subtasks: [], cleanNotes: '' };
  const lines = String(notesText).split('\n');
  const subtasks = [];
  const otherLines = [];

  for (const line of lines) {
    const match = line.match(/^\s*(?:[-*]\s*)?\[([ xX])\]\s*(.*)$/);
    if (match) {
      subtasks.push({
        done: match[1].toLowerCase() === 'x',
        text: match[2].trim()
      });
    } else {
      otherLines.push(line);
    }
  }
  return { subtasks, cleanNotes: otherLines.join('\n').trim() };
}

function serializeNotesAndSubtasks(cleanNotes, subtasks) {
  const parts = [];
  if (subtasks && subtasks.length > 0) {
    for (const st of subtasks) {
      if (st.text) {
        parts.push(`- [${st.done ? 'x' : ' '}] ${st.text}`);
      }
    }
  }
  if (cleanNotes && cleanNotes.trim()) {
    if (parts.length > 0) parts.push('');
    parts.push(cleanNotes.trim());
  }
  return parts.join('\n');
}

function initDrawerSubtasks(notesText) {
  const parsed = parseNotesAndSubtasks(notesText);
  currentDrawerCleanNotes = parsed.cleanNotes;
  currentDrawerSubtasks = parsed.subtasks;
  renderDrawerSubtasks();
}

function syncDrawerNotesFromSubtasks() {
  const notesField = q('#formTaskNotes');
  if (notesField) {
    notesField.value = serializeNotesAndSubtasks(currentDrawerCleanNotes, currentDrawerSubtasks);
  }
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

  if (currentDrawerSubtasks.length === 0) {
    listEl.innerHTML = '<div style="font-size:0.8rem;color:var(--muted);padding:8px 0;text-align:center;">אין תת-משימות. ניתן להוסיף סעיפים למטה.</div>';
    if (counterEl) counterEl.textContent = '0/0 הושלמו (0%)';
    return;
  }

  const total = currentDrawerSubtasks.length;
  const done = currentDrawerSubtasks.filter(s => s.done).length;
  const pct = Math.round((done / total) * 100);
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
        syncDrawerNotesFromSubtasks();
        renderDrawerSubtasks();
      }
    });
  });

  listEl.querySelectorAll('.subtask-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      currentDrawerSubtasks.splice(idx, 1);
      updateChecklistProgress();
      syncDrawerNotesFromSubtasks();
      renderDrawerSubtasks();
    });
  });
}

/**
 * Drawer Operations
 */
function openNewTask() {
  if (state?.connection?.mode === 'excel' || cloudConnectPending) { toast('בחיבור הניסיוני ניתן לערוך משימות קיימות בלבד'); return; }

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
  q('#formTaskDueDate').value = '';
  q('#formTaskBlocked').value = NO;
  q('#formTaskProgress').value = 0;
  q('#formProgressValText').textContent = '0%';
  q('#formTaskNotes').value = '';

  initDrawerSubtasks('');

  const auditBox = q('#taskAuditBox');
  if (auditBox) auditBox.style.display = 'none';

  q('#drawerTitle').textContent = 'משימה חדשה';
  const btnDelete = q('#formBtnDelete');
  if (btnDelete) btnDelete.style.visibility = 'hidden';

  openDrawer();
}

function openNewTaskFor(owner) {
  if (state?.connection?.mode === 'excel' || cloudConnectPending) { toast('יצירת משימות אינה זמינה בחיבור הענן הניסיוני'); return; }
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
  q('#formTaskDueDate').value = task.dueDate || '';
  q('#formTaskBlocked').value = task.blocked === YES ? YES : NO;
  q('#formTaskProgress').value = task.progress || 0;
  q('#formProgressValText').textContent = `${task.progress || 0}%`;
  q('#formTaskNotes').value = task.notes || '';

  initDrawerSubtasks(task.notes || '');

  const auditBox = q('#taskAuditBox');
  if (auditBox) {
    if (task.createdBy || task.updatedBy) {
      auditBox.style.display = 'flex';
      const cBy = q('#drawerCreatedBy');
      const cAt = q('#drawerCreatedAt');
      const uBy = q('#drawerUpdatedBy');
      const uAt = q('#drawerUpdatedAt');
      if (cBy) cBy.textContent = task.createdBy || 'Excel';
      if (cAt) cAt.textContent = task.createdAt ? `(${task.createdAt})` : '';
      if (uBy) uBy.textContent = task.updatedBy || task.createdBy || '-';
      if (uAt) uAt.textContent = task.updatedAt ? `(${task.updatedAt})` : '';
    } else {
      auditBox.style.display = 'none';
    }
  }

  q('#drawerTitle').textContent = 'עריכת משימה';
  const btnDelete = q('#formBtnDelete');
  if (btnDelete) btnDelete.style.visibility = 'visible';

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

  // Ensure notes contain any subtasks updated in drawer
  syncDrawerNotesFromSubtasks();

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
    notes: q('#formTaskNotes').value,
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
    await api('/api/bulk', {
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
    toast(`${count} משימות עודכנו באקסל`);
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
        workbook: settingsDirty ? settingsWorkbook : state.workbook }
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
}

function clearSelection() {
  selectedTasks.clear();
  renderBulkBar();
  qa('.row-checkbox').forEach(cb => cb.checked = false);
  qa('.card-checkbox').forEach(cb => cb.checked = false);
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
    { type: 'action', title: 'מעבר לדשבורד', desc: 'תצוגת לוח הבקרה והמדדים המרכזיים', icon: '📊', run: () => setView('dashboard') },
    { type: 'action', title: 'מעבר לציר זמן (גאנט)', desc: 'מפת דרכים כרונולוגית לפי תאריכי ביצוע', icon: '📅', run: () => setView('timeline') },
    { type: 'action', title: 'מעבר ללוח קנבן', desc: 'תצוגת עמודות לפי חברי צוות', icon: '📋', run: () => setView('board') },
    { type: 'action', title: 'מעבר לטבלת משימות', desc: 'טבלה מפורטת עם כל הנתונים וההיסטוריה', icon: '📑', run: () => setView('tasks') },
    { type: 'action', title: 'המשימות שלי', desc: 'סינון מהיר של משימות המשויכות אלי', icon: '👤', run: () => applyMetricFilter('my-tasks') },
    { type: 'action', title: 'שיתוף סטטוס ב-WhatsApp', desc: 'ניסוח סיכום מוכן לשליחה מהירה', icon: '💬', run: () => openShareModal() },
    { type: 'action', title: 'פתח קובץ באקסל ישירות', desc: 'הפעלת Microsoft Excel עם הקובץ הפעיל', icon: '↗️', run: () => api('/api/open-excel', { method: 'POST' }).then(() => toast('נפתח ב-Excel ↗')).catch(e => toast(e.message)) },
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
        desc: `${t.owner} • ${t.priority} • ${t.status} • יעד: ${t.dueDate || 'ללא'}`,
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
      <div style="padding:4px 8px;text-align:right;">משימה / עובד</div>
      ${days.map(d => `
        <div class="timeline-day-col ${d.isToday ? 'today-col' : ''}" style="${d.isToday ? 'background:#e0f2fe;font-weight:800;color:#0284c7;' : ''}">
          <div>${d.dayName}</div>
          <div style="font-size:0.72rem;">${d.dayNum}/${d.monthNum}</div>
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
            <span class="priority-badge ${priorityClass[t.priority] || 'p2'}" style="margin-left:4px;padding:1px 5px;font-size:0.68rem;">${escapeHtml(t.priority)}</span>
            ${escapeHtml(t.title)}
          </div>
          <div class="timeline-bar-slot">
            ${inRange ? `<div class="timeline-bar ${hClass}" data-id="${escapeAttr(t.id)}" style="right:${leftPct}%;width:${widthPct}%;">
              <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${clampPercent(t.progress)}% • ${escapeHtml(t.title)}</span>
            </div>` : `<span class="timeline-out-of-range" style="color:var(--muted);font-size:0.75rem;">${tDue ? 'מחוץ לטווח התאריכים המוצג' : 'ללא תאריך יעד'}</span>`}
          </div>
        </div>
      `;
    }).join('');

    return `
      <div class="timeline-group-row">
        <div class="timeline-group-header">
          <span class="user-avatar-sm" style="width:24px;height:24px;font-size:0.7rem;">${escapeHtml(initials(grpName))}</span>
          <span>${escapeHtml(grpName)}</span>
          <span style="font-size:0.76rem;color:var(--muted);font-weight:normal;">(${grpTasks.length} משימות)</span>
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
      text += `• [${t.priority}] ${t.title} (${t.owner || 'לא משויך'}${t.dueDate ? ' | יעד: ' + t.dueDate : ''})\n`;
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
  if (errBox) {
    errBox.style.display = 'none';
    errBox.textContent = '';
  }

  // Pre-populate with current state
  const fnEl = q('#wbModalFileName');
  const pathEl = q('#wbModalFullPath');
  const tagEl = q('#wbModalStatusTag');
  const sizeEl = q('#wbModalFileSize');
  const timeEl = q('#wbModalFileModTime');
  const pathInput = q('#wbManualPathInput');

  if (state && state.workbook) {
    const p = state.workbook;
    const name = p.split(/[\\\/]/).pop();
    if (fnEl) fnEl.textContent = name;
    if (pathEl) pathEl.textContent = p;
    if (pathInput) pathInput.value = state.connection?.mode === 'excel' ? '' : p;
  }

  try {
    const info = await api('/api/workbook-info');
    if (info && info.ok) {
      if (fnEl) fnEl.textContent = info.name || '';
      if (pathEl) pathEl.textContent = info.path || '';
      if (pathInput && !pathInput.value && info.mode !== 'excel') pathInput.value = info.path || '';
      if (sizeEl) {
        const kb = (info.sizeBytes / 1024).toFixed(1);
        sizeEl.textContent = info.mode === 'excel' ? 'SharePoint דרך Excel' : 'גודל: ' + kb + ' KB';
      }
      if (timeEl) {
        timeEl.textContent = info.modified ? `עודכן: ${info.modified}` : '';
      }
      if (tagEl) {
        if (!info.exists) {
          tagEl.textContent = 'הקובץ לא נמצא';
          tagEl.className = 'wb-status-tag readonly';
        } else if (!info.writable) {
          tagEl.textContent = 'לקריאה בלבד';
          tagEl.className = 'wb-status-tag readonly';
        } else {
          tagEl.textContent = 'פעיל ותקין';
          tagEl.className = 'wb-status-tag';
        }
      }
    }
  } catch (err) {
    // If info call fails, current display remains
  }
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
function initEvents() {
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
    syncDrawerNotesFromSubtasks();
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

  // Live sync notes textarea with subtasks if user types manually
  const formNotes = q('#formTaskNotes');
  if (formNotes) {
    formNotes.addEventListener('input', () => {
      const parsed = parseNotesAndSubtasks(formNotes.value);
      const checklistChanged = JSON.stringify(parsed.subtasks) !== JSON.stringify(currentDrawerSubtasks);
      currentDrawerCleanNotes = parsed.cleanNotes;
      currentDrawerSubtasks = parsed.subtasks;
      if (checklistChanged) updateChecklistProgress();
      renderDrawerSubtasks();
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
          <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;font-family:Rubik,sans-serif;direction:rtl;background:#f8fafc;color:#0f172a;text-align:center;padding:20px;">
            <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:36px 44px;box-shadow:0 10px 25px -5px rgba(0,0,0,0.06);max-width:480px;">
              <div style="width:52px;height:52px;border-radius:50%;background:#dcfce7;color:#16a34a;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;font-size:24px;">✓</div>
              <h2 style="margin:0 0 12px;color:#0f172a;font-size:22px;font-weight:700;">האפליקציה נסגרה בהצלחה</h2>
              <p style="margin:0 0 24px;color:#64748b;font-size:15px;line-height:1.6;">השרת כובה לחלוטין וקובץ האקסל שמור ומעודכן. כעת ניתן לסגור חלון זה בבטחה.</p>
              <button onclick="window.close()" style="background:#0284c7;color:#fff;border:none;border-radius:6px;padding:10px 26px;font-size:14px;font-weight:600;cursor:pointer;">סגירת חלון הדפדפן</button>
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
    if (e.key === '1') setView('dashboard');
    if (e.key === '2') setView('timeline');
    if (e.key === '3') setView('board');
    if (e.key === '4') setView('tasks');
    if (e.key === '5') setView('settings');
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

  // Manual Path Apply (SharePoint / Network / Local)
  const btnApplyManual = q('#btnApplyManualPath');
  const manualInput = q('#wbManualPathInput');
  const wbErrBox = q('#wbModalError');

  if (btnApplyManual && manualInput) {
    const handleApplyPath = async () => {
      const p = (manualInput.value || '').trim();
      if (!p) {
        if (wbErrBox) {
          wbErrBox.textContent = 'נא להזין נתיב מלא לקובץ Excel (.xlsx)';
          wbErrBox.style.display = 'block';
        }
        return;
      }
      btnApplyManual.disabled = true;
      btnApplyManual.textContent = 'טוען... ⏳';
      if (wbErrBox) wbErrBox.style.display = 'none';

      try {
        const res = await api('/api/set-workbook', {
          method: 'POST',
          body: { path: p }
        });
        if (res && res.data) {
          adoptWorkbook(res.data);
          closeWorkbookModal();
          toast('קובץ האקסל הוחלף ונטען בהצלחה! ✓');
        }
      } catch (err) {
        if (wbErrBox) {
          wbErrBox.textContent = err.message || 'שגיאה בטעינת הקובץ';
          wbErrBox.style.display = 'block';
        }
      } finally {
        btnApplyManual.disabled = false;
        btnApplyManual.textContent = 'טען קובץ זה ✓';
      }
    };

    btnApplyManual.addEventListener('click', handleApplyPath);
    manualInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleApplyPath();
      }
    });
  }

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
        btnWinPicker.textContent = '🗂️ פתח סייר קבצים';
      }
    });
  }

  // File Upload via Browser
  const fileInput = q('#wbFileInput');
  if (fileInput) {
    fileInput.addEventListener('change', async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        if (wbErrBox) {
          wbErrBox.textContent = 'יש לבחור קובץ Excel עם סיומת .xlsx בלבד.';
          wbErrBox.style.display = 'block';
        }
        fileInput.value = '';
        return;
      }
      if (file.size > 20 * 1024 * 1024) {
        if (wbErrBox) {
          wbErrBox.textContent = 'הקובץ גדול מדי. הגודל המרבי הוא 20 MB.';
          wbErrBox.style.display = 'block';
        }
        fileInput.value = '';
        return;
      }

      const uploadLabel = q('#wbUploadText');
      const origText = uploadLabel ? uploadLabel.textContent : 'מעלה קובץ...';
      fileInput.disabled = true;
      if (uploadLabel) uploadLabel.textContent = 'מעלה... ⏳';
      if (wbErrBox) wbErrBox.style.display = 'none';

      try {
        const reader = new FileReader();
        reader.onload = async () => {
          try {
            const base64Data = reader.result.split(',')[1];
            const res = await api('/api/upload-workbook', {
              method: 'POST',
              body: {
                filename: file.name,
                content: base64Data
              }
            });
            if (res && res.data) {
              adoptWorkbook(res.data);
              closeWorkbookModal();
              toast('הקובץ הועלה ונטען בהצלחה! ✓');
            }
          } catch (uploadErr) {
            if (wbErrBox) {
              wbErrBox.textContent = uploadErr.message || 'שגיאה בהעלאת הקובץ';
              wbErrBox.style.display = 'block';
            }
          } finally {
            if (uploadLabel) uploadLabel.textContent = origText;
            fileInput.value = '';
            fileInput.disabled = false;
          }
        };
        reader.onerror = () => {
          if (wbErrBox) {
            wbErrBox.textContent = 'שגיאה בקריאת הקובץ מהמחשב';
            wbErrBox.style.display = 'block';
          }
          if (uploadLabel) uploadLabel.textContent = origText;
          fileInput.value = '';
          fileInput.disabled = false;
        };
        reader.readAsDataURL(file);
      } catch (err) {
        if (wbErrBox) {
          wbErrBox.textContent = err.message || 'שגיאה בהעלאת הקובץ';
          wbErrBox.style.display = 'block';
        }
        if (uploadLabel) uploadLabel.textContent = origText;
        fileInput.value = '';
        fileInput.disabled = false;
      }
    });
  }

  // Reset to default workbook
  const btnReset = q('#btnResetDefaultWb');
  if (btnReset) {
    btnReset.addEventListener('click', async () => {
      if (confirm('האם לחזור לקובץ ברירת המחדל המקורי של TeamFlow?')) {
        try {
          const res = await api('/api/reset-workbook', { method: 'POST' });
          if (res && res.data) {
            adoptWorkbook(res.data);
            closeWorkbookModal();
            toast('קובץ ברירת המחדל נטען בהצלחה! ✓');
          }
        } catch (err) {
          if (wbErrBox) {
            wbErrBox.textContent = err.message || 'שגיאה באיפוס הקובץ';
            wbErrBox.style.display = 'block';
          }
        }
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
    if (info.mode === 'excel') {
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
  q('#btnShowCloudExcel').addEventListener('click', async () => {
    try { await api('/api/cloud/show', {method: 'POST'}); }
    catch (err) { q('#cloudConnectFeedback').textContent = err.message; }
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
  const validViews = ['dashboard', 'timeline', 'board', 'tasks', 'settings'];
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
