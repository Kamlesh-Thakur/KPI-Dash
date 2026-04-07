/**
 * KPI Dashboard — Main Application Entries
 */
import './style.css';
import * as XLSX from 'xlsx';
import { createGrid } from 'ag-grid-community';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import {
  loadData, loadTeamPerformanceData, getFilteredRawData, getFilteredIncidentData,
  getFilteredRawDataForRange, getFilteredIncidentDataForRange, getDateFilterRange,
  setFilter, getUniqueValues, formatNumber,
  excelDateToJS, formatDuration, getState
} from './dataStore.js';
import { formatDisplayDate, formatDisplayDateRange } from './dateDisplay.js';
import {
  renderTasksTrend, renderTaskTypePie,
  renderPriorityDist, renderTaskTypeResolutionTargets,
  renderRepeatedSupportWindow, renderImmediateSupportWindow, renderRepeatedSupportByBranch, renderImmediateSupportByBranch,
  renderIncidentTrend, renderIncidentCategory,
  renderIncidentComplexity, renderIncidentClosure,
  renderBranchPerformance, renderBranchClosureRates, renderBranchEfficiencyFromSheet, renderBranchWorkloadFromSheet,
  renderSameDayClosure, renderBranchTasksVolume,
  renderTeamTopScores, renderTeamPerformanceByBranch, renderTeamOpsHealth,
  resizeCharts
} from './charts.js';
import { createCalendarPicker } from './calendarPicker.js';
import { mountCustomSelect, syncCustomSelect } from './customSelect.js';
import { getCalendarSystem, setCalendarSystem } from './calendarPrefs.js';
import { syncBsMonthHiddenFromGregorianYm, getNepaliMonthlyFilterDefaults } from './nepaliCalendarUi.js';

// ==========================================
// STATE
// ==========================================
let tasksGridApi = null;
let incidentsGridApi = null;
let currentTab = 'dashboard';
let dataLoaded = false;
let currentTheme = 'dark';
let showComparisons = false;

/** Themed calendar popover (week / day / month / range) */
let calendarPickerApi = null;
let compareDetailsOpen = false;
let compareDetailsHover = false;

// ==========================================
// INIT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initCompareToggle();
  initCompareDetails();
  initSidebarCollapse();
  initNavigation();
  initFilters();
  initDateFilters();
  initCalendarSystemToggle();
  mountFilterCustomSelects();
  populateFilters();
  initUpload();
  initExport();
  initMenuToggle();

  // Paint KPI + charts immediately so chart loaders clear even if Excel fetch fails or errors later
  try {
    renderAll();
  } catch (e) {
    console.error('[init] renderAll failed', e);
  }

  // Auto-load the Excel file
  autoLoadExcel();

  // Resize handler
  window.addEventListener('resize', () => {
    resizeCharts();
  });
});

// ==========================================
// AUTO LOAD
// ==========================================

/** KPI workbooks under public/data — loaded in order, rows merged (later file wins for Branch Effi. / Sheet2 / Drop Down). */
const KPI_DATA_FILES = ['kpi-falgun-2082.xlsx', 'kpi-chaitra-2082.xlsx'];

async function autoLoadExcel() {
  showToast('Loading data from Excel…', 'info');

  try {
    let merged = false;
    for (let i = 0; i < KPI_DATA_FILES.length; i++) {
      const name = KPI_DATA_FILES[i];
      const response = await fetch(`${import.meta.env.BASE_URL}data/${name}`);
      if (!response.ok) {
        showToast(`Missing ${name} in public/data/ or use Import`, 'info');
        return;
      }
      const buf = await response.arrayBuffer();
      const workbook = XLSX.read(buf, { type: 'array' });
      loadData(workbook, XLSX, { append: i > 0 });
      merged = true;
    }
    if (!merged) return;

    dataLoaded = true;
    populateFilters();
    try {
      renderAll();
    } catch (e) {
      console.error('[autoLoadExcel] renderAll failed', e);
      showToast('Dashboard render error — check console', 'info');
    }
    setTimeout(() => resizeCharts(), 50);
    setTimeout(() => resizeCharts(), 250);
    showToast(
      `Data loaded: ${formatNumber(getFilteredRawData().length)} tasks, ${formatNumber(getFilteredIncidentData().length)} incidents`,
      'success'
    );

    loadTeamPerformanceWorkbook();
  } catch (e) {
    console.log('No pre-loaded data file found. Use Import to load data.', e);
  }
}

async function loadTeamPerformanceWorkbook() {
  try {
    const response = await fetch(`${import.meta.env.BASE_URL}data/team-performance-monthly.xlsx`);
    if (!response.ok) return;
    const buf = await response.arrayBuffer();
    const workbook = XLSX.read(buf, { type: 'array' });
    loadTeamPerformanceData(workbook, XLSX);
    if (currentTab === 'team') {
      renderTabContent('team');
      resizeCharts();
    }
  } catch (e) {
    console.log('No team performance workbook found in public/data.');
  }
}

function processExcelBuffer(buf) {
  const workbook = XLSX.read(buf, { type: 'array' });
  loadData(workbook, XLSX);

  dataLoaded = true;
  populateFilters();
  try {
    renderAll();
  } catch (e) {
    console.error('[processExcelBuffer] renderAll failed', e);
    showToast('Dashboard render error — check console', 'info');
  }
  // Ensure charts pick up final layout sizes after first render
  setTimeout(() => resizeCharts(), 50);
  setTimeout(() => resizeCharts(), 250);
  showToast(`Data loaded: ${formatNumber(getFilteredRawData().length)} tasks, ${formatNumber(getFilteredIncidentData().length)} incidents`, 'success');
}

// ==========================================
// NAVIGATION
// ==========================================
function initNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(item => {
    item.addEventListener('click', () => {
      const tab = item.dataset.tab;
      switchTab(tab);
      closeMobileDrawer();
    });
  });
}

function switchTab(tab) {
  currentTab = tab;

  // Update nav
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.querySelector(`[data-tab="${tab}"]`).classList.add('active');

  // Update content
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.getElementById(`tab-${tab}`).classList.add('active');

  // Update title
  const titles = {
    dashboard: 'Dashboard',
    tasks: 'Task Data',
    incidents: 'Incidents',
    branch: 'Branch KPI',
    support: 'Support KPI',
    team: 'Team Performance'
  };
  document.getElementById('page-title').textContent = titles[tab] || 'Dashboard';

  // Render tab-specific content
  if (dataLoaded) {
    setTimeout(() => {
      renderTabContent(tab);
      resizeCharts();
    }, 50);
  }
}

// ==========================================
// FILTERS
// ==========================================
function initCalendarSystemToggle() {
  const en = document.getElementById('calendar-system-en');
  const ne = document.getElementById('calendar-system-ne');
  const sync = () => {
    const sys = getCalendarSystem();
    en?.classList.toggle('is-active', sys === 'english');
    ne?.classList.toggle('is-active', sys === 'nepali');
  };
  sync();
  en?.addEventListener('click', () => {
    setCalendarSystem('english');
    sync();
    calendarPickerApi?.refreshCalendarDisplay?.();
    renderAll();
    tasksGridApi?.refreshCells?.({ force: true });
    incidentsGridApi?.refreshCells?.({ force: true });
  });
  ne?.addEventListener('click', () => {
    setCalendarSystem('nepali');
    sync();
    calendarPickerApi?.refreshCalendarDisplay?.();
    renderAll();
    tasksGridApi?.refreshCells?.({ force: true });
    incidentsGridApi?.refreshCells?.({ force: true });
  });
}

function mountFilterCustomSelects() {
  ['filter-date-mode', 'filter-by-dimension', 'filter-by-value'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) mountCustomSelect(el);
  });
}

function initFilters() {
  const dimEl = document.getElementById('filter-by-dimension');
  const valEl = document.getElementById('filter-by-value');
  if (!dimEl || !valEl) return;

  dimEl.addEventListener('change', () => {
    setFilter('filterBy', dimEl.value);
    setFilter('division', '');
    setFilter('region', '');
    setFilter('branch', '');
    setFilter('taskType', '');
    populateFilters();
    renderAll();
  });

  valEl.addEventListener('change', () => {
    const fb = getState().filters.filterBy;
    if (!fb) return;
    const keyMap = {
      division: 'division',
      region: 'region',
      branch: 'branch',
      taskType: 'taskType'
    };
    setFilter(keyMap[fb], valEl.value);
    renderAll();
  });
}

function initDateFilters() {
  const modeEl = document.getElementById('filter-date-mode');
  const anchorEl = document.getElementById('filter-date-anchor');
  const weekEl = document.getElementById('filter-date-week');
  const weekFilterWrap = document.getElementById('week-filter-wrap');
  const dailyFilterWrap = document.getElementById('daily-filter-wrap');
  const monthFilterWrap = document.getElementById('month-filter-wrap');
  const rangeFilterWrap = document.getElementById('custom-range-filter-wrap');
  const monthEl = document.getElementById('filter-date-month');
  const fromEl = document.getElementById('filter-date-from');
  const toEl = document.getElementById('filter-date-to');

  if (
    !modeEl ||
    !anchorEl ||
    !weekEl ||
    !weekFilterWrap ||
    !dailyFilterWrap ||
    !monthFilterWrap ||
    !rangeFilterWrap ||
    !monthEl ||
    !fromEl ||
    !toEl
  ) {
    return;
  }

  const today = new Date();
  const todayText = [
    today.getFullYear(),
    String(today.getMonth() + 1).padStart(2, '0'),
    String(today.getDate()).padStart(2, '0')
  ].join('-');
  const monthStartText = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;
  const monthYm = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

  modeEl.value = 'monthly';
  anchorEl.value = todayText;
  weekEl.value = formatYmdLocal(getSundayOfWeek(today));
  fromEl.value = todayText;
  toEl.value = todayText;
  setFilter('dateMode', modeEl.value);
  setFilter('dateFrom', todayText);
  setFilter('dateTo', todayText);

  const bsMonthHidden = document.getElementById('filter-bs-month');
  if (getCalendarSystem() === 'nepali') {
    const nep = getNepaliMonthlyFilterDefaults(today);
    monthEl.value = nep.monthYYYYMM;
    if (bsMonthHidden) bsMonthHidden.value = nep.bsMonthCode;
    setFilter('dateAnchor', nep.dateAnchor);
  } else {
    monthEl.value = monthYm;
    syncBsMonthHiddenFromGregorianYm(monthYm);
    setFilter('dateAnchor', monthStartText);
  }

  calendarPickerApi = createCalendarPicker({
    setFilter,
    renderAll,
    formatYmdLocal,
    parseYmdLocal,
    getSundayOfWeek
  });
  calendarPickerApi.initCalendarPickerInner(today);

  const syncDateInputs = () => {
    const mode = modeEl.value;
    const showDaily = mode === 'daily';
    const showWeek = mode === 'weekly';
    const showMonth = mode === 'monthly';
    const showRange = mode === 'custom';

    dailyFilterWrap.classList.toggle('visible', showDaily);
    weekFilterWrap.classList.toggle('visible', showWeek);
    monthFilterWrap.classList.toggle('visible', showMonth);
    rangeFilterWrap.classList.toggle('visible', showRange);
    calendarPickerApi.syncCalTriggers({
      showWeek,
      showDaily,
      showMonth,
      showRange
    });
  };

  syncDateInputs();

  modeEl.addEventListener('change', (e) => {
    const mode = e.target.value;
    setFilter('dateMode', mode);
    if (mode === 'daily') {
      setFilter('dateAnchor', anchorEl.value);
    } else if (mode === 'weekly') {
      const base = parseYmdLocal(getState().filters.dateAnchor) || today;
      const sun = getSundayOfWeek(base);
      weekEl.value = formatYmdLocal(sun);
      setFilter('dateAnchor', weekEl.value);
    } else if (mode === 'monthly') {
      const bsMonthHidden = document.getElementById('filter-bs-month');
      if (getCalendarSystem() === 'nepali') {
        const nep = getNepaliMonthlyFilterDefaults(new Date());
        monthEl.value = nep.monthYYYYMM;
        if (bsMonthHidden) bsMonthHidden.value = nep.bsMonthCode;
        setFilter('dateAnchor', nep.dateAnchor);
      } else {
        syncBsMonthHiddenFromGregorianYm(monthEl.value);
        setFilter('dateAnchor', monthToDate(monthEl.value));
      }
    } else if (mode === 'custom') {
      setFilter('dateFrom', fromEl.value);
      setFilter('dateTo', toEl.value);
    }
    syncDateInputs();
    renderAll();
  });
}

/** Local calendar Sunday of the week containing `date` (week = Sunday–Saturday). */
function getSundayOfWeek(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() - d.getDay());
  return d;
}

function formatYmdLocal(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseYmdLocal(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || '').trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const day = Number(m[3]);
  const dt = new Date(y, mo, day);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== day) return null;
  return dt;
}

function monthToDate(monthValue) {
  const match = /^(\d{4})-(\d{2})$/.exec(monthValue || '');
  if (!match) return '';
  return `${match[1]}-${match[2]}-01`;
}

function populateFilters() {
  const dimEl = document.getElementById('filter-by-dimension');
  const valEl = document.getElementById('filter-by-value');
  if (!dimEl || !valEl) return;

  const f = getState().filters;
  const filterBy = f.filterBy || '';

  dimEl.value = filterBy;

  const placeholder = {
    '': 'Select value…',
    division: 'All divisions',
    region: 'All regions',
    branch: 'All branches',
    taskType: 'All task types'
  };

  valEl.innerHTML = '';
  if (!filterBy) {
    valEl.disabled = true;
    valEl.title = 'Choose a filter type first';
    valEl.appendChild(new Option(placeholder[''], ''));
    syncCustomSelect(dimEl);
    syncCustomSelect(valEl);
    return;
  }

  valEl.disabled = false;
  valEl.title = 'Choose a value';
  valEl.appendChild(new Option(placeholder[filterBy], ''));

  const columnMap = {
    division: 'Division',
    region: 'Region',
    branch: 'Branch',
    taskType: 'Task Type'
  };
  const col = columnMap[filterBy];
  getUniqueValues(col, 'raw').forEach((v) => {
    valEl.appendChild(new Option(v, v));
  });

  const keyMap = {
    division: 'division',
    region: 'region',
    branch: 'branch',
    taskType: 'taskType'
  };
  const cur = f[keyMap[filterBy]] || '';
  valEl.value = cur;
  syncCustomSelect(dimEl);
  syncCustomSelect(valEl);
}

// ==========================================
// FILE UPLOAD
// ==========================================
function initUpload() {
  const uploadArea = document.getElementById('upload-area');
  const fileInput = document.getElementById('file-input');

  uploadArea.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      processExcelBuffer(evt.target.result);
    };
    reader.readAsArrayBuffer(file);
  });

  // Drag & drop
  uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.style.borderColor = '#6384ff'; });
  uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = ''; });
  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '';
    const file = e.dataTransfer.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (evt) => processExcelBuffer(evt.target.result);
      reader.readAsArrayBuffer(file);
    }
  });
}

// ==========================================
// EXPORT
// ==========================================
function initExport() {
  document.getElementById('export-btn').addEventListener('click', () => {
    if (!dataLoaded) {
      showToast('No data loaded to export', 'error');
      return;
    }

    const wb = XLSX.utils.book_new();

    const rawData = getFilteredRawData();
    if (rawData.length > 0) {
      const ws = XLSX.utils.json_to_sheet(rawData);
      XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    }

    const incData = getFilteredIncidentData();
    if (incData.length > 0) {
      const ws = XLSX.utils.json_to_sheet(incData);
      XLSX.utils.book_append_sheet(wb, ws, 'Incidents');
    }

    XLSX.writeFile(wb, `KPI_Export_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast('Data exported successfully!', 'success');
  });
}

// ==========================================
// MENU TOGGLE (mobile drawer)
// ==========================================
const MOBILE_NAV_MQ = window.matchMedia('(max-width: 768px)');
const SIDEBAR_COLLAPSE_KEY = 'kpi-sidebar-collapsed';
const THEME_COOKIE_KEY = 'kpi-theme';
const COMPARE_COOKIE_KEY = 'kpi-show-comparisons';

function isMobileNavLayout() {
  return MOBILE_NAV_MQ.matches;
}

function setMobileDrawerOpen(open) {
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('sidebar-backdrop');
  if (!isMobileNavLayout()) {
    sidebar.classList.remove('open');
    backdrop.classList.remove('active');
    backdrop.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('sidebar-open');
    return;
  }
  sidebar.classList.toggle('open', open);
  backdrop.classList.toggle('active', open);
  backdrop.setAttribute('aria-hidden', open ? 'false' : 'true');
  document.body.classList.toggle('sidebar-open', open);
}

function closeMobileDrawer() {
  setMobileDrawerOpen(false);
}

function initMenuToggle() {
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('sidebar-backdrop');
  const toggle = document.getElementById('menu-toggle');
  const closeBtn = document.getElementById('sidebar-close');

  toggle.addEventListener('click', () => {
    setMobileDrawerOpen(!sidebar.classList.contains('open'));
  });
  backdrop.addEventListener('click', closeMobileDrawer);
  closeBtn.addEventListener('click', closeMobileDrawer);

  MOBILE_NAV_MQ.addEventListener('change', (e) => {
    if (!e.matches) setMobileDrawerOpen(false);
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && isMobileNavLayout() && sidebar.classList.contains('open')) {
      closeMobileDrawer();
    }
  });
}

function setSidebarCollapsed(collapsed) {
  if (isMobileNavLayout()) {
    document.body.classList.remove('sidebar-collapsed');
    setTimeout(() => resizeCharts(), 0);
    return;
  }
  document.body.classList.toggle('sidebar-collapsed', collapsed);
  setCookie(SIDEBAR_COLLAPSE_KEY, collapsed ? '1' : '0');
  const collapseBtn = document.getElementById('sidebar-collapse-toggle');
  if (collapseBtn) {
    collapseBtn.setAttribute('aria-label', collapsed ? 'Expand sidebar' : 'Collapse sidebar');
    collapseBtn.setAttribute('title', collapsed ? 'Expand sidebar' : 'Collapse sidebar');
  }
  // Resize immediately and after CSS transition so charts use full width.
  setTimeout(() => resizeCharts(), 0);
  setTimeout(() => resizeCharts(), 420);
}

function initSidebarCollapse() {
  const collapseBtn = document.getElementById('sidebar-collapse-toggle');
  if (!collapseBtn) return;

  const initialCollapsed = getCookie(SIDEBAR_COLLAPSE_KEY) === '1';
  setSidebarCollapsed(initialCollapsed);

  collapseBtn.addEventListener('click', () => {
    const collapsed = !document.body.classList.contains('sidebar-collapsed');
    setSidebarCollapsed(collapsed);
  });

  MOBILE_NAV_MQ.addEventListener('change', (e) => {
    if (e.matches) {
      document.body.classList.remove('sidebar-collapsed');
    } else {
      const persisted = getCookie(SIDEBAR_COLLAPSE_KEY) === '1';
      setSidebarCollapsed(persisted);
    }
  });
}

// ==========================================
// THEME
// ==========================================
function initTheme() {
  currentTheme = getCookie(THEME_COOKIE_KEY) || 'dark';
  applyTheme(currentTheme);

  const themeToggle = document.getElementById('theme-toggle');
  if (!themeToggle) return;

  themeToggle.addEventListener('click', () => {
    const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
    applyTheme(nextTheme);
  });
}

function applyTheme(theme) {
  currentTheme = theme === 'light' ? 'light' : 'dark';
  document.body.classList.toggle('theme-light', currentTheme === 'light');
  setCookie(THEME_COOKIE_KEY, currentTheme);
  updateThemeToggleUI();
  updateGridThemes();
  if (dataLoaded) {
    renderTabContent(currentTab);
  }
  resizeCharts();
}

function updateThemeToggleUI() {
  const themeToggle = document.getElementById('theme-toggle');
  const themeIcon = document.getElementById('theme-toggle-icon');
  if (!themeToggle || !themeIcon) return;

  const isLight = currentTheme === 'light';
  themeIcon.textContent = isLight ? '☀️' : '🌙';
  const target = isLight ? 'dark' : 'light';
  themeToggle.setAttribute('aria-label', `Switch to ${target} mode`);
  themeToggle.setAttribute('title', `Switch to ${target} mode`);
}

function initCompareToggle() {
  showComparisons = getCookie(COMPARE_COOKIE_KEY) === '1';
  updateCompareToggleUI();
  const compareToggle = document.getElementById('compare-toggle');
  if (!compareToggle) return;
  compareToggle.addEventListener('click', () => {
    showComparisons = !showComparisons;
    setCookie(COMPARE_COOKIE_KEY, showComparisons ? '1' : '0');
    updateCompareToggleUI();
    if (dataLoaded) renderKPICards();
  });
}

function updateCompareToggleUI() {
  const compareToggle = document.getElementById('compare-toggle');
  if (!compareToggle) return;
  compareToggle.classList.toggle('active', showComparisons);
  compareToggle.setAttribute('aria-pressed', showComparisons ? 'true' : 'false');
  compareToggle.setAttribute('title', showComparisons ? 'Hide comparisons' : 'Show comparisons');
}

function syncCompareDropdown() {
  const panel = document.getElementById('compare-context');
  const btn = document.getElementById('compare-details-btn');
  const anchor = document.getElementById('compare-details-anchor');
  if (!panel || !btn || !anchor) return;
  const showPanel = showComparisons && (compareDetailsOpen || compareDetailsHover);
  panel.hidden = !showPanel;
  btn.setAttribute('aria-expanded', showPanel ? 'true' : 'false');
  btn.classList.toggle('active', showPanel);
}

function initCompareDetails() {
  const anchor = document.getElementById('compare-details-anchor');
  const btn = document.getElementById('compare-details-btn');
  if (!anchor || !btn) return;
  let leaveTimer = null;
  anchor.addEventListener('mouseenter', () => {
    clearTimeout(leaveTimer);
    compareDetailsHover = true;
    syncCompareDropdown();
  });
  anchor.addEventListener('mouseleave', () => {
    leaveTimer = setTimeout(() => {
      compareDetailsHover = false;
      syncCompareDropdown();
    }, 220);
  });
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    compareDetailsOpen = !compareDetailsOpen;
    syncCompareDropdown();
  });
  document.addEventListener('click', (e) => {
    if (!showComparisons || !compareDetailsOpen) return;
    if (!anchor.contains(e.target)) {
      compareDetailsOpen = false;
      syncCompareDropdown();
    }
  });
}

function getCompareContextSummaryShort() {
  const win = getComparisonWindowDates();
  if (!win) return 'Comparison: pick a date filter to compare current, prior, and last year.';
  const cur = formatDisplayDateRange(win.thisStart, win.thisEnd);
  return `Current: ${cur}. Click or hover for full comparison windows.`;
}

/** Percentage (0–100) with two decimals for KPI rate cards. */
function formatKpiPercent(value) {
  if (value == null || Number.isNaN(value)) return '0.00%';
  return `${Number(value).toFixed(2)}%`;
}

function updateGridThemes() {
  const nextGridTheme = currentTheme === 'light' ? 'ag-theme-quartz' : 'ag-theme-quartz-dark';
  const altGridTheme = currentTheme === 'light' ? 'ag-theme-quartz-dark' : 'ag-theme-quartz';
  document.querySelectorAll('.ag-theme-custom').forEach((gridEl) => {
    gridEl.classList.remove(altGridTheme);
    gridEl.classList.add(nextGridTheme);
  });
}

// ==========================================
// RENDER ALL
// ==========================================
function renderAll() {
  renderKPICards();
  renderIncidentKPICards();
  renderTabContent(currentTab);
  // Charts may be initialized before layout settles; force a resize.
  setTimeout(() => resizeCharts(), 50);
}

function renderTabContent(tab) {
  switch (tab) {
    case 'dashboard':
      renderTasksTrend('chart-tasks-trend');
      renderTaskTypePie('chart-task-type-pie');
      renderIncidentCategory('chart-incident-category-dashboard');
      renderIncidentComplexity('chart-incident-complexity-dashboard');
      renderPriorityDist('chart-priority-dist');
      renderBranchTasksVolume('chart-branch-heatmap');
      renderTaskTypeResolutionTargets('chart-type-resolution-targets');
      break;
    case 'tasks':
      renderTasksGrid();
      break;
    case 'incidents':
      renderIncidentTrend('chart-incident-trend');
      renderIncidentCategory('chart-incident-category');
      renderIncidentComplexity('chart-incident-complexity');
      renderIncidentClosure('chart-incident-closure');
      renderIncidentsGrid();
      break;
    case 'branch':
      renderBranchPerformance('chart-branch-performance');
      renderBranchClosureRates('chart-branch-closure-rates');
      renderBranchEfficiencyFromSheet('chart-branch-efficiency-sheet');
      renderBranchWorkloadFromSheet('chart-branch-workload-sheet');
      renderSameDayClosure('chart-same-day-closure');
      renderBranchTasksVolume('chart-branch-tasks-volume');
      break;
    case 'support':
      renderRepeatedSupportWindow('chart-repeated-support-window');
      renderImmediateSupportWindow('chart-immediate-support-window');
      renderRepeatedSupportByBranch('chart-repeated-support-branch');
      renderImmediateSupportByBranch('chart-immediate-support-branch');
      break;
    case 'team':
      renderTeamTopScores('chart-team-top-scores');
      renderTeamPerformanceByBranch('chart-team-branch-performance');
      renderTeamOpsHealth('chart-team-ops-health');
      break;
  }
}

function setCookie(name, value, maxAgeSeconds = 31536000) {
  document.cookie = `${name}=${encodeURIComponent(value)}; path=/; max-age=${maxAgeSeconds}; SameSite=Lax`;
}

function getCookie(name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = document.cookie.match(new RegExp(`(?:^|; )${escaped}=([^;]*)`));
  return match ? decodeURIComponent(match[1]) : '';
}

// ==========================================
// KPI CARDS
// ==========================================
function rowDurationHours(r) {
  const d = parseFloat(r['Duration']);
  return Number.isNaN(d) ? null : d * 24;
}

function isKpiExclusionConsidered(r) {
  return (r['Exceptions'] || '').toString().toLowerCase().includes('consider');
}

/** Same duration / exception rules as Task KPI matrix (buildTaskKPIRows). */
function toRowMetrics(rows) {
  const total = rows.length;
  if (!total) {
    return {
      total: 0,
      sameDayRate: 0,
      avgDuration: 0,
      within4hRate: 0,
      within24hRate: 0,
      kpiExclusionRate: 0,
      sameDay: 0,
      within4h: 0,
      within24h: 0,
      kpiExclusion: 0
    };
  }
  const sameDay = rows.filter(r => r['Same day Closure'] === true).length;
  const avgDuration = rows.reduce((s, r) => s + (parseFloat(r['Duration']) || 0), 0) / total;
  let within4h = 0;
  let within24h = 0;
  let kpiExclusion = 0;
  for (const r of rows) {
    const h = rowDurationHours(r);
    if (h != null) {
      if (h <= 4) within4h += 1;
      if (h <= 24) within24h += 1;
    }
    if (isKpiExclusionConsidered(r)) kpiExclusion += 1;
  }
  return {
    total,
    sameDayRate: (sameDay / total) * 100,
    avgDuration,
    within4hRate: (within4h / total) * 100,
    within24hRate: (within24h / total) * 100,
    kpiExclusionRate: (kpiExclusion / total) * 100,
    sameDay,
    within4h,
    within24h,
    kpiExclusion
  };
}

function renderKPICards() {
  const data = getFilteredRawData();
  const incidentData = getFilteredIncidentData();
  const container = document.getElementById('kpi-cards');

  const taskM = toRowMetrics(data);
  const total = taskM.total;
  const sameDay = taskM.sameDay;
  const avgDuration = taskM.avgDuration;
  const taskTypes = new Set(data.map(r => r['Task Type'])).size;
  const incM = toRowMetrics(incidentData);
  const incidentTotal = incM.total;
  const incidentSameDay = incM.sameDay;
  const incidentAvgDuration = incM.avgDuration;
  const compare = showComparisons ? buildComparisons() : emptyCompare();
  const compareChipLabels = showComparisons ? getCompareChipLabels() : null;
  if (!showComparisons) {
    compareDetailsOpen = false;
    compareDetailsHover = false;
  }
  const compareContextEl = document.getElementById('compare-context');
  const compareDetailsAnchor = document.getElementById('compare-details-anchor');
  const compareDetailsBtn = document.getElementById('compare-details-btn');
  if (compareContextEl) {
    compareContextEl.innerHTML = showComparisons ? buildComparisonContextLabel() : '';
  }
  if (compareDetailsAnchor) {
    compareDetailsAnchor.hidden = !showComparisons;
  }
  if (compareDetailsBtn) {
    compareDetailsBtn.title = showComparisons
      ? getCompareContextSummaryShort()
      : 'Available when Compare is on';
  }
  syncCompareDropdown();

  container.innerHTML = `
    <div class="kpi-section kpi-section--tasks" role="group" aria-label="Tasks KPIs">
      <div class="kpi-section-label">Tasks</div>
      <div class="kpi-section-inner">
        <div class="kpi-card blue" title="Total — count for current filters">
          <div class="kpi-label">Total</div>
          <div class="kpi-value">${formatNumber(total)}</div>
          <div class="kpi-change">${taskTypes} types</div>
          ${showComparisons ? renderCompareRow(compare.tasks.total, false, compareChipLabels) : ''}
        </div>
        <div class="kpi-card cyan" title="Task closed within 4 hours (duration)">
          <div class="kpi-label">Within 4h</div>
          <div class="kpi-value">${total > 0 ? formatKpiPercent(taskM.within4hRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(taskM.within4h)} / ${formatNumber(total)}</div>
          ${showComparisons ? renderCompareRow(compare.tasks.within4hRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card purple" title="Task closed within 24 hours (duration)">
          <div class="kpi-label">Within 24h</div>
          <div class="kpi-value">${total > 0 ? formatKpiPercent(taskM.within24hRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(taskM.within24h)} / ${formatNumber(total)}</div>
          ${showComparisons ? renderCompareRow(compare.tasks.within24hRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card indigo" title="KPI after exclusion — Exceptions contains &quot;consider&quot;">
          <div class="kpi-label">KPI After Excl.</div>
          <div class="kpi-value">${total > 0 ? formatKpiPercent(taskM.kpiExclusionRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(taskM.kpiExclusion)} / ${formatNumber(total)}</div>
          ${showComparisons ? renderCompareRow(compare.tasks.kpiExclusionRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card emerald" title="Same-day closure">
          <div class="kpi-label">Same Day</div>
          <div class="kpi-value">${total > 0 ? formatKpiPercent(taskM.sameDayRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(sameDay)} / ${formatNumber(total)}</div>
          ${showComparisons ? renderCompareRow(compare.tasks.sameDayRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card amber" title="Average task duration">
          <div class="kpi-label">Avg Duration</div>
          <div class="kpi-value">${formatDuration(avgDuration)}</div>
          <div class="kpi-change">per task</div>
          ${showComparisons ? renderCompareRow(compare.tasks.avgDuration, true, compareChipLabels) : ''}
        </div>
      </div>
    </div>
    <div class="kpi-section kpi-section--incidents" role="group" aria-label="Incidents KPIs">
      <div class="kpi-section-label">Incidents</div>
      <div class="kpi-section-inner">
        <div class="kpi-card red" title="Total incidents — fiber network">
          <div class="kpi-label">Total</div>
          <div class="kpi-value">${formatNumber(incidentTotal)}</div>
          <div class="kpi-change">fiber network</div>
          ${showComparisons ? renderCompareRow(compare.incidents.total, false, compareChipLabels) : ''}
        </div>
        <div class="kpi-card cyan" title="Incident closed within 4 hours (duration)">
          <div class="kpi-label">Within 4h</div>
          <div class="kpi-value">${incidentTotal > 0 ? formatKpiPercent(incM.within4hRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(incM.within4h)} / ${formatNumber(incidentTotal)}</div>
          ${showComparisons ? renderCompareRow(compare.incidents.within4hRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card purple" title="Incident closed within 24 hours (duration)">
          <div class="kpi-label">Within 24h</div>
          <div class="kpi-value">${incidentTotal > 0 ? formatKpiPercent(incM.within24hRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(incM.within24h)} / ${formatNumber(incidentTotal)}</div>
          ${showComparisons ? renderCompareRow(compare.incidents.within24hRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card indigo" title="KPI after exclusion — Exceptions contains &quot;consider&quot;">
          <div class="kpi-label">KPI After Excl.</div>
          <div class="kpi-value">${incidentTotal > 0 ? formatKpiPercent(incM.kpiExclusionRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(incM.kpiExclusion)} / ${formatNumber(incidentTotal)}</div>
          ${showComparisons ? renderCompareRow(compare.incidents.kpiExclusionRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card emerald" title="Same-day closure">
          <div class="kpi-label">Same Day</div>
          <div class="kpi-value">${incidentTotal > 0 ? formatKpiPercent(incM.sameDayRate) : '0.00%'}</div>
          <div class="kpi-change up">${formatNumber(incidentSameDay)} / ${formatNumber(incidentTotal)}</div>
          ${showComparisons ? renderCompareRow(compare.incidents.sameDayRate, true, compareChipLabels) : ''}
        </div>
        <div class="kpi-card amber" title="Average incident duration">
          <div class="kpi-label">Avg Duration</div>
          <div class="kpi-value">${formatDuration(incidentAvgDuration)}</div>
          <div class="kpi-change">per incident</div>
          ${showComparisons ? renderCompareRow(compare.incidents.avgDuration, true, compareChipLabels) : ''}
        </div>
      </div>
    </div>
  `;
}

/** Comparison windows aligned with buildComparisons / buildComparisonContextLabel */
function getComparisonWindowDates() {
  const range = getDateFilterRange();
  const thisStart = range.start;
  const thisEnd = range.end;
  if (!thisStart || !thisEnd) return null;
  const daySpan = Math.max(1, Math.round((thisEnd - thisStart) / 86400000) + 1);
  const prevEnd = new Date(thisStart);
  prevEnd.setDate(prevEnd.getDate() - 1);
  const prevStart = new Date(prevEnd);
  prevStart.setDate(prevEnd.getDate() - (daySpan - 1));
  const yoyStart = new Date(thisStart);
  yoyStart.setFullYear(yoyStart.getFullYear() - 1);
  const yoyEnd = new Date(thisEnd);
  yoyEnd.setFullYear(yoyEnd.getFullYear() - 1);
  return { thisStart, thisEnd, prevStart, prevEnd, yoyStart, yoyEnd, daySpan };
}

function getCompareChipLabels() {
  const mode = getState().filters?.dateMode || 'overall';
  const byMode = {
    daily: {
      prevCode: 'DoD',
      yoyCode: 'YoY',
      prevTitle: 'Change vs the previous calendar day',
      yoyTitle: 'Change vs the same calendar day last year',
      prevPeriod: 'Prior day',
      yoyPeriod: 'Same day last year'
    },
    weekly: {
      prevCode: 'WoW',
      yoyCode: 'YoY',
      prevTitle: 'Change vs the previous week (same length)',
      yoyTitle: 'Change vs the same dates last year',
      prevPeriod: 'Prior week',
      yoyPeriod: 'Same span last year'
    },
    monthly: {
      prevCode: 'MoM',
      yoyCode: 'YoY',
      prevTitle: 'Change vs the previous calendar month',
      yoyTitle: 'Change vs the same calendar month last year',
      prevPeriod: 'Prior month',
      yoyPeriod: 'Same month last year'
    },
    custom: {
      prevCode: 'Prior',
      yoyCode: 'YoY',
      prevTitle: 'Change vs the previous period of equal length',
      yoyTitle: 'Change vs the same calendar span last year',
      prevPeriod: 'Prior period',
      yoyPeriod: 'Same span last year'
    },
    overall: {
      prevCode: 'Prior',
      yoyCode: 'YoY',
      prevTitle: 'Change vs the previous period of equal length (immediately before your data range)',
      yoyTitle: 'Change vs the same calendar span last year',
      prevPeriod: 'Prior period',
      yoyPeriod: 'Same span last year'
    }
  };
  return byMode[mode] || byMode.overall;
}

function getCompareModeHeadline() {
  const mode = getState().filters?.dateMode || 'overall';
  const map = {
    daily: 'Comparing to the prior day and the same day last year',
    weekly: 'Comparing to the prior week and the same dates last year',
    monthly: 'Comparing to the prior month and the same month last year',
    custom: 'Comparing to the prior period (same length) and the same span last year',
    overall: 'Comparing to the prior period (same length) and the same span last year'
  };
  return map[mode] || map.overall;
}

function escapeAttr(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function buildComparisons() {
  const win = getComparisonWindowDates();
  if (!win) {
    return emptyCompare();
  }
  const { thisStart, thisEnd, prevStart, prevEnd, yoyStart, yoyEnd } = win;

  const currTasks = getFilteredRawDataForRange(thisStart, thisEnd);
  const prevTasks = getFilteredRawDataForRange(prevStart, prevEnd);
  const yoyTasks = getFilteredRawDataForRange(yoyStart, yoyEnd);
  const currInc = getFilteredIncidentDataForRange(thisStart, thisEnd);
  const prevInc = getFilteredIncidentDataForRange(prevStart, prevEnd);
  const yoyInc = getFilteredIncidentDataForRange(yoyStart, yoyEnd);

  return {
    tasks: {
      total: deltaPair(toRowMetrics(currTasks).total, toRowMetrics(prevTasks).total, toRowMetrics(yoyTasks).total),
      sameDayRate: deltaPair(toRowMetrics(currTasks).sameDayRate, toRowMetrics(prevTasks).sameDayRate, toRowMetrics(yoyTasks).sameDayRate),
      avgDuration: deltaPair(toRowMetrics(currTasks).avgDuration, toRowMetrics(prevTasks).avgDuration, toRowMetrics(yoyTasks).avgDuration),
      within4hRate: deltaPair(toRowMetrics(currTasks).within4hRate, toRowMetrics(prevTasks).within4hRate, toRowMetrics(yoyTasks).within4hRate),
      within24hRate: deltaPair(toRowMetrics(currTasks).within24hRate, toRowMetrics(prevTasks).within24hRate, toRowMetrics(yoyTasks).within24hRate),
      kpiExclusionRate: deltaPair(toRowMetrics(currTasks).kpiExclusionRate, toRowMetrics(prevTasks).kpiExclusionRate, toRowMetrics(yoyTasks).kpiExclusionRate)
    },
    incidents: {
      total: deltaPair(toRowMetrics(currInc).total, toRowMetrics(prevInc).total, toRowMetrics(yoyInc).total),
      sameDayRate: deltaPair(toRowMetrics(currInc).sameDayRate, toRowMetrics(prevInc).sameDayRate, toRowMetrics(yoyInc).sameDayRate),
      avgDuration: deltaPair(toRowMetrics(currInc).avgDuration, toRowMetrics(prevInc).avgDuration, toRowMetrics(yoyInc).avgDuration),
      within4hRate: deltaPair(toRowMetrics(currInc).within4hRate, toRowMetrics(prevInc).within4hRate, toRowMetrics(yoyInc).within4hRate),
      within24hRate: deltaPair(toRowMetrics(currInc).within24hRate, toRowMetrics(prevInc).within24hRate, toRowMetrics(yoyInc).within24hRate),
      kpiExclusionRate: deltaPair(toRowMetrics(currInc).kpiExclusionRate, toRowMetrics(prevInc).kpiExclusionRate, toRowMetrics(yoyInc).kpiExclusionRate)
    }
  };
}

function deltaPair(current, previous, yoy) {
  return {
    mom: calcDelta(current, previous),
    yoy: calcDelta(current, yoy)
  };
}

function calcDelta(current, base) {
  if (!base) return 0;
  return ((current - base) / Math.abs(base)) * 100;
}

function renderCompareRow(pair, isPercent, chipLabels) {
  const L = chipLabels || getCompareChipLabels();
  return `<div class="kpi-compare">
    <span class="compare-chip ${pair.mom >= 0 ? 'up' : 'down'}" title="${escapeAttr(L.prevTitle)}"><span class="compare-chip-label">${escapeAttr(L.prevCode)}</span><span class="compare-chip-delta">${formatDelta(pair.mom, isPercent)}</span></span>
    <span class="compare-chip ${pair.yoy >= 0 ? 'up' : 'down'}" title="${escapeAttr(L.yoyTitle)}"><span class="compare-chip-label">${escapeAttr(L.yoyCode)}</span><span class="compare-chip-delta">${formatDelta(pair.yoy, isPercent)}</span></span>
  </div>`;
}

function formatDelta(value, _isPercent) {
  const sign = value >= 0 ? '+' : '';
  return `${sign}${value.toFixed(1)}%`;
}

function emptyCompare() {
  const zero = { mom: 0, yoy: 0 };
  return {
    tasks: {
      total: zero,
      sameDayRate: zero,
      avgDuration: zero,
      within4hRate: zero,
      within24hRate: zero,
      kpiExclusionRate: zero
    },
    incidents: {
      total: zero,
      sameDayRate: zero,
      avgDuration: zero,
      within4hRate: zero,
      within24hRate: zero,
      kpiExclusionRate: zero
    }
  };
}

function buildComparisonContextLabel() {
  const win = getComparisonWindowDates();
  const L = getCompareChipLabels();
  if (!win) {
    return '<p class="compare-context-muted">Choose a date filter so comparisons can use a current window, a prior window, and last year.</p>';
  }
  const { thisStart, thisEnd, prevStart, prevEnd, yoyStart, yoyEnd } = win;
  const cur = formatDisplayDateRange(thisStart, thisEnd);
  const prev = formatDisplayDateRange(prevStart, prevEnd);
  const yoy = formatDisplayDateRange(yoyStart, yoyEnd);
  const headline = getCompareModeHeadline();

  return `
    <div class="compare-context-inner">
      <p class="compare-context-headline">${headline}</p>
      <dl class="compare-context-dl">
        <div class="compare-context-row"><dt>Current</dt><dd>${cur}</dd></div>
        <div class="compare-context-row"><dt>${L.prevPeriod} <span class="compare-context-code">(${L.prevCode})</span></dt><dd>${prev}</dd></div>
        <div class="compare-context-row"><dt>${L.yoyPeriod} <span class="compare-context-code">(${L.yoyCode})</span></dt><dd>${yoy}</dd></div>
      </dl>
      <p class="compare-context-hint">Percent changes on each card match the two chips (${L.prevCode} and ${L.yoyCode}). Hover a chip for a full explanation.</p>
    </div>
  `;
}

function renderIncidentKPICards() {
  const data = getFilteredIncidentData();
  const container = document.getElementById('incident-kpi-cards');

  const total = data.length;
  const sameDay = data.filter(r => r['Same day Closure'] === true).length;
  const delayed = data.filter(r => (r['Exceptions'] || '').toLowerCase().includes('delayed')).length;
  const categories = new Set(data.map(r => r['Category'])).size;

  container.innerHTML = `
    <div class="kpi-card red">
      <div class="kpi-label">Total Incidents</div>
      <div class="kpi-value">${formatNumber(total)}</div>
      <div class="kpi-change">${categories} categories</div>
    </div>
    <div class="kpi-card emerald">
      <div class="kpi-label">Same-Day Resolved</div>
      <div class="kpi-value">${formatNumber(sameDay)}</div>
      <div class="kpi-change up">${total > 0 ? Math.round((sameDay / total) * 100) : 0}%</div>
    </div>
    <div class="kpi-card amber">
      <div class="kpi-label">Delayed</div>
      <div class="kpi-value">${formatNumber(delayed)}</div>
      <div class="kpi-change down">${total > 0 ? Math.round((delayed / total) * 100) : 0}%</div>
    </div>
    <div class="kpi-card blue">
      <div class="kpi-label">Considered</div>
      <div class="kpi-value">${formatNumber(total - delayed)}</div>
      <div class="kpi-change">on-time incidents</div>
    </div>
  `;
}

// ==========================================
// DATA GRIDS
// ==========================================
function renderTasksGrid() {
  const data = getFilteredRawData();
  const container = document.getElementById('tasks-grid');

  document.getElementById('tasks-row-count').textContent = `${formatNumber(data.length)} rows`;

  if (tasksGridApi) {
    tasksGridApi.setGridOption('rowData', data);
    return;
  }

  // Clear container
  container.innerHTML = '';
  const gridDiv = document.createElement('div');
  gridDiv.className = `ag-theme-custom ${currentTheme === 'light' ? 'ag-theme-quartz' : 'ag-theme-quartz-dark'}`;
  gridDiv.style.height = '600px';
  gridDiv.style.width = '100%';
  container.appendChild(gridDiv);

  const columnDefs = [
    { field: 'SR', headerName: '#', width: 70, pinned: 'left' },
    { field: 'Task Type', width: 140, filter: true },
    { field: 'Task ID', width: 120 },
    { field: 'Task Priority', width: 100, filter: true,
      cellRenderer: (p) => {
        const colors = { 0: 'blue', 1: 'green', 2: 'amber', 3: 'red', 4: 'red' };
        return `<span class="status-badge ${colors[p.value] || 'blue'}">P${p.value}</span>`;
      }
    },
    { field: 'Branch', width: 130, filter: true },
    { field: 'Region', width: 100, filter: true },
    { field: 'Division', width: 90, filter: true },
    { field: 'Cluster', width: 120 },
    { field: 'Service Type', width: 150 },
    { field: 'Customer Category', width: 130 },
    { field: 'Task Field Agent', width: 150 },
    { field: 'Task Created', width: 120,
      valueFormatter: p => formatDisplayDate(excelDateToJS(p.value))
    },
    { field: 'Task Assigned', width: 120,
      valueFormatter: p => formatDisplayDate(excelDateToJS(p.value))
    },
    { field: 'Task Completed', width: 120,
      valueFormatter: p => formatDisplayDate(excelDateToJS(p.value))
    },
    { field: 'Service Days', width: 100 },
    { field: 'Duration', width: 100,
      valueFormatter: p => formatDuration(p.value)
    },
    { field: 'Same day Closure', width: 120,
      cellRenderer: (p) => {
        const color = p.value === true ? 'green' : 'red';
        const text = p.value === true ? 'Yes' : 'No';
        return `<span class="status-badge ${color}">${text}</span>`;
      }
    },
    { field: 'Support Action', width: 140 },
    { field: 'Assigned Day', width: 110 },
    { field: 'Completed Day', width: 120 },
    { field: 'Priority Escalated', width: 120,
      cellRenderer: (p) => {
        const color = p.value === 'Y' ? 'red' : 'green';
        return `<span class="status-badge ${color}">${p.value}</span>`;
      }
    },
    { field: 'Repeated Fiber Support', width: 140 },
    { field: 'Exceptions', width: 120,
      cellRenderer: (p) => {
        const v = p.value || '';
        const color = v.toLowerCase().includes('delay') ? 'red' : v.toLowerCase().includes('consider') ? 'green' : 'blue';
        return `<span class="status-badge ${color}">${v}</span>`;
      }
    },
    { field: 'Remarks', width: 200 }
  ];

  const gridOptions = {
    columnDefs,
    rowData: data,
    defaultColDef: {
      sortable: true,
      resizable: true,
      editable: true,
      minWidth: 70
    },
    pagination: true,
    paginationPageSize: 50,
    paginationPageSizeSelector: [20, 50, 100, 500],
    rowSelection: 'multiple',
    animateRows: true,
    suppressMovableColumns: false,
    enableCellTextSelection: true
  };

  tasksGridApi = createGrid(gridDiv, gridOptions);

  // Search
  document.getElementById('tasks-search').addEventListener('input', (e) => {
    tasksGridApi.setGridOption('quickFilterText', e.target.value);
  });

  // Add row
  document.getElementById('add-task-btn').addEventListener('click', () => {
    const newRow = { SR: data.length + 1, 'Task Type': '', Branch: '', Region: '' };
    tasksGridApi.applyTransaction({ add: [newRow], addIndex: 0 });
    showToast('New row added', 'success');
  });

  // Delete selected
  document.getElementById('delete-task-btn').addEventListener('click', () => {
    const selected = tasksGridApi.getSelectedRows();
    if (selected.length === 0) {
      showToast('Select rows to delete', 'error');
      return;
    }
    tasksGridApi.applyTransaction({ remove: selected });
    showToast(`${selected.length} row(s) deleted`, 'success');
  });
}

function renderIncidentsGrid() {
  const data = getFilteredIncidentData();
  const container = document.getElementById('incidents-grid');

  document.getElementById('incidents-row-count').textContent = `${formatNumber(data.length)} rows`;

  if (incidentsGridApi) {
    incidentsGridApi.setGridOption('rowData', data);
    return;
  }

  container.innerHTML = '';
  const gridDiv = document.createElement('div');
  gridDiv.className = `ag-theme-custom ${currentTheme === 'light' ? 'ag-theme-quartz' : 'ag-theme-quartz-dark'}`;
  gridDiv.style.height = '500px';
  gridDiv.style.width = '100%';
  container.appendChild(gridDiv);

  const columnDefs = [
    { field: 'Sr.No.', headerName: '#', width: 70, pinned: 'left' },
    { field: 'Incident ID', width: 110 },
    { field: 'Incident Description', width: 250 },
    { field: 'Incident Type', width: 180, filter: true },
    { field: 'Category', width: 200, filter: true },
    { field: 'Complexity', width: 100, filter: true,
      cellRenderer: (p) => {
        const colors = { 'Low': 'green', 'Normal': 'blue', 'High': 'amber', 'Critical': 'red' };
        return `<span class="status-badge ${colors[p.value] || 'blue'}">${p.value}</span>`;
      }
    },
    { field: 'Branch', width: 180 },
    { field: 'Clients', width: 100 },
    { field: 'POP', width: 180 },
    { field: 'Assignments', width: 150 },
    { field: 'Duration', width: 120 },
    { field: 'Same day Closure', width: 120,
      cellRenderer: (p) => {
        const color = p.value === true ? 'green' : 'red';
        const text = p.value === true ? 'Yes' : 'No';
        return `<span class="status-badge ${color}">${text}</span>`;
      }
    },
    { field: 'Region', width: 100, filter: true },
    { field: 'Division', width: 90, filter: true },
    { field: 'Assigned Day', width: 110 },
    { field: 'Completed Day', width: 120 },
    { field: 'Related Tickets', width: 110 },
    { field: 'Incident Delayed?', width: 200,
      cellRenderer: (p) => {
        const v = (p.value || '').toString();
        if (v.toLowerCase().startsWith('delay')) {
          return `<span class="status-badge red" title="${v}">Delayed</span>`;
        }
        return `<span class="status-badge green">${v}</span>`;
      }
    },
    { field: 'Exceptions', width: 120,
      cellRenderer: (p) => {
        const v = p.value || '';
        const color = v.toLowerCase().includes('delay') ? 'red' : 'green';
        return `<span class="status-badge ${color}">${v}</span>`;
      }
    }
  ];

  const gridOptions = {
    columnDefs,
    rowData: data,
    defaultColDef: {
      sortable: true,
      resizable: true,
      editable: true,
      minWidth: 70
    },
    pagination: true,
    paginationPageSize: 50,
    paginationPageSizeSelector: [20, 50, 100, 500],
    rowSelection: 'multiple',
    animateRows: true,
    enableCellTextSelection: true
  };

  incidentsGridApi = createGrid(gridDiv, gridOptions);

  // Search
  document.getElementById('incidents-search').addEventListener('input', (e) => {
    incidentsGridApi.setGridOption('quickFilterText', e.target.value);
  });
}

// ==========================================
// TOAST NOTIFICATIONS
// ==========================================
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  const icons = {
    success: '✓',
    error: '✕',
    info: 'ℹ'
  };

  toast.innerHTML = `<span>${icons[type] || 'ℹ'}</span> ${message}`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(40px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}
