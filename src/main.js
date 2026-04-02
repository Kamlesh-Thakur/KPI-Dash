/**
 * KPI Dashboard — Main Application Entries
 */
import './style.css';
import * as XLSX from 'xlsx';
import { createGrid } from 'ag-grid-community';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';
import {
  loadData, getFilteredRawData, getFilteredIncidentData,
  setFilter, getUniqueValues, formatNumber,
  excelDateToJS, formatDate, formatDuration, getState
} from './dataStore.js';
import {
  renderTasksTrend, renderTaskTypePie,
  renderPriorityDist, renderBranchHeatmap, renderTaskTypeResolutionTargets,
  renderRepeatedSupportWindow, renderImmediateSupportWindow, renderRepeatedSupportByBranch, renderImmediateSupportByBranch,
  renderTaskKPIBars,
  renderIncidentTrend, renderIncidentCategory,
  renderIncidentComplexity, renderIncidentClosure,
  renderBranchPerformance, renderBranchClosureRates, renderBranchEfficiencyFromSheet, renderBranchWorkloadFromSheet,
  renderSameDayClosure, renderBranchTasksVolume,
  resizeCharts
} from './charts.js';

// ==========================================
// STATE
// ==========================================
let tasksGridApi = null;
let incidentsGridApi = null;
let currentTab = 'dashboard';
let dataLoaded = false;
let currentTheme = 'dark';

// ==========================================
// INIT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initSidebarCollapse();
  initNavigation();
  initFilters();
  initDateFilters();
  initUpload();
  initExport();
  initMenuToggle();

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
async function autoLoadExcel() {
  showToast('Loading data from Excel file...', 'info');

  try {
    const response = await fetch(`${import.meta.env.BASE_URL}data/kpi-data.xlsx`);
    if (!response.ok) {
      showToast('Place your Excel file in public/data/ folder or use Import', 'info');
      return;
    }
    const buf = await response.arrayBuffer();
    processExcelBuffer(buf);
  } catch (e) {
    console.log('No pre-loaded data file found. Use Import to load data.');
  }
}

function processExcelBuffer(buf) {
  const workbook = XLSX.read(buf, { type: 'array' });
  loadData(workbook, XLSX);

  dataLoaded = true;
  populateFilters();
  renderAll();
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
    support: 'Support KPI'
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
function initFilters() {
  ['filter-division', 'filter-region', 'filter-branch', 'filter-task-type'].forEach(id => {
    document.getElementById(id).addEventListener('change', (e) => {
      const key = id.replace('filter-', '').replace('-', '');
      const keyMap = {
        'division': 'division',
        'region': 'region',
        'branch': 'branch',
        'tasktype': 'taskType'
      };
      setFilter(keyMap[key], e.target.value);
      renderAll();
    });
  });
}

function initDateFilters() {
  const modeEl = document.getElementById('filter-date-mode');
  const anchorEl = document.getElementById('filter-date-anchor');
  const weekEl = document.getElementById('filter-date-week');
  const monthEl = document.getElementById('filter-date-month');
  const fromEl = document.getElementById('filter-date-from');
  const toEl = document.getElementById('filter-date-to');
  const separatorEl = document.getElementById('date-separator');

  if (!modeEl || !anchorEl || !weekEl || !monthEl || !fromEl || !toEl || !separatorEl) return;

  const today = new Date();
  const todayText = [
    today.getFullYear(),
    String(today.getMonth() + 1).padStart(2, '0'),
    String(today.getDate()).padStart(2, '0')
  ].join('-');
  const monthStartText = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;

  modeEl.value = 'monthly';
  anchorEl.value = todayText;
  weekEl.value = toISOWeek(today);
  monthEl.value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  setFilter('dateMode', modeEl.value);
  setFilter('dateAnchor', monthStartText);

  const syncDateInputs = () => {
    const mode = modeEl.value;
    const showAnchor = mode === 'daily';
    const showWeek = mode === 'weekly';
    const showMonth = mode === 'monthly';
    const showRange = mode === 'custom';

    anchorEl.classList.toggle('visible', showAnchor);
    weekEl.classList.toggle('visible', showWeek);
    monthEl.classList.toggle('visible', showMonth);
    fromEl.classList.toggle('visible', showRange);
    toEl.classList.toggle('visible', showRange);
    separatorEl.classList.toggle('visible', showRange);
  };

  syncDateInputs();

  modeEl.addEventListener('change', (e) => {
    const mode = e.target.value;
    setFilter('dateMode', mode);
    if (mode === 'daily') {
      setFilter('dateAnchor', anchorEl.value);
    } else if (mode === 'weekly') {
      setFilter('dateAnchor', isoWeekToDate(weekEl.value));
    } else if (mode === 'monthly') {
      setFilter('dateAnchor', monthToDate(monthEl.value));
    }
    syncDateInputs();
    renderAll();
  });

  anchorEl.addEventListener('change', (e) => {
    setFilter('dateAnchor', e.target.value);
    renderAll();
  });

  weekEl.addEventListener('change', (e) => {
    setFilter('dateAnchor', isoWeekToDate(e.target.value));
    renderAll();
  });

  monthEl.addEventListener('change', (e) => {
    setFilter('dateAnchor', monthToDate(e.target.value));
    renderAll();
  });

  fromEl.addEventListener('change', (e) => {
    setFilter('dateFrom', e.target.value);
    renderAll();
  });

  toEl.addEventListener('change', (e) => {
    setFilter('dateTo', e.target.value);
    renderAll();
  });
}

function toISOWeek(date) {
  const dt = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = dt.getUTCDay() || 7;
  dt.setUTCDate(dt.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((dt - yearStart) / 86400000) + 1) / 7);
  return `${dt.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function isoWeekToDate(weekValue) {
  const match = /^(\d{4})-W(\d{2})$/.exec(weekValue || '');
  if (!match) return '';
  const year = Number(match[1]);
  const week = Number(match[2]);
  const jan4 = new Date(year, 0, 4);
  const jan4Day = jan4.getDay() || 7;
  const mondayWeek1 = new Date(year, 0, 4 - (jan4Day - 1));
  const mondayTarget = new Date(mondayWeek1);
  mondayTarget.setDate(mondayWeek1.getDate() + (week - 1) * 7);
  return `${mondayTarget.getFullYear()}-${String(mondayTarget.getMonth() + 1).padStart(2, '0')}-${String(mondayTarget.getDate()).padStart(2, '0')}`;
}

function monthToDate(monthValue) {
  const match = /^(\d{4})-(\d{2})$/.exec(monthValue || '');
  if (!match) return '';
  return `${match[1]}-${match[2]}-01`;
}

function populateFilters() {
  populateSelect('filter-division', getUniqueValues('Division'), 'All Divisions');
  populateSelect('filter-region', getUniqueValues('Region'), 'All Regions');
  populateSelect('filter-branch', getUniqueValues('Branch'), 'All Branches');
  populateSelect('filter-task-type', getUniqueValues('Task Type'), 'All Task Types');
}

function populateSelect(id, values, defaultText) {
  const el = document.getElementById(id);
  el.innerHTML = `<option value="">${defaultText}</option>`;
  values.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    el.appendChild(opt);
  });
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
      renderBranchHeatmap('chart-branch-heatmap');
      renderTaskKPIBars('chart-task-kpi-bars');
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
function renderKPICards() {
  const data = getFilteredRawData();
  const incidentData = getFilteredIncidentData();
  const container = document.getElementById('kpi-cards');

  const total = data.length;
  const sameDay = data.filter(r => r['Same day Closure'] === true).length;
  const sameDayRate = total > 0 ? Math.round((sameDay / total) * 100) : 0;
  const avgDuration = total > 0 ? (data.reduce((sum, r) => sum + (parseFloat(r['Duration']) || 0), 0) / total) : 0;
  const taskTypes = new Set(data.map(r => r['Task Type'])).size;
  const incidentTotal = incidentData.length;
  const incidentSameDay = incidentData.filter(r => r['Same day Closure'] === true).length;
  const incidentSameDayRate = incidentTotal > 0 ? Math.round((incidentSameDay / incidentTotal) * 100) : 0;
  const incidentAvgDuration = incidentTotal > 0
    ? (incidentData.reduce((sum, r) => sum + (parseFloat(r['Duration']) || 0), 0) / incidentTotal)
    : 0;

  container.innerHTML = `
    <div class="kpi-card blue">
      <div class="kpi-label">Total Tasks</div>
      <div class="kpi-value">${formatNumber(total)}</div>
      <div class="kpi-change">${taskTypes} task types</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">Same-Day Closure</div>
      <div class="kpi-value">${sameDayRate}%</div>
      <div class="kpi-change up">${formatNumber(sameDay)} of ${formatNumber(total)}</div>
    </div>
    <div class="kpi-card amber">
      <div class="kpi-label">Avg Duration</div>
      <div class="kpi-value">${formatDuration(avgDuration)}</div>
      <div class="kpi-change">per task</div>
    </div>
    <div class="kpi-card purple">
      <div class="kpi-label">Incidents</div>
      <div class="kpi-value">${formatNumber(getFilteredIncidentData().length)}</div>
      <div class="kpi-change">fiber network</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">Incident Same-Day Closure</div>
      <div class="kpi-value">${incidentSameDayRate}%</div>
      <div class="kpi-change up">${formatNumber(incidentSameDay)} of ${formatNumber(incidentTotal)}</div>
    </div>
    <div class="kpi-card amber">
      <div class="kpi-label">Incident Avg Duration</div>
      <div class="kpi-value">${formatDuration(incidentAvgDuration)}</div>
      <div class="kpi-change">per incident</div>
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
    <div class="kpi-card green">
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
      valueFormatter: p => formatDate(excelDateToJS(p.value))
    },
    { field: 'Task Assigned', width: 120,
      valueFormatter: p => formatDate(excelDateToJS(p.value))
    },
    { field: 'Task Completed', width: 120,
      valueFormatter: p => formatDate(excelDateToJS(p.value))
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
