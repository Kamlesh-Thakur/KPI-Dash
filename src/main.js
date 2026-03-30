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
  renderTasksTrend, renderTaskTypePie, renderRegionBar,
  renderPriorityDist, renderSLAGauge, renderBranchHeatmap,
  renderIncidentTrend, renderIncidentCategory,
  renderIncidentComplexity, renderIncidentClosure,
  renderBranchPerformance, renderDivisionCompare,
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

// ==========================================
// INIT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  initNavigation();
  initFilters();
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
    branch: 'Branch KPI'
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

// ==========================================
// RENDER ALL
// ==========================================
function renderAll() {
  renderKPICards();
  renderIncidentKPICards();
  renderTabContent(currentTab);
}

function renderTabContent(tab) {
  switch (tab) {
    case 'dashboard':
      renderTasksTrend('chart-tasks-trend');
      renderTaskTypePie('chart-task-type-pie');
      renderRegionBar('chart-region-bar');
      renderPriorityDist('chart-priority-dist');
      renderSLAGauge('chart-sla-gauge');
      renderBranchHeatmap('chart-branch-heatmap');
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
      renderDivisionCompare('chart-division-compare');
      renderSameDayClosure('chart-same-day-closure');
      renderBranchTasksVolume('chart-branch-tasks-volume');
      break;
  }
}

// ==========================================
// KPI CARDS
// ==========================================
function renderKPICards() {
  const data = getFilteredRawData();
  const container = document.getElementById('kpi-cards');

  const total = data.length;
  const sameDay = data.filter(r => r['Same day Closure'] === true).length;
  const sameDayRate = total > 0 ? Math.round((sameDay / total) * 100) : 0;
  const avgDuration = total > 0 ? (data.reduce((sum, r) => sum + (parseFloat(r['Duration']) || 0), 0) / total) : 0;
  const uniqueBranches = new Set(data.map(r => r['Branch'])).size;
  const taskTypes = new Set(data.map(r => r['Task Type'])).size;
  const escalated = data.filter(r => r['Priority Escalated'] === 'Y').length;

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
    <div class="kpi-card cyan">
      <div class="kpi-label">Active Branches</div>
      <div class="kpi-value">${uniqueBranches}</div>
      <div class="kpi-change">across all regions</div>
    </div>
    <div class="kpi-card red">
      <div class="kpi-label">Escalated</div>
      <div class="kpi-value">${formatNumber(escalated)}</div>
      <div class="kpi-change down">${total > 0 ? Math.round((escalated / total) * 100) : 0}% escalation rate</div>
    </div>
    <div class="kpi-card purple">
      <div class="kpi-label">Incidents</div>
      <div class="kpi-value">${formatNumber(getFilteredIncidentData().length)}</div>
      <div class="kpi-change">fiber network</div>
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
  gridDiv.className = 'ag-theme-quartz-dark ag-theme-custom';
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
  gridDiv.className = 'ag-theme-quartz-dark ag-theme-custom';
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
