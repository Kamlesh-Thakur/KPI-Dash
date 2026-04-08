/**
 * Data Store — holds all loaded data and provides filtered views.
 */

import { getCalendarSystem } from './calendarPrefs.js';
import { getBsMonthAdDateRange, parseBsMonthCode } from './nepaliCalendarUi.js';

const STATE = {
  rawData: [],
  incidentData: [],
  /** Snapshot of the last loaded workbook's "Overall KPI" sheet (values; formulas not evaluated in-browser). */
  overallKpi: null,
  branchEfficiency: [],
  teamPerformance: {
    fsNew: [],
    rsAre: []
  },
  branchMapping: [],
  dropdowns: [],
  filters: {
    filterBy: '',
    division: '',
    region: '',
    cluster: '',
    branch: '',
    taskType: '',
    dateMode: 'monthly',
    dateAnchor: '',
    dateFrom: '',
    dateTo: '',
    /** Bikram Sambat month code `YYYY-MM` (month index 00–11). Used for monthly range when Nepali calendar is active. */
    bsMonthCode: ''
  }
};

const FILTERED_CACHE = {
  key: '',
  raw: null,
  incident: null
};

function invalidateFilteredCache() {
  FILTERED_CACHE.key = '';
  FILTERED_CACHE.raw = null;
  FILTERED_CACHE.incident = null;
}

function buildFilteredCacheKey(filters, startDate, endDate) {
  const f = filters || {};
  const start = startDate instanceof Date ? startDate.toISOString().slice(0, 10) : '';
  const end = endDate instanceof Date ? endDate.toISOString().slice(0, 10) : '';
  return [
    f.filterBy || '',
    f.division || '',
    f.region || '',
    f.cluster || '',
    f.branch || '',
    f.taskType || '',
    f.dateMode || 'overall',
    f.dateAnchor || '',
    f.dateFrom || '',
    f.dateTo || '',
    f.bsMonthCode || '',
    start,
    end,
    // Include data lengths so cache invalidates after new workbook load.
    STATE.rawData.length,
    STATE.incidentData.length
  ].join('|');
}

function ensureFilteredCache(filters = STATE.filters) {
  const ranged = getDateFilterRange(filters);
  const key = buildFilteredCacheKey(filters, ranged.start, ranged.end);
  if (FILTERED_CACHE.key === key && FILTERED_CACHE.raw && FILTERED_CACHE.incident) {
    return FILTERED_CACHE;
  }
  FILTERED_CACHE.key = key;
  FILTERED_CACHE.raw = getFilteredRawDataForRange(ranged.start, ranged.end, filters);
  FILTERED_CACHE.incident = getFilteredIncidentDataForRange(ranged.start, ranged.end, filters);
  return FILTERED_CACHE;
}

/** Parse Excel serial date to JS Date */
export function excelDateToJS(serial) {
  if (!serial || typeof serial !== 'number') return null;
  // Excel uses 1900 epoch, JS uses 1970
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400 * 1000;
  return new Date(utcValue);
}

/** Format date as YYYY-MM-DD */
export function formatDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return '—';
  return d.toISOString().slice(0, 10);
}

/** Format date as shorter display */
export function formatDateShort(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

/** Format number with commas */
export function formatNumber(n) {
  if (n == null || isNaN(n)) return '—';
  return Number(n).toLocaleString('en-US');
}

/** Format duration in hours */
export function formatDuration(days) {
  if (days == null || isNaN(days)) return '—';
  const hrs = Math.round(days * 24 * 10) / 10;
  return `${hrs}h`;
}

/** Clean column names (trim leading spaces) */
function cleanKey(key) {
  return key.trim();
}

function cleanRow(row) {
  const cleaned = {};
  for (const [key, val] of Object.entries(row)) {
    if (key.startsWith('__EMPTY')) continue;
    cleaned[cleanKey(key)] = val;
  }
  return cleaned;
}

function toDateOnly(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/** Calendar week Sunday–Saturday (local dates), US convention. Anchor may be any day in that week. */
function weekRangeSunToSat(anchor) {
  const start = new Date(anchor);
  const dow = start.getDay(); // 0 = Sunday … 6 = Saturday
  start.setDate(start.getDate() - dow);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return { start, end };
}

function parseDateInput(value) {
  if (!value) return null;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  return toDateOnly(parsed);
}

function getRowDate(row) {
  const dateFields = ['Completed date', 'Task Completed', 'Task Assigned', 'Task Created'];
  for (const field of dateFields) {
    const value = row[field];
    if (typeof value === 'number') {
      const d = excelDateToJS(value);
      const dateOnly = toDateOnly(d);
      if (dateOnly) return dateOnly;
    } else if (value instanceof Date) {
      const dateOnly = toDateOnly(value);
      if (dateOnly) return dateOnly;
    } else if (typeof value === 'string' && value.trim() !== '') {
      const parsed = parseDateInput(value);
      if (parsed) return parsed;
    }
  }
  return null;
}

function isWithinDateFilter(rowDate, filters) {
  const mode = filters.dateMode || 'overall';
  if (mode === 'overall') return true;
  if (!rowDate) return false;

  if (mode === 'custom') {
    const from = parseDateInput(filters.dateFrom);
    const to = parseDateInput(filters.dateTo);
    if (!from || !to) return true;
    return rowDate >= from && rowDate <= to;
  }

  const anchor = parseDateInput(filters.dateAnchor) || toDateOnly(new Date());
  if (!anchor) return true;

  if (mode === 'daily') {
    return rowDate.getTime() === anchor.getTime();
  }

  if (mode === 'weekly') {
    const { start, end } = weekRangeSunToSat(anchor);
    return rowDate >= start && rowDate <= end;
  }

  if (mode === 'monthly') {
    return rowDate.getFullYear() === anchor.getFullYear() && rowDate.getMonth() === anchor.getMonth();
  }

  return true;
}

/**
 * Load data from parsed XLSX sheets.
 * @param {object} workbook - Parsed XLSX workbook
 * @param {object} XLSX - xlsx library
 * @param {{ append?: boolean }} [options] - If append is true, Raw Data and Incident rows are concatenated;
 *   Branch Effi. / Sheet2 / Drop Down are re-applied from this workbook when present (last file wins).
 */
export function loadData(workbook, XLSX, options = {}) {
  const append = options.append === true;

  // Raw Data sheet
  const rawSheet = workbook.Sheets['Raw Data'];
  if (rawSheet) {
    const json = XLSX.utils.sheet_to_json(rawSheet, { defval: '' });
    const rows = json.map(r => cleanRow(r));
    STATE.rawData = append ? [...STATE.rawData, ...rows] : rows;
  } else if (!append) {
    STATE.rawData = [];
  }

  // Incident sheet
  const incSheet = workbook.Sheets['Incident'];
  if (incSheet) {
    const json = XLSX.utils.sheet_to_json(incSheet, { defval: '' });
    const rows = json.map(r => cleanRow(r));
    if (append && STATE.incidentData.length) {
      // Append and de‑duplicate by Incident ID (and Completed date if missing ID)
      const merged = [...STATE.incidentData, ...rows];
      const byKey = new Map();
      for (const r of merged) {
        const id = (r['Incident ID'] || '').toString().trim();
        const completed = r['Completed date'] ?? r['Task Completed'] ?? '';
        const key = id || `${completed}::${JSON.stringify(r)}`;
        // Last file wins for duplicates
        byKey.set(key, r);
      }
      STATE.incidentData = [...byKey.values()];
    } else {
      STATE.incidentData = rows;
    }
  } else if (!append) {
    STATE.incidentData = [];
  }

  // Branch Efficiency sheet
  const branchEffSheet = workbook.Sheets['Branch Effi.'];
  if (branchEffSheet) {
    const rows = XLSX.utils.sheet_to_json(branchEffSheet, { header: 1, defval: '' });
    const parsedRows = rows
      .map((row) => {
        const branch = (row[1] || '').toString().trim().replace(/\s+/g, ' ');
        const efficiency = Number(row[2]);
        const efficiencyPerWorkingDay = Number(row[3]);
        const workload = Number(row[27]);
        if (!branch || branch.toLowerCase() === 'branch' || branch.toLowerCase() === 'grand total' || Number.isNaN(efficiency)) return null;
        return {
          branch,
          efficiency,
          efficiencyPerWorkingDay: Number.isNaN(efficiencyPerWorkingDay) ? null : efficiencyPerWorkingDay,
          workload: Number.isNaN(workload) ? null : workload
        };
      })
      .filter(Boolean);
    // Keep one row per branch name if sheet contains repeated summary sections.
    const byBranch = new Map();
    parsedRows.forEach((row) => byBranch.set(row.branch, row));
    STATE.branchEfficiency = [...byBranch.values()];
  }

  // Sheet2 — branch mappings
  const mapSheet = workbook.Sheets['Sheet2'];
  if (mapSheet) {
    const json = XLSX.utils.sheet_to_json(mapSheet, { defval: '' });
    STATE.branchMapping = json.map(r => cleanRow(r));
  }

  // Drop Down
  const ddSheet = workbook.Sheets['Drop Down'];
  if (ddSheet) {
    const json = XLSX.utils.sheet_to_json(ddSheet, { defval: '' });
    STATE.dropdowns = json.map(r => cleanRow(r));
  }

  // Overall KPI — summary sheet (same layout as Excel workbook; optional)
  const overallSheet = workbook.Sheets['Overall KPI'];
  if (overallSheet) {
    STATE.overallKpi = {
      rows: XLSX.utils.sheet_to_json(overallSheet, { defval: '', raw: true }),
      ref: overallSheet['!ref'] || null
    };
  } else if (!append) {
    STATE.overallKpi = null;
  }

  invalidateFilteredCache();
  return STATE;
}

/**
 * Load precomputed monthly snapshot (generated offline from source workbooks).
 * Keeps the same in-memory shape used by the app so filters/charts continue to run on-demand.
 */
export function loadPrecomputedData(snapshot = {}) {
  STATE.rawData = Array.isArray(snapshot.rawData) ? snapshot.rawData : [];
  STATE.incidentData = Array.isArray(snapshot.incidentData) ? snapshot.incidentData : [];
  STATE.branchEfficiency = Array.isArray(snapshot.branchEfficiency) ? snapshot.branchEfficiency : [];
  STATE.branchMapping = Array.isArray(snapshot.branchMapping) ? snapshot.branchMapping : [];
  STATE.dropdowns = Array.isArray(snapshot.dropdowns) ? snapshot.dropdowns : [];
  STATE.overallKpi = snapshot.overallKpi || null;
  invalidateFilteredCache();
  return STATE;
}

/**
 * Hours to resolve (matches Excel Incident column AD = (Closed At − First Incident At) × 24).
 * SheetJS duplicates the header "Duration": computed hours are in `Duration_1`; the first `Duration`
 * column is a display string — do not use it for KPI math.
 */
export function incidentDurationHours(r) {
  const d1 = parseFloat(r['Duration_1']);
  if (!Number.isNaN(d1)) return d1;
  const closed = r['Closed At'];
  const first = r['First Incident At'];
  if (typeof closed === 'number' && typeof first === 'number' && !Number.isNaN(closed - first)) {
    return (closed - first) * 24;
  }
  return null;
}

/** Same-day flag: Excel uses the literal "true" on Incident `Same day Closure`. */
export function isSameDayClosure(r) {
  const v = r['Same day Closure'];
  return v === true || String(v).toLowerCase() === 'true';
}

/**
 * Incident is excluded from the "KPI after exclusion" pool when Exceptions indicates Delayed
 * and resolution duration (hours) is strictly greater than 24.
 */
export function isIncidentExcludedFromKpi(r) {
  const exc = (r['Exceptions'] || '').toString().toLowerCase();
  if (!exc.includes('delay')) return false;
  const h = incidentDurationHours(r);
  if (h == null) return false;
  return h > 24;
}

export function getOverallKpiSnapshot() {
  return STATE.overallKpi;
}

export function loadTeamPerformanceData(workbook, XLSX) {
  const fsSheet = workbook.Sheets['FS and New'];
  if (fsSheet) {
    const rows = XLSX.utils.sheet_to_json(fsSheet, { header: 1, defval: '' });
    STATE.teamPerformance.fsNew = rows
      .slice(4)
      .map((row) => {
        const rank = Number(row[0]);
        const branch = (row[1] || '').toString().trim();
        const region = (row[3] || '').toString().trim();
        const agent = (row[4] || '').toString().trim();
        const teamType = (row[5] || '').toString().trim();
        const score = Number(row[6]);
        if (Number.isNaN(rank) || !branch || !agent || Number.isNaN(score)) return null;
        return { rank, branch, region, agent, teamType, score };
      })
      .filter(Boolean);
  }

  const rsSheet = workbook.Sheets[' RS and ARE'];
  if (rsSheet) {
    const rows = XLSX.utils.sheet_to_json(rsSheet, { header: 1, defval: '' });
    STATE.teamPerformance.rsAre = rows
      .slice(4)
      .map((row) => {
        const rank = Number(row[0]);
        const branch = (row[1] || '').toString().trim();
        const region = (row[3] || '').toString().trim();
        const agent = (row[4] || '').toString().trim();
        const team = (row[5] || '').toString().trim();
        const workingDaysPct = Number(row[6]);
        const taskHandledPct = Number(row[8]);
        const avgTasksHandled = Number(row[9]);
        if (Number.isNaN(rank) || !branch || !agent) return null;
        return {
          rank,
          branch,
          region,
          agent,
          team,
          workingDaysPct: Number.isNaN(workingDaysPct) ? null : workingDaysPct,
          taskHandledPct: Number.isNaN(taskHandledPct) ? null : taskHandledPct,
          avgTasksHandled: Number.isNaN(avgTasksHandled) ? null : avgTasksHandled
        };
      })
      .filter(Boolean);
  }
}

/** Get filtered raw data */
export function getFilteredRawData() {
  return ensureFilteredCache(STATE.filters).raw || [];
}

/** Get filtered incident data */
export function getFilteredIncidentData() {
  return ensureFilteredCache(STATE.filters).incident || [];
}

/** Set a filter value */
export function setFilter(key, value) {
  if (STATE.filters[key] === value) return;
  STATE.filters[key] = value;
  invalidateFilteredCache();
}

/** Get unique values from a column */
export function getUniqueValues(column, dataset = 'raw') {
  const data = dataset === 'raw' ? STATE.rawData : STATE.incidentData;
  const values = new Set();
  data.forEach(r => {
    const v = r[column];
    if (v && v !== '' && v !== 'N/A') values.add(v);
  });
  return [...values].sort();
}

/** Get all state */
export function getState() {
  return STATE;
}

export function getBranchEfficiencyData() {
  return STATE.branchEfficiency || [];
}

export function getTeamPerformanceData() {
  return STATE.teamPerformance || { fsNew: [], rsAre: [] };
}

export function getDateFilterRange(filters = STATE.filters) {
  const mode = filters.dateMode || 'overall';
  const today = toDateOnly(new Date());
  if (!today) return { start: null, end: null };

  if (mode === 'custom') {
    const from = parseDateInput(filters.dateFrom);
    const to = parseDateInput(filters.dateTo);
    return { start: from, end: to };
  }

  const anchor = parseDateInput(filters.dateAnchor) || today;
  if (mode === 'daily') return { start: anchor, end: anchor };

  if (mode === 'weekly') {
    return weekRangeSunToSat(anchor);
  }

  if (mode === 'monthly') {
    const bsCode = filters.bsMonthCode;
    const useBsMonth =
      getCalendarSystem() === 'nepali' &&
      bsCode &&
      parseBsMonthCode(bsCode);
    if (useBsMonth) {
      const bsRange = getBsMonthAdDateRange(bsCode);
      if (bsRange) return bsRange;
    }
    const start = new Date(anchor.getFullYear(), anchor.getMonth(), 1);
    const end = new Date(anchor.getFullYear(), anchor.getMonth() + 1, 0);
    return { start, end };
  }

  // overall: derive full available range
  const allDates = [...STATE.rawData, ...STATE.incidentData]
    .map(getRowDate)
    .filter(Boolean)
    .sort((a, b) => a - b);
  if (!allDates.length) return { start: null, end: null };
  return { start: allDates[0], end: allDates[allDates.length - 1] };
}

export function getFilteredRawDataForRange(startDate, endDate, filters = STATE.filters) {
  let data = applyRawDimensionFilters(STATE.rawData, filters);
  if (!startDate || !endDate) return data;
  return data.filter((r) => {
    const d = getRowDate(r);
    return d && d >= startDate && d <= endDate;
  });
}

export function getFilteredIncidentDataForRange(startDate, endDate, filters = STATE.filters) {
  let data = applyIncidentDimensionFilters(STATE.incidentData, filters);
  if (!startDate || !endDate) return data;
  return data.filter((r) => {
    const d = getRowDate(r);
    return d && d >= startDate && d <= endDate;
  });
}

function applyRawDimensionFilters(data, filters) {
  let out = data;
  const fb = filters.filterBy || '';
  if (fb === 'division' && filters.division) {
    out = out.filter((r) => r['Division'] === filters.division);
  } else if (fb === 'region' && filters.region) {
    out = out.filter((r) => r['Region'] === filters.region);
  } else if (fb === 'cluster' && filters.cluster) {
    out = out.filter((r) => (r['Cluster-1'] || r[' Cluster'] || r['Cluster']) === filters.cluster);
  } else if (fb === 'branch' && filters.branch) {
    out = out.filter((r) => r['Branch'] === filters.branch);
  } else if (fb === 'taskType' && filters.taskType) {
    out = out.filter((r) => r['Task Type'] === filters.taskType);
  }
  if (filters.division) out = out.filter((r) => r['Division'] === filters.division);
  if (filters.region) out = out.filter((r) => r['Region'] === filters.region);
  if (filters.cluster) out = out.filter((r) => (r['Cluster-1'] || r[' Cluster'] || r['Cluster']) === filters.cluster);
  if (filters.branch) out = out.filter((r) => r['Branch'] === filters.branch);
  if (filters.taskType) out = out.filter((r) => r['Task Type'] === filters.taskType);
  return out;
}

function applyIncidentDimensionFilters(data, filters) {
  let out = data;
  const fb = filters.filterBy || '';
  if (fb === 'division' && filters.division) {
    out = out.filter((r) => r['Division'] === filters.division);
  } else if (fb === 'region' && filters.region) {
    out = out.filter((r) => r['Region'] === filters.region);
  } else if (fb === 'cluster' && filters.cluster) {
    out = out.filter((r) => (r['Cluster-1'] || r[' Cluster'] || r['Cluster']) === filters.cluster);
  } else if (fb === 'branch' && filters.branch) {
    out = out.filter((r) => (r['Branch_1'] || r['Branch']) === filters.branch);
  }
  if (filters.division) out = out.filter((r) => r['Division'] === filters.division);
  if (filters.region) out = out.filter((r) => r['Region'] === filters.region);
  if (filters.cluster) out = out.filter((r) => (r['Cluster-1'] || r[' Cluster'] || r['Cluster']) === filters.cluster);
  if (filters.branch) out = out.filter((r) => (r['Branch_1'] || r['Branch']) === filters.branch);
  return out;
}

export default STATE;
