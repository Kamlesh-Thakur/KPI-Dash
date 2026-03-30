/**
 * Data Store — holds all loaded data and provides filtered views.
 */

const STATE = {
  rawData: [],
  incidentData: [],
  branchMapping: [],
  dropdowns: [],
  filters: {
    division: '',
    region: '',
    branch: '',
    taskType: ''
  }
};

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

/** Load data from parsed XLSX sheets */
export function loadData(workbook, XLSX) {
  // Raw Data sheet
  const rawSheet = workbook.Sheets['Raw Data'];
  if (rawSheet) {
    const json = XLSX.utils.sheet_to_json(rawSheet, { defval: '' });
    STATE.rawData = json.map(r => cleanRow(r));
  }

  // Incident sheet
  const incSheet = workbook.Sheets['Incident'];
  if (incSheet) {
    const json = XLSX.utils.sheet_to_json(incSheet, { defval: '' });
    STATE.incidentData = json.map(r => cleanRow(r));
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

  return STATE;
}

/** Get filtered raw data */
export function getFilteredRawData() {
  let data = STATE.rawData;
  const f = STATE.filters;

  if (f.division) {
    data = data.filter(r => r['Division'] === f.division);
  }
  if (f.region) {
    data = data.filter(r => r['Region'] === f.region);
  }
  if (f.branch) {
    data = data.filter(r => r['Branch'] === f.branch);
  }
  if (f.taskType) {
    data = data.filter(r => r['Task Type'] === f.taskType);
  }
  return data;
}

/** Get filtered incident data */
export function getFilteredIncidentData() {
  let data = STATE.incidentData;
  const f = STATE.filters;

  if (f.division) {
    data = data.filter(r => r['Division'] === f.division);
  }
  if (f.region) {
    data = data.filter(r => r['Region'] === f.region);
  }
  if (f.branch) {
    data = data.filter(r => (r['Branch_1'] || r['Branch']) === f.branch);
  }
  return data;
}

/** Set a filter value */
export function setFilter(key, value) {
  STATE.filters[key] = value;
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

export default STATE;
