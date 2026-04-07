import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';

const ROOT = process.cwd();
const DATA_DIR = path.join(ROOT, 'public', 'data');
const KPI_FILES = ['kpi-falgun-2082.xlsx', 'kpi-chaitra-2082.xlsx'];
const OUT_FILE = path.join(DATA_DIR, 'precomputed-kpi.json');

function cleanRow(row) {
  const out = {};
  for (const [k, v] of Object.entries(row || {})) {
    if (String(k).startsWith('__EMPTY')) continue;
    out[String(k).trim()] = v;
  }
  return out;
}

function parseBranchEfficiency(sheet, xlsx) {
  if (!sheet) return [];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const parsed = rows
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
  const byBranch = new Map();
  parsed.forEach((row) => byBranch.set(row.branch, row));
  return [...byBranch.values()];
}

function dedupeIncidentRows(rows) {
  const byKey = new Map();
  for (const r of rows) {
    const id = (r['Incident ID'] || '').toString().trim();
    const completed = r['Completed date'] ?? r['Task Completed'] ?? '';
    const key = id || `${completed}::${JSON.stringify(r)}`;
    byKey.set(key, r);
  }
  return [...byKey.values()];
}

function main() {
  const snapshot = {
    generatedAt: new Date().toISOString(),
    sourceFiles: KPI_FILES.slice(),
    sourceMeta: {},
    rawData: [],
    incidentData: [],
    branchEfficiency: [],
    branchMapping: [],
    dropdowns: [],
    overallKpi: null
  };

  for (let i = 0; i < KPI_FILES.length; i++) {
    const fileName = KPI_FILES[i];
    const filePath = path.join(DATA_DIR, fileName);
    if (!fs.existsSync(filePath)) {
      throw new Error(`Missing source workbook: ${filePath}`);
    }
    const stat = fs.statSync(filePath);
    snapshot.sourceMeta[fileName] = {
      size: stat.size,
      mtimeMs: stat.mtimeMs
    };
    const wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });

    const raw = wb.Sheets['Raw Data'];
    if (raw) {
      const rows = XLSX.utils.sheet_to_json(raw, { defval: '' }).map(cleanRow);
      snapshot.rawData = i > 0 ? snapshot.rawData.concat(rows) : rows;
    } else if (i === 0) {
      snapshot.rawData = [];
    }

    const inc = wb.Sheets['Incident'];
    if (inc) {
      const rows = XLSX.utils.sheet_to_json(inc, { defval: '' }).map(cleanRow);
      snapshot.incidentData = i > 0
        ? dedupeIncidentRows(snapshot.incidentData.concat(rows))
        : rows;
    } else if (i === 0) {
      snapshot.incidentData = [];
    }

    const branchEff = wb.Sheets['Branch Effi.'];
    if (branchEff) snapshot.branchEfficiency = parseBranchEfficiency(branchEff, XLSX);

    const map = wb.Sheets['Sheet2'];
    if (map) snapshot.branchMapping = XLSX.utils.sheet_to_json(map, { defval: '' }).map(cleanRow);

    const dd = wb.Sheets['Drop Down'];
    if (dd) snapshot.dropdowns = XLSX.utils.sheet_to_json(dd, { defval: '' }).map(cleanRow);

    const overall = wb.Sheets['Overall KPI'];
    if (overall) {
      snapshot.overallKpi = {
        rows: XLSX.utils.sheet_to_json(overall, { defval: '', raw: true }),
        ref: overall['!ref'] || null
      };
    }
  }

  fs.writeFileSync(OUT_FILE, JSON.stringify(snapshot));
  console.log(`Wrote ${OUT_FILE}`);
  console.log(`Tasks: ${snapshot.rawData.length}, Incidents: ${snapshot.incidentData.length}`);
}

main();
