import crypto from 'node:crypto';

function normalizeValue(v) {
  if (v == null) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  return String(v).trim();
}

export function normalizeRow(row) {
  const out = {};
  for (const [k, v] of Object.entries(row || {})) {
    if (k.startsWith('__EMPTY')) continue;
    out[String(k).trim()] = normalizeValue(v);
  }
  return out;
}

function makeHash(payload) {
  return crypto.createHash('sha256').update(payload).digest('hex');
}

function getIncidentKey(row) {
  const id = normalizeValue(row['Incident ID']);
  if (id) return `incident:${id}`;
  const composite = [
    normalizeValue(row['First Incident At']),
    normalizeValue(row['Closed At']),
    normalizeValue(row['Branch']),
    normalizeValue(row['Category']),
    normalizeValue(row['Task Type'])
  ].join('|');
  return `incident:${composite}`;
}

function getTaskKey(row) {
  const id = normalizeValue(row['Ticket ID']) || normalizeValue(row['Task ID']) || normalizeValue(row['Reference']);
  if (id) return `task:${id}`;
  const composite = [
    normalizeValue(row['Task Created']),
    normalizeValue(row['Task Completed']),
    normalizeValue(row['Branch']),
    normalizeValue(row['Task Type']),
    normalizeValue(row['Category'])
  ].join('|');
  return `task:${composite}`;
}

export function buildRowIdentity(dataset, row) {
  const normalized = normalizeRow(row);
  const rowKey = dataset === 'incident' ? getIncidentKey(normalized) : getTaskKey(normalized);
  const rowHash = makeHash(JSON.stringify(normalized));
  return { normalized, rowKey, rowHash };
}

export function fileSha256(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

