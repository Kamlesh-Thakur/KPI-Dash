import { pool } from '../db/pool.js';
import { requireAuth, requireRoles } from '../middleware/auth.js';

function applyDimensionFilters(rows, filters, incident = false) {
  const clusterOf = (r) => r['Cluster-1'] || r['Cluster'] || r[' Cluster'];
  return rows.filter((r) => {
    if (filters.division && r.Division !== filters.division) return false;
    if (filters.region && r.Region !== filters.region) return false;
    if (filters.cluster && clusterOf(r) !== filters.cluster) return false;
    if (filters.branch && r.Branch !== filters.branch) return false;
    if (!incident && filters.taskType && r['Task Type'] !== filters.taskType) return false;
    return true;
  });
}

function uniqueValues(rows, keyFn) {
  return [...new Set(rows.map((r) => keyFn(r)).filter(Boolean))].sort();
}

export async function kpiRoutes(fastify) {
  const guard = [requireAuth, requireRoles('admin', 'analyst', 'viewer')];

  fastify.get('/kpi/rows', { preHandler: guard }, async (request) => {
    const { dataset = 'raw_data' } = request.query || {};
    const res = await pool.query(
      `SELECT row_payload FROM kpi_rows_curated WHERE dataset = $1 ORDER BY id`,
      [dataset]
    );
    const rows = res.rows.map((r) => r.row_payload);
    return { dataset, rows };
  });

  fastify.get('/kpi/filters', { preHandler: guard }, async (request) => {
    const filters = {
      division: String(request.query?.division || ''),
      region: String(request.query?.region || ''),
      cluster: String(request.query?.cluster || ''),
      branch: String(request.query?.branch || ''),
      taskType: String(request.query?.taskType || '')
    };
    const [rawRes] = await Promise.all([
      pool.query(`SELECT row_payload FROM kpi_rows_curated WHERE dataset = 'raw_data' ORDER BY id`)
    ]);
    const rawRows = rawRes.rows.map((r) => r.row_payload);
    const rowsByDivision = rawRows.filter((r) => !filters.division || r.Division === filters.division);
    const rowsByRegion = rowsByDivision.filter((r) => !filters.region || r.Region === filters.region);
    const clusterOf = (r) => r['Cluster-1'] || r['Cluster'] || r[' Cluster'];
    const rowsByCluster = rowsByRegion.filter((r) => !filters.cluster || clusterOf(r) === filters.cluster);

    return {
      divisions: uniqueValues(rawRows, (r) => r.Division),
      regions: uniqueValues(rowsByDivision, (r) => r.Region),
      clusters: uniqueValues(rowsByRegion, (r) => clusterOf(r)),
      branches: uniqueValues(rowsByCluster, (r) => r.Branch),
      taskTypes: uniqueValues(rawRows, (r) => r['Task Type'])
    };
  });

  fastify.get('/kpi/summary', { preHandler: guard }, async (request) => {
    const filters = {
      division: String(request.query?.division || ''),
      region: String(request.query?.region || ''),
      cluster: String(request.query?.cluster || ''),
      branch: String(request.query?.branch || ''),
      taskType: String(request.query?.taskType || '')
    };
    const [rawRes, incRes] = await Promise.all([
      pool.query(`SELECT row_payload FROM kpi_rows_curated WHERE dataset = 'raw_data'`),
      pool.query(`SELECT row_payload FROM kpi_rows_curated WHERE dataset = 'incident'`)
    ]);
    const rawRows = applyDimensionFilters(rawRes.rows.map((r) => r.row_payload), filters, false);
    const incidentRows = applyDimensionFilters(incRes.rows.map((r) => r.row_payload), filters, true);
    return {
      totals: {
        tasks: rawRows.length,
        incidents: incidentRows.length
      }
    };
  });
}

