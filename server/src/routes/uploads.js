import * as XLSX from 'xlsx';
import { requireAuth, requireRoles } from '../middleware/auth.js';
import { tx } from '../db/pool.js';
import { buildRowIdentity, fileSha256 } from '../services/dedup.js';
import { logAudit } from '../services/auditService.js';

function cleanRow(row) {
  const out = {};
  for (const [k, v] of Object.entries(row || {})) {
    if (String(k).startsWith('__EMPTY')) continue;
    out[String(k).trim()] = v;
  }
  return out;
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const rawSheet = workbook.Sheets['Raw Data'];
  const incidentSheet = workbook.Sheets['Incident'];
  const rawRows = rawSheet ? XLSX.utils.sheet_to_json(rawSheet, { defval: '' }).map(cleanRow) : [];
  const incidentRows = incidentSheet ? XLSX.utils.sheet_to_json(incidentSheet, { defval: '' }).map(cleanRow) : [];
  return { rawRows, incidentRows };
}

async function upsertDataset(client, manifestId, dataset, rows) {
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (let i = 0; i < rows.length; i += 1) {
    const sourceRowNo = i + 2;
    const row = rows[i];
    await client.query(
      `INSERT INTO kpi_rows_raw(manifest_id, dataset, source_row_no, row_payload)
       VALUES ($1, $2, $3, $4::jsonb)`,
      [manifestId, dataset, sourceRowNo, JSON.stringify(row)]
    );

    const identity = buildRowIdentity(dataset === 'incident' ? 'incident' : 'task', row);
    const existing = await client.query(
      `SELECT id, row_hash FROM kpi_rows_curated WHERE dataset = $1 AND row_key = $2`,
      [dataset, identity.rowKey]
    );
    if (!existing.rowCount) {
      await client.query(
        `INSERT INTO kpi_rows_curated(dataset, row_key, row_hash, row_payload, latest_manifest_id)
         VALUES ($1, $2, $3, $4::jsonb, $5)`,
        [dataset, identity.rowKey, identity.rowHash, JSON.stringify(identity.normalized), manifestId]
      );
      inserted += 1;
      continue;
    }
    const current = existing.rows[0];
    if (current.row_hash === identity.rowHash) {
      skipped += 1;
      continue;
    }
    await client.query(
      `UPDATE kpi_rows_curated
       SET row_hash = $1, row_payload = $2::jsonb, latest_manifest_id = $3, updated_at = NOW()
       WHERE id = $4`,
      [identity.rowHash, JSON.stringify(identity.normalized), manifestId, current.id]
    );
    updated += 1;
  }
  return { inserted, updated, skipped, total: rows.length };
}

export async function uploadRoutes(fastify) {
  fastify.post('/uploads/workbook', {
    preHandler: [requireAuth, requireRoles('admin', 'analyst')]
  }, async (request, reply) => {
    const part = await request.file();
    if (!part) return reply.code(400).send({ error: 'Workbook file is required' });
    const buffer = await part.toBuffer();
    const sourceName = part.filename || 'upload.xlsx';
    const sourceSize = buffer.length;
    const sourceSha = fileSha256(buffer);
    const periodCode = String(request.query?.periodCode || request.body?.periodCode || '').trim() || null;
    const uploadMode = String(request.query?.uploadMode || request.body?.uploadMode || 'manual');

    const summary = await tx(async (client) => {
      const existing = await client.query(
        `SELECT id, summary FROM upload_manifests WHERE source_sha256 = $1`,
        [sourceSha]
      );
      if (existing.rowCount) {
        return { duplicateManifest: true, manifestId: existing.rows[0].id, ...existing.rows[0].summary };
      }

      const manifestRes = await client.query(
        `INSERT INTO upload_manifests(uploaded_by, source_name, source_sha256, source_size, period_code, upload_mode, status)
         VALUES ($1, $2, $3, $4, $5, $6, 'processing')
         RETURNING id`,
        [request.user.sub, sourceName, sourceSha, sourceSize, periodCode, uploadMode]
      );
      const manifestId = manifestRes.rows[0].id;
      const { rawRows, incidentRows } = parseWorkbook(buffer);
      const rawSummary = await upsertDataset(client, manifestId, 'raw_data', rawRows);
      const incSummary = await upsertDataset(client, manifestId, 'incident', incidentRows);
      const payload = {
        manifestId,
        duplicateManifest: false,
        raw: rawSummary,
        incident: incSummary
      };
      await client.query(
        `UPDATE upload_manifests SET status = 'completed', summary = $2::jsonb WHERE id = $1`,
        [manifestId, JSON.stringify(payload)]
      );
      return payload;
    });

    await logAudit({
      actorUserId: request.user.sub,
      action: 'upload.workbook',
      entityType: 'upload_manifest',
      entityId: summary.manifestId,
      metadata: { sourceName, sourceSha, summary }
    });

    return reply.code(201).send(summary);
  });
}

