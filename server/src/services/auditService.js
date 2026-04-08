import { pool } from '../db/pool.js';

export async function logAudit({ actorUserId = null, action, entityType, entityId = null, metadata = {} }) {
  await pool.query(
    `INSERT INTO audit_logs(actor_user_id, action, entity_type, entity_id, metadata)
     VALUES ($1, $2, $3, $4, $5::jsonb)`,
    [actorUserId, action, entityType, entityId, JSON.stringify(metadata || {})]
  );
}

