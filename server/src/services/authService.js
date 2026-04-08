import crypto from 'node:crypto';
import bcrypt from 'bcryptjs';
import { pool } from '../db/pool.js';

export async function getUserByEmail(email) {
  const res = await pool.query(
    `SELECT u.id, u.email, u.password_hash, u.is_active,
            ARRAY_REMOVE(ARRAY_AGG(r.name), NULL) AS roles
     FROM users u
     LEFT JOIN user_roles ur ON ur.user_id = u.id
     LEFT JOIN roles r ON r.id = ur.role_id
     WHERE u.email = $1
     GROUP BY u.id`,
    [email.toLowerCase()]
  );
  return res.rows[0] || null;
}

export async function verifyPassword(plain, passwordHash) {
  return bcrypt.compare(plain, passwordHash);
}

export function sha256(text) {
  return crypto.createHash('sha256').update(text).digest('hex');
}

export async function storeRefreshToken(userId, refreshToken, expiresAt) {
  const tokenHash = sha256(refreshToken);
  await pool.query(
    `INSERT INTO refresh_tokens(user_id, token_hash, expires_at)
     VALUES ($1, $2, $3)`,
    [userId, tokenHash, expiresAt]
  );
}

export async function rotateRefreshToken(userId, oldToken, newToken, newExpiry) {
  const oldHash = sha256(oldToken);
  const newHash = sha256(newToken);
  const revokeRes = await pool.query(
    `UPDATE refresh_tokens
     SET revoked_at = NOW()
     WHERE user_id = $1 AND token_hash = $2 AND revoked_at IS NULL AND expires_at > NOW()`,
    [userId, oldHash]
  );
  if (revokeRes.rowCount === 0) return false;
  await pool.query(
    `INSERT INTO refresh_tokens(user_id, token_hash, expires_at)
     VALUES ($1, $2, $3)`,
    [userId, newHash, newExpiry]
  );
  return true;
}

export async function revokeRefreshToken(userId, token) {
  const tokenHash = sha256(token);
  await pool.query(
    `UPDATE refresh_tokens SET revoked_at = NOW()
     WHERE user_id = $1 AND token_hash = $2 AND revoked_at IS NULL`,
    [userId, tokenHash]
  );
}

