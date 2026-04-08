import bcrypt from 'bcryptjs';
import { pool } from './pool.js';
import { assertEnv } from '../config.js';

async function run() {
  assertEnv();
  const email = String(process.env.SEED_ADMIN_EMAIL || '').trim().toLowerCase();
  const password = String(process.env.SEED_ADMIN_PASSWORD || '');
  if (!email || !password) {
    throw new Error('SEED_ADMIN_EMAIL and SEED_ADMIN_PASSWORD are required.');
  }

  const passwordHash = await bcrypt.hash(password, 12);
  const userRes = await pool.query(
    `INSERT INTO users(email, password_hash)
     VALUES ($1, $2)
     ON CONFLICT (email) DO UPDATE SET password_hash = EXCLUDED.password_hash, updated_at = NOW()
     RETURNING id, email`,
    [email, passwordHash]
  );
  const user = userRes.rows[0];
  await pool.query(
    `INSERT INTO user_roles(user_id, role_id)
     SELECT $1, id FROM roles WHERE name = 'admin'
     ON CONFLICT (user_id, role_id) DO NOTHING`,
    [user.id]
  );
  console.log(`Admin seeded: ${user.email}`);
  await pool.end();
}

run().catch((err) => {
  console.error('Admin seed failed:', err.message);
  process.exitCode = 1;
});

