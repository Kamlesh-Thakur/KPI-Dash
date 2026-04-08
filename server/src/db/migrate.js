import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { pool } from './pool.js';
import { assertEnv } from '../config.js';

async function run() {
  assertEnv();
  const __dirname = path.dirname(fileURLToPath(import.meta.url));
  const sqlPath = path.join(__dirname, 'schema.sql');
  const sql = await fs.readFile(sqlPath, 'utf8');
  await pool.query('CREATE EXTENSION IF NOT EXISTS pgcrypto;');
  await pool.query(sql);
  await pool.end();
  console.log('Database migration completed.');
}

run().catch((err) => {
  console.error('Migration failed:', err.message);
  process.exitCode = 1;
});

