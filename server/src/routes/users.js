import bcrypt from 'bcryptjs';
import { pool } from '../db/pool.js';
import { requireAuth, requireRoles } from '../middleware/auth.js';
import { logAudit } from '../services/auditService.js';

export async function userRoutes(fastify) {
  fastify.get('/users', { preHandler: [requireAuth, requireRoles('admin')] }, async () => {
    const res = await pool.query(
      `SELECT u.id, u.email, u.is_active, u.created_at,
              ARRAY_REMOVE(ARRAY_AGG(r.name), NULL) as roles
       FROM users u
       LEFT JOIN user_roles ur ON ur.user_id = u.id
       LEFT JOIN roles r ON r.id = ur.role_id
       GROUP BY u.id
       ORDER BY u.created_at DESC`
    );
    return { users: res.rows };
  });

  fastify.post('/users', { preHandler: [requireAuth, requireRoles('admin')] }, async (request, reply) => {
    const { email, password, roles = ['viewer'] } = request.body || {};
    const e = String(email || '').trim().toLowerCase();
    if (!e || !password) return reply.code(400).send({ error: 'email and password required' });

    const passHash = await bcrypt.hash(String(password), 12);
    const userRes = await pool.query(
      `INSERT INTO users(email, password_hash) VALUES ($1, $2)
       ON CONFLICT (email) DO NOTHING
       RETURNING id, email`,
      [e, passHash]
    );
    const user = userRes.rows[0];
    if (!user) return reply.code(409).send({ error: 'User already exists' });

    const validRoles = ['admin', 'analyst', 'viewer'].filter((r) => roles.includes(r));
    if (!validRoles.length) validRoles.push('viewer');
    for (const role of validRoles) {
      await pool.query(
        `INSERT INTO user_roles(user_id, role_id)
         SELECT $1, id FROM roles WHERE name = $2
         ON CONFLICT (user_id, role_id) DO NOTHING`,
        [user.id, role]
      );
    }
    await logAudit({
      actorUserId: request.user.sub,
      action: 'user.create',
      entityType: 'user',
      entityId: user.id,
      metadata: { email: user.email, roles: validRoles }
    });
    return reply.code(201).send({ user: { ...user, roles: validRoles } });
  });
}

