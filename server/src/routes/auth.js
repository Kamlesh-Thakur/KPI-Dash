import { getUserByEmail, verifyPassword, storeRefreshToken, rotateRefreshToken, revokeRefreshToken } from '../services/authService.js';
import { requireAuth } from '../middleware/auth.js';
import { logAudit } from '../services/auditService.js';
import { config } from '../config.js';

function accessPayload(user) {
  return { sub: user.id, email: user.email, roles: user.roles || [] };
}

export async function authRoutes(fastify) {
  fastify.post('/auth/login', async (request, reply) => {
    const body = request.body || {};
    const email = String(body.email || '').trim().toLowerCase();
    const password = String(body.password || '');
    if (!email || !password) return reply.code(400).send({ error: 'email and password are required' });

    const user = await getUserByEmail(email);
    if (!user || !user.is_active) return reply.code(401).send({ error: 'Invalid credentials' });
    const ok = await verifyPassword(password, user.password_hash);
    if (!ok) return reply.code(401).send({ error: 'Invalid credentials' });

    const accessToken = await reply.jwtSign(accessPayload(user), { expiresIn: config.jwtAccessTtl });
    const refreshToken = await fastify.jwt.sign(accessPayload(user), { secret: config.jwtRefreshSecret, expiresIn: config.jwtRefreshTtl });
    const refreshExp = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
    await storeRefreshToken(user.id, refreshToken, refreshExp);
    await logAudit({ actorUserId: user.id, action: 'login', entityType: 'auth', metadata: { email: user.email } });

    return { accessToken, refreshToken, user: { id: user.id, email: user.email, roles: user.roles || [] } };
  });

  fastify.post('/auth/refresh', async (request, reply) => {
    const { refreshToken } = request.body || {};
    if (!refreshToken) return reply.code(400).send({ error: 'refreshToken required' });
    let payload;
    try {
      payload = await fastify.jwt.verify(refreshToken, { secret: config.jwtRefreshSecret });
    } catch (_err) {
      return reply.code(401).send({ error: 'Invalid refresh token' });
    }
    const user = await getUserByEmail(payload.email);
    if (!user || !user.is_active) return reply.code(401).send({ error: 'Invalid refresh token' });
    const nextRefreshToken = await fastify.jwt.sign(accessPayload(user), { secret: config.jwtRefreshSecret, expiresIn: config.jwtRefreshTtl });
    const nextExp = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
    const rotated = await rotateRefreshToken(user.id, refreshToken, nextRefreshToken, nextExp);
    if (!rotated) return reply.code(401).send({ error: 'Refresh token expired or revoked' });

    const accessToken = await reply.jwtSign(accessPayload(user), { expiresIn: config.jwtAccessTtl });
    return { accessToken, refreshToken: nextRefreshToken, user: { id: user.id, email: user.email, roles: user.roles || [] } };
  });

  fastify.post('/auth/logout', { preHandler: [requireAuth] }, async (request) => {
    const { refreshToken } = request.body || {};
    if (refreshToken) await revokeRefreshToken(request.user.sub, refreshToken);
    await logAudit({ actorUserId: request.user.sub, action: 'logout', entityType: 'auth' });
    return { ok: true };
  });

  fastify.get('/auth/me', { preHandler: [requireAuth] }, async (request) => {
    const user = await getUserByEmail(request.user.email);
    if (!user) return { user: null };
    return { user: { id: user.id, email: user.email, roles: user.roles || [] } };
  });
}

