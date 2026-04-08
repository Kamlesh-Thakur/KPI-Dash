import Fastify from 'fastify';
import cors from '@fastify/cors';
import jwt from '@fastify/jwt';
import multipart from '@fastify/multipart';
import { assertEnv, config } from './config.js';
import { healthRoutes } from './routes/health.js';
import { authRoutes } from './routes/auth.js';
import { userRoutes } from './routes/users.js';
import { uploadRoutes } from './routes/uploads.js';
import { kpiRoutes } from './routes/kpi.js';

async function buildServer() {
  assertEnv();
  const app = Fastify({ logger: true });

  await app.register(cors, { origin: config.allowedOrigin, credentials: true });
  await app.register(jwt, { secret: config.jwtSecret });
  await app.register(multipart, { limits: { fileSize: 50 * 1024 * 1024 } });

  app.setErrorHandler((err, request, reply) => {
    request.log.error(err);
    if (reply.sent) return;
    reply.code(err.statusCode || 500).send({ error: err.message || 'Internal server error' });
  });

  await app.register(healthRoutes, { prefix: '/api' });
  await app.register(authRoutes, { prefix: '/api' });
  await app.register(userRoutes, { prefix: '/api' });
  await app.register(uploadRoutes, { prefix: '/api' });
  await app.register(kpiRoutes, { prefix: '/api' });
  return app;
}

buildServer()
  .then((app) => app.listen({ port: config.port, host: config.host }))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });

