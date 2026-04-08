export async function healthRoutes(fastify) {
  fastify.get('/health', async () => ({ ok: true, service: 'kpi-api' }));
}

