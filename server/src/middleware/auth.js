export async function requireAuth(request, reply) {
  try {
    await request.jwtVerify();
  } catch (_err) {
    return reply.code(401).send({ error: 'Unauthorized' });
  }
}

export function requireRoles(...roles) {
  const allowed = new Set(roles);
  return async function guard(request, reply) {
    const userRoles = request.user?.roles || [];
    const hasRole = userRoles.some((r) => allowed.has(r));
    if (!hasRole) {
      return reply.code(403).send({ error: 'Forbidden' });
    }
    return undefined;
  };
}

