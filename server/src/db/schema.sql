CREATE TABLE IF NOT EXISTS roles (
  id SERIAL PRIMARY KEY,
  name TEXT UNIQUE NOT NULL CHECK (name IN ('admin', 'analyst', 'viewer')),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email TEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS user_roles (
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role_id INT NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  PRIMARY KEY (user_id, role_id)
);

CREATE TABLE IF NOT EXISTS refresh_tokens (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash TEXT NOT NULL,
  expires_at TIMESTAMPTZ NOT NULL,
  revoked_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS audit_logs (
  id BIGSERIAL PRIMARY KEY,
  actor_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  action TEXT NOT NULL,
  entity_type TEXT NOT NULL,
  entity_id TEXT,
  metadata JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS upload_manifests (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  uploaded_by UUID NOT NULL REFERENCES users(id) ON DELETE RESTRICT,
  source_name TEXT NOT NULL,
  source_sha256 TEXT NOT NULL,
  source_size BIGINT NOT NULL,
  period_code TEXT,
  upload_mode TEXT NOT NULL DEFAULT 'manual' CHECK (upload_mode IN ('manual', 'daily', 'monthly')),
  status TEXT NOT NULL DEFAULT 'completed' CHECK (status IN ('processing', 'completed', 'failed')),
  summary JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (source_sha256)
);

CREATE TABLE IF NOT EXISTS kpi_rows_raw (
  id BIGSERIAL PRIMARY KEY,
  manifest_id UUID NOT NULL REFERENCES upload_manifests(id) ON DELETE CASCADE,
  dataset TEXT NOT NULL CHECK (dataset IN ('raw_data', 'incident')),
  source_row_no INT NOT NULL,
  row_payload JSONB NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS kpi_rows_curated (
  id BIGSERIAL PRIMARY KEY,
  dataset TEXT NOT NULL CHECK (dataset IN ('raw_data', 'incident')),
  row_key TEXT NOT NULL,
  row_hash TEXT NOT NULL,
  row_payload JSONB NOT NULL,
  latest_manifest_id UUID NOT NULL REFERENCES upload_manifests(id) ON DELETE RESTRICT,
  first_seen_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (dataset, row_key)
);

CREATE INDEX IF NOT EXISTS idx_kpi_rows_curated_dataset ON kpi_rows_curated(dataset);
CREATE INDEX IF NOT EXISTS idx_kpi_rows_curated_updated ON kpi_rows_curated(updated_at DESC);

INSERT INTO roles(name) VALUES ('admin'), ('analyst'), ('viewer')
ON CONFLICT (name) DO NOTHING;

