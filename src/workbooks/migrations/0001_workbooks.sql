-- Run with the migration role in this service's own database.
-- One row per saved version of a workbook; the newest non-deleted version is current.
CREATE TABLE file_versions (
    tenant_id text NOT NULL,
    id uuid NOT NULL,
    path text NOT NULL,
    version integer NOT NULL,
    size bigint NOT NULL,
    modified_ms bigint NOT NULL,
    deleted boolean NOT NULL DEFAULT false,
    content_sha256 text NOT NULL,
    PRIMARY KEY (tenant_id, id),
    UNIQUE (tenant_id, path, version)
);
-- Workbook bytes, sealed in chunks of at most 1 MiB (each chunk is its own encryption resource).
CREATE TABLE file_chunks (
    tenant_id text NOT NULL,
    id uuid NOT NULL,
    version_id uuid NOT NULL,
    seq integer NOT NULL,
    content_ciphertext text NOT NULL,
    PRIMARY KEY (tenant_id, id),
    UNIQUE (tenant_id, version_id, seq),
    FOREIGN KEY (tenant_id, version_id) REFERENCES file_versions (tenant_id, id) ON DELETE CASCADE
);
-- Every plaintext read of a chunk is recorded before the bytes are released.
CREATE TABLE access_audit (
    tenant_id text NOT NULL,
    id uuid NOT NULL,
    at timestamptz NOT NULL DEFAULT now(),
    entity text NOT NULL,
    field text NOT NULL,
    level text NOT NULL,
    resource uuid NOT NULL,
    actor text NOT NULL,
    reason text NOT NULL,
    PRIMARY KEY (tenant_id, id)
);
CREATE TABLE env_meta (
    tenant_id text NOT NULL PRIMARY KEY,
    scenario text NOT NULL,
    generation integer NOT NULL
);
ALTER TABLE file_versions ENABLE ROW LEVEL SECURITY;
ALTER TABLE file_versions FORCE ROW LEVEL SECURITY;
CREATE POLICY forklaunch_tenant_policy ON file_versions
    USING (tenant_id = current_setting('forklaunch.tenant_id', true))
    WITH CHECK (tenant_id = current_setting('forklaunch.tenant_id', true));
ALTER TABLE file_chunks ENABLE ROW LEVEL SECURITY;
ALTER TABLE file_chunks FORCE ROW LEVEL SECURITY;
CREATE POLICY forklaunch_tenant_policy ON file_chunks
    USING (tenant_id = current_setting('forklaunch.tenant_id', true))
    WITH CHECK (tenant_id = current_setting('forklaunch.tenant_id', true));
ALTER TABLE access_audit ENABLE ROW LEVEL SECURITY;
ALTER TABLE access_audit FORCE ROW LEVEL SECURITY;
CREATE POLICY forklaunch_tenant_policy ON access_audit
    USING (tenant_id = current_setting('forklaunch.tenant_id', true))
    WITH CHECK (tenant_id = current_setting('forklaunch.tenant_id', true));
ALTER TABLE env_meta ENABLE ROW LEVEL SECURITY;
ALTER TABLE env_meta FORCE ROW LEVEL SECURITY;
CREATE POLICY forklaunch_tenant_policy ON env_meta
    USING (tenant_id = current_setting('forklaunch.tenant_id', true))
    WITH CHECK (tenant_id = current_setting('forklaunch.tenant_id', true));
-- Grant SELECT/INSERT/UPDATE/DELETE to the configured
-- restricted runtime role through deployment provisioning; no superuser credentials belong in DATABASE_URL.
