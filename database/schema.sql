-- SVARA 2026 Lucky Draw - SQLite schema
PRAGMA journal_mode = WAL;

-- One row per token category; last_no is the last serial number issued.
CREATE TABLE IF NOT EXISTS counters (
    type_key  TEXT PRIMARY KEY,            -- RE / SI / SA
    label     TEXT NOT NULL,               -- Royal Enfield / Silver / Saree
    prefix    TEXT NOT NULL,               -- B / S / SA
    last_no   INTEGER NOT NULL DEFAULT 0
);

INSERT OR IGNORE INTO counters (type_key, label, prefix, last_no) VALUES
    ('RE', 'Royal Enfield',   'B',  0),
    ('SI', 'Silver',          'S',  0),
    ('SA', 'Saree',           'SA', 0),
    ('KA', 'Kunkuma Archana', 'KA', 0);

-- Every issued token with price, status, and counter identification.
CREATE TABLE IF NOT EXISTS tokens (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    serial        TEXT NOT NULL UNIQUE,    -- B00001, S00001, SA00001, KA00001 ...
    type_key      TEXT NOT NULL REFERENCES counters(type_key),
    token_type    TEXT NOT NULL,           -- Royal Enfield / Silver / Saree / Kunkuma Archana
    price         INTEGER NOT NULL DEFAULT 0,
    status        TEXT NOT NULL DEFAULT 'ACTIVE', -- ACTIVE / VOID
    voided_at     TEXT,
    voided_by     TEXT,
    name          TEXT NOT NULL,
    mobile        TEXT NOT NULL,
    payment       TEXT NOT NULL CHECK (payment IN ('Cash', 'UPI', 'Payment Pending', 'Pending')),
    created_by    TEXT NOT NULL DEFAULT 'admin',
    counter_name  TEXT NOT NULL DEFAULT 'Main Counter',
    created_date  TEXT NOT NULL,           -- dd/mm/yyyy
    created_time  TEXT NOT NULL,           -- hh:mm:ss AM/PM
    created_at    TEXT NOT NULL            -- ISO timestamp
);

CREATE INDEX IF NOT EXISTS idx_tokens_type ON tokens(type_key);
CREATE INDEX IF NOT EXISTS idx_tokens_user ON tokens(created_by);
CREATE INDEX IF NOT EXISTS idx_tokens_status ON tokens(status);
CREATE INDEX IF NOT EXISTS idx_tokens_mobile ON tokens(mobile);
CREATE INDEX IF NOT EXISTS idx_tokens_serial ON tokens(serial);

-- Active session per user so admin and all 3 counters have independent active logins
CREATE TABLE IF NOT EXISTS active_sessions (
    user_id       TEXT PRIMARY KEY,
    session_token TEXT NOT NULL,
    logged_in_at  TEXT NOT NULL
);

-- Reprint audit logs: records each reprint event for tracking
CREATE TABLE IF NOT EXISTS reprint_logs (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    token_id      INTEGER NOT NULL REFERENCES tokens(id) ON DELETE CASCADE,
    token_serial  TEXT NOT NULL,
    clerk_id      TEXT NOT NULL,
    clerk_name    TEXT NOT NULL,
    reason        TEXT DEFAULT 'Lost or torn receipt',
    reprinted_at  TEXT NOT NULL,
    timestamp     TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_reprint_token ON reprint_logs(token_serial);
CREATE INDEX IF NOT EXISTS idx_reprint_clerk ON reprint_logs(clerk_id);

-- System and Administrative security audit logs: immutable tracking of wipe attempts, backups, etc.
CREATE TABLE IF NOT EXISTS audit_logs (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    action        TEXT NOT NULL,
    user_id       TEXT NOT NULL,
    ip_address    TEXT,
    status        TEXT NOT NULL,
    details       TEXT,
    created_at    TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_logs(action);
CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_logs(user_id);
