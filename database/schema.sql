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
    ('RE', 'Royal Enfield', 'B',  0),
    ('SI', 'Silver',        'S',  0),
    ('SA', 'Saree',         'SA', 0);

-- Every issued token with price and counter identification.
CREATE TABLE IF NOT EXISTS tokens (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    serial        TEXT NOT NULL UNIQUE,    -- B00001, S00001, SA00001 ...
    type_key      TEXT NOT NULL REFERENCES counters(type_key),
    token_type    TEXT NOT NULL,           -- Royal Enfield / Silver / Saree
    price         INTEGER NOT NULL DEFAULT 0,
    name          TEXT NOT NULL,
    mobile        TEXT NOT NULL,
    payment       TEXT NOT NULL CHECK (payment IN ('Cash', 'UPI')),
    created_by    TEXT NOT NULL DEFAULT 'admin',
    counter_name  TEXT NOT NULL DEFAULT 'Main Counter',
    created_date  TEXT NOT NULL,           -- dd/mm/yyyy
    created_time  TEXT NOT NULL,           -- hh:mm:ss AM/PM
    created_at    TEXT NOT NULL            -- ISO timestamp
);

CREATE INDEX IF NOT EXISTS idx_tokens_type ON tokens(type_key);

-- Active session per user so admin and all 3 counters have independent active logins
CREATE TABLE IF NOT EXISTS active_sessions (
    user_id       TEXT PRIMARY KEY,
    session_token TEXT NOT NULL,
    logged_in_at  TEXT NOT NULL
);

