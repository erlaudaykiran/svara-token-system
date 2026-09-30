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

-- Every issued token.
CREATE TABLE IF NOT EXISTS tokens (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    serial        TEXT NOT NULL UNIQUE,    -- B001, S001, SA001 ...
    type_key      TEXT NOT NULL REFERENCES counters(type_key),
    token_type    TEXT NOT NULL,           -- Royal Enfield / Silver / Saree
    name          TEXT NOT NULL,
    mobile        TEXT NOT NULL,
    payment       TEXT NOT NULL CHECK (payment IN ('Cash', 'UPI')),
    created_date  TEXT NOT NULL,           -- dd/mm/yyyy
    created_time  TEXT NOT NULL,           -- hh:mm:ss AM/PM
    created_at    TEXT NOT NULL            -- ISO timestamp
);

CREATE INDEX IF NOT EXISTS idx_tokens_type ON tokens(type_key);
