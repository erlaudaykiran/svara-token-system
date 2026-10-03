"""
SVARA 2026 Lucky Draw - token printing backend.

Flask + SQLite / PostgreSQL.
Multi-User Counter Management & Secure Token Generation:
  - Administrator (admin): full management, all counters combined, edit and void controls
  - 3 Independent Counters (counter1, counter2, counter3): issuance and read-only records
  - Strictly Sequential & Thread-Safe: Concurrent requests locked for guaranteed sequential numbers
  - Non-Rolling Void: Voided tokens are marked VOID with audit trail; counters NEVER roll back
  - Automated Save: Tokens are saved atomically upon generation
  - Dynamic Pricing: Royal Enfield ₹301, Silver ₹201, Saree ₹101
"""
import base64
import glob
import gzip
import io
import json
import os
import re
import secrets
import shutil
import sqlite3
import sys
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

# Reconfigure stdout/stderr on Windows to prevent UnicodeEncodeError in console output
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
try:
    from zoneinfo import ZoneInfo
except ImportError:
    try:
        from backports.zoneinfo import ZoneInfo
    except ImportError:
        ZoneInfo = None
from functools import wraps

from flask import Flask, jsonify, request, send_file, send_from_directory, session
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Thread-level lock for process concurrency and SQLite atomic writes
db_lock = threading.Lock()

# Admin Mobile & Thank You Blessing Configuration
ADMIN_MOBILE = os.environ.get("ADMIN_MOBILE", "9848433020").strip()
THANK_YOU_MESSAGE = os.environ.get(
    "THANK_YOU_MESSAGE",
    "Thank you for registration! May Goddess Durgamatha bless you and your family."
).strip()


# Database Configuration:
# When DATABASE_URL is set (Render PostgreSQL, Neon, Supabase), use PostgreSQL.
# Otherwise, fall back to SQLite on local disk (database/svara.db).
DATABASE_URL = os.environ.get("DATABASE_URL")
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)
import escpos

ROOT = os.path.dirname(BASE)
DB_PATH = os.environ.get("SVARA_DB", os.path.join(ROOT, "database", "svara.db"))
FRONTEND = os.path.join(ROOT, "frontend")
EXPORT_DIR = os.path.join(ROOT, "exports")
EXPORT_FILE = os.path.join(EXPORT_DIR, "SVARA_2026_Tokens.xlsx")
EXPORT_ARCHANA_FILE = os.path.join(EXPORT_DIR, "SVARA_2026_Kunkuma_Archana.xlsx")
BACKUP_DIR = os.path.join(ROOT, "backups")

TZ_NAME = os.environ.get("SVARA_TZ", "Asia/Kolkata")

# Token Prices & Classification
PRICES = {
    "RE": 301,  # Royal Enfield
    "SI": 201,  # Silver
    "SA": 0,    # Saree (No price / free)
    "KA": 251   # Kunkuma Archana (Special Seva Puja Token)
}

LUCKY_DRAW_KEYS = ("RE", "SI", "SA")
ARCHANA_KEYS = ("KA",)


# User accounts configuration:
# 1 Admin (handles all counters) + 3 Separate Counter users with independent credentials
USERS = {
    os.environ.get("ADMIN_USER", os.environ.get("SVARA_USER", "admin")).strip().lower(): {
        "password": os.environ.get("ADMIN_PASSWORD", os.environ.get("SVARA_PASSWORD", "admin@svara2026")).strip(),
        "role": "admin",
        "name": os.environ.get("ADMIN_NAME", "Administrator"),
        "counter_name": "Main / All Counters"
    },
    os.environ.get("COUNTER1_USER", "counter1").strip().lower(): {
        "password": os.environ.get("COUNTER1_PASSWORD", "counter1@2026").strip(),
        "role": "counter",
        "name": os.environ.get("COUNTER1_NAME", "Counter 1"),
        "counter_name": os.environ.get("COUNTER1_NAME", "Counter 1")
    },
    os.environ.get("COUNTER2_USER", "counter2").strip().lower(): {
        "password": os.environ.get("COUNTER2_PASSWORD", "counter2@2026").strip(),
        "role": "counter",
        "name": os.environ.get("COUNTER2_NAME", "Counter 2"),
        "counter_name": os.environ.get("COUNTER2_NAME", "Counter 2")
    },
    os.environ.get("COUNTER3_USER", "counter3").strip().lower(): {
        "password": os.environ.get("COUNTER3_PASSWORD", "counter3@2026").strip(),
        "role": "counter",
        "name": os.environ.get("COUNTER3_NAME", "Counter 3"),
        "counter_name": os.environ.get("COUNTER3_NAME", "Counter 3")
    }
}

app = Flask(__name__, static_folder=None)
app.secret_key = os.environ.get("SECRET_KEY", "svara-token-system-secret-key-2026")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"


class DBWrapper:
    def __init__(self, conn, is_pg=False):
        self.conn = conn
        self.is_pg = is_pg

    def execute(self, sql, params=()):
        if self.is_pg:
            sql = sql.replace("?", "%s")
            cur = self.conn.cursor()
            cur.execute(sql, params)
            return cur
        else:
            return self.conn.execute(sql, params)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


def connect():
    url = os.environ.get("DATABASE_URL")
    if url and url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    if url:
        try:
            import psycopg2
            import psycopg2.extras
            conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
            conn.autocommit = False
            return DBWrapper(conn, is_pg=True)
        except Exception as e:
            print(f"[WARN] Failed to connect to PostgreSQL ({e}). Falling back to SQLite.")

    conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    conn.row_factory = sqlite3.Row
    return DBWrapper(conn, is_pg=False)


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    os.makedirs(BACKUP_DIR, exist_ok=True)
    db = connect()
    try:
        if db.is_pg:
            db.execute("""
                CREATE TABLE IF NOT EXISTS counters (
                    type_key  VARCHAR(10) PRIMARY KEY,
                    label     VARCHAR(50) NOT NULL,
                    prefix    VARCHAR(10) NOT NULL,
                    last_no   INTEGER NOT NULL DEFAULT 0
                );
            """)
            db.execute("""
                INSERT INTO counters (type_key, label, prefix, last_no) VALUES
                    ('RE', 'Royal Enfield',   'B',  0),
                    ('SI', 'Silver',          'S',  0),
                    ('SA', 'Saree',           'SA', 0),
                    ('KA', 'Kunkuma Archana', 'KA', 0)
                ON CONFLICT (type_key) DO NOTHING;
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS tokens (
                    id            SERIAL PRIMARY KEY,
                    serial        VARCHAR(20) NOT NULL UNIQUE,
                    type_key      VARCHAR(10) NOT NULL REFERENCES counters(type_key),
                    token_type    VARCHAR(50) NOT NULL,
                    price         INTEGER NOT NULL DEFAULT 0,
                    status        VARCHAR(20) NOT NULL DEFAULT 'ACTIVE',
                    voided_at     VARCHAR(40),
                    voided_by     VARCHAR(50),
                    name          VARCHAR(100) NOT NULL,
                    mobile        VARCHAR(20) NOT NULL,
                    payment       VARCHAR(30) NOT NULL CHECK (payment IN ('Cash', 'UPI', 'Payment Pending', 'Pending')),
                    created_by    VARCHAR(50) NOT NULL DEFAULT 'admin',
                    counter_name  VARCHAR(50) NOT NULL DEFAULT 'Main Counter',
                    created_date  VARCHAR(20) NOT NULL,
                    created_time  VARCHAR(20) NOT NULL,
                    created_at    VARCHAR(40) NOT NULL
                );
            """)
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_type ON tokens(type_key);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_user ON tokens(created_by);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_status ON tokens(status);")

            # Safe column additions if table was created in an earlier build
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS price INTEGER NOT NULL DEFAULT 0;")
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS status VARCHAR(20) NOT NULL DEFAULT 'ACTIVE';")
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS voided_at VARCHAR(40);")
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS voided_by VARCHAR(50);")
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS created_by VARCHAR(50) NOT NULL DEFAULT 'admin';")
            db.execute("ALTER TABLE tokens ADD COLUMN IF NOT EXISTS counter_name VARCHAR(50) NOT NULL DEFAULT 'Main Counter';")
            try:
                db.execute("ALTER TABLE tokens DROP CONSTRAINT IF EXISTS tokens_payment_check;")
                db.execute("ALTER TABLE tokens ADD CONSTRAINT tokens_payment_check CHECK (payment IN ('Cash', 'UPI', 'Payment Pending', 'Pending'));")
            except Exception:
                pass

            # Active sessions per user ID for concurrent logins (drop legacy id-based table if exists)
            db.execute("""
                DO $$
                BEGIN
                    IF EXISTS (
                        SELECT 1 FROM information_schema.columns 
                        WHERE table_name = 'active_sessions' AND column_name = 'id'
                    ) THEN
                        DROP TABLE active_sessions;
                    END IF;
                END $$;
            """)
            db.execute("""
                CREATE TABLE IF NOT EXISTS active_sessions (
                    user_id       VARCHAR(50) PRIMARY KEY,
                    session_token TEXT NOT NULL,
                    logged_in_at  TEXT NOT NULL
                );
            """)

            # Reprint Audit Logging Table
            db.execute("""
                CREATE TABLE IF NOT EXISTS reprint_logs (
                    id            SERIAL PRIMARY KEY,
                    token_id      INTEGER NOT NULL,
                    token_serial  VARCHAR(20) NOT NULL,
                    clerk_id      VARCHAR(50) NOT NULL,
                    clerk_name    VARCHAR(100) NOT NULL,
                    reason        TEXT DEFAULT 'Lost or torn receipt',
                    reprinted_at  VARCHAR(40) NOT NULL,
                    timestamp     VARCHAR(40) NOT NULL DEFAULT ''
                );
            """)
            try:
                db.execute("ALTER TABLE reprint_logs ADD COLUMN IF NOT EXISTS timestamp VARCHAR(40) NOT NULL DEFAULT '';")
            except Exception:
                pass
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_token ON reprint_logs(token_serial);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_clerk ON reprint_logs(clerk_id);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_time ON reprint_logs(timestamp);")

            # System & Admin Security Audit Logs (Wipe attempts, backups, security events)
            db.execute("""
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id            SERIAL PRIMARY KEY,
                    action        VARCHAR(50) NOT NULL,
                    user_id       VARCHAR(50) NOT NULL,
                    ip_address    VARCHAR(50),
                    status        VARCHAR(20) NOT NULL,
                    details       TEXT,
                    created_at    VARCHAR(40) NOT NULL
                );
            """)
            db.execute("CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_logs(action);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_logs(user_id);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_mobile ON tokens(mobile);")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_serial ON tokens(serial);")
            db.commit()
        else:
            with open(os.path.join(ROOT, "database", "schema.sql"), encoding="utf-8") as f:
                db.conn.executescript(f.read())

            # Migrations for SQLite if upgraded from earlier versions
            cur = db.execute("PRAGMA table_info(tokens)")
            cols = [r["name"] for r in cur.fetchall()]
            if "price" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN price INTEGER NOT NULL DEFAULT 0")
            if "status" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN status TEXT NOT NULL DEFAULT 'ACTIVE'")
            if "voided_at" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN voided_at TEXT")
            if "voided_by" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN voided_by TEXT")
            if "created_by" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN created_by TEXT NOT NULL DEFAULT 'admin'")
            if "counter_name" not in cols:
                db.execute("ALTER TABLE tokens ADD COLUMN counter_name TEXT NOT NULL DEFAULT 'Main Counter'")

            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_user ON tokens(created_by)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_status ON tokens(status)")

            cur_s = db.execute("PRAGMA table_info(active_sessions)")
            s_rows = cur_s.fetchall()
            s_cols = [r["name"] for r in s_rows]
            if not s_rows or "user_id" not in s_cols:
                db.execute("DROP TABLE IF EXISTS active_sessions")
                db.execute("""
                    CREATE TABLE active_sessions (
                        user_id       TEXT PRIMARY KEY,
                        session_token TEXT NOT NULL,
                        logged_in_at  TEXT NOT NULL
                    )
                """)

            # SQLite Reprint Audit Logging Table
            db.execute("""
                CREATE TABLE IF NOT EXISTS reprint_logs (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    token_id      INTEGER NOT NULL,
                    token_serial  TEXT NOT NULL,
                    clerk_id      TEXT NOT NULL,
                    clerk_name    TEXT NOT NULL,
                    reason        TEXT DEFAULT 'Lost or torn receipt',
                    reprinted_at  TEXT NOT NULL,
                    timestamp     TEXT NOT NULL DEFAULT ''
                )
            """)
            cur_rp = db.execute("PRAGMA table_info(reprint_logs)")
            rp_cols = [r["name"] for r in cur_rp.fetchall()]
            if "timestamp" not in rp_cols:
                db.execute("ALTER TABLE reprint_logs ADD COLUMN timestamp TEXT NOT NULL DEFAULT ''")
                db.execute("UPDATE reprint_logs SET timestamp = reprinted_at WHERE timestamp = ''")
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_token ON reprint_logs(token_serial)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_clerk ON reprint_logs(clerk_id)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_reprint_time ON reprint_logs(timestamp)")

            # SQLite Audit Logs Table
            db.execute("""
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    action        TEXT NOT NULL,
                    user_id       TEXT NOT NULL,
                    ip_address    TEXT,
                    status        TEXT NOT NULL,
                    details       TEXT,
                    created_at    TEXT NOT NULL
                )
            """)
            db.execute("CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_logs(action)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_logs(user_id)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_mobile ON tokens(mobile)")
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_serial ON tokens(serial)")

        # Ensure all category rows exist in counters table
        for tkey, lbl, pfx in [('RE', 'Royal Enfield', 'B'), ('SI', 'Silver', 'S'), ('SA', 'Saree', 'SA'), ('KA', 'Kunkuma Archana', 'KA')]:
            if db.is_pg:
                db.execute(
                    "INSERT INTO counters (type_key, label, prefix, last_no) VALUES (?, ?, ?, 0) "
                    "ON CONFLICT (type_key) DO NOTHING",
                    (tkey, lbl, pfx)
                )
            else:
                db.execute(
                    "INSERT OR IGNORE INTO counters (type_key, label, prefix, last_no) VALUES (?, ?, ?, 0)",
                    (tkey, lbl, pfx)
                )
        db.commit()

        # Counter synchronization & Floor protection:
        # Guarantee counters.last_no is NEVER lower than any existing token serial number
        # and NEVER lower than any configured environment minimum floor (MIN_RE_NO, MIN_SI_NO, MIN_SA_NO, MIN_KA_NO)
        floors = {
            "RE": int(os.environ.get("MIN_RE_NO", os.environ.get("START_RE_NO", 0))),
            "SI": int(os.environ.get("MIN_SI_NO", os.environ.get("START_SI_NO", 0))),
            "SA": int(os.environ.get("MIN_SA_NO", os.environ.get("START_SA_NO", 0))),
            "KA": int(os.environ.get("MIN_KA_NO", os.environ.get("START_KA_NO", 0)))
        }

        counters_rows = db.execute("SELECT type_key, prefix, last_no FROM counters").fetchall()
        for c in counters_rows:
            tk = c["type_key"]
            pfx = c["prefix"]
            curr_last = c["last_no"]
            target_last = max(curr_last, floors.get(tk, 0))

            tok_rows = db.execute("SELECT serial FROM tokens WHERE type_key = ?", (tk,)).fetchall()
            for tr in tok_rows:
                s = tr["serial"]
                if s and s.startswith(pfx):
                    digits = s[len(pfx):]
                    if digits.isdigit():
                        target_last = max(target_last, int(digits))

            if target_last > curr_last:
                db.execute("UPDATE counters SET last_no = ? WHERE type_key = ?", (target_last, tk))
        db.commit()
    finally:
        db.close()


def get_tz():
    if ZoneInfo:
        try:
            return ZoneInfo(TZ_NAME)
        except Exception:
            pass
    return timezone(timedelta(hours=5, minutes=30))


def now_parts():
    d = datetime.now(get_tz())
    return d.strftime("%d/%m/%Y"), d.strftime("%I:%M:%S %p"), d.isoformat(timespec="seconds")


def authenticate_user(username, password):
    user_key = str(username).strip().lower()
    password = str(password).strip()

    # Allow "administrator" alias for whatever admin user is configured
    if user_key in ("admin", "administrator"):
        for k, v in USERS.items():
            if v["role"] == "admin":
                user_key = k
                break

    if user_key not in USERS:
        return None

    cfg = USERS[user_key]
    configured_password = cfg["password"].strip()

    allowed_passwords = [configured_password]
    # Retain convenience fallbacks only if default passwords have not been changed
    if cfg["role"] == "admin" and configured_password == "admin@svara2026":
        allowed_passwords.extend(["svara@2026", "admin", "admin123"])
    elif user_key == "counter1" and configured_password == "counter1@2026":
        allowed_passwords.append("counter1")
    elif user_key == "counter2" and configured_password == "counter2@2026":
        allowed_passwords.append("counter2")
    elif user_key == "counter3" and configured_password == "counter3@2026":
        allowed_passwords.append("counter3")

    allowed_passwords = [p.strip() for p in allowed_passwords if p and p.strip()]

    for p in allowed_passwords:
        if secrets.compare_digest(password, p):
            return {
                "username": user_key,
                "role": cfg["role"],
                "name": cfg["name"],
                "counter_name": cfg["counter_name"]
            }
    return None


def get_current_user():
    token = session.get("token")
    username = session.get("user")
    if not token or not username:
        return None
    username = username.strip().lower()
    if username not in USERS:
        return None

    db = connect()
    try:
        cur = db.execute("SELECT session_token FROM active_sessions WHERE user_id = ?", (username,))
        row = cur.fetchone()
        if row:
            token_val = row["session_token"] if isinstance(row, dict) or hasattr(row, "__getitem__") else row[0]
            if secrets.compare_digest(token_val, token):
                cfg = USERS[username]
                return {
                    "username": username,
                    "role": cfg["role"],
                    "name": cfg["name"],
                    "counter_name": cfg["counter_name"]
                }
        return None
    except Exception as e:
        print(f"[WARN] Error fetching active session for {username}: {e}")
        return None
    finally:
        db.close()


def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        user = get_current_user()
        if not user:
            return jsonify(error="Authentication required. Please sign in.", auth=False), 401
        return f(user, *args, **kwargs)
    return decorated


def fmt_serial(prefix, n):
    return f"{prefix}{n:05d}"



def row_val(r, key, default=None):
    if r is None:
        return default
    if isinstance(r, dict):
        return r.get(key, default)
    try:
        val = r[key]
        return val if val is not None else default
    except (IndexError, KeyError):
        return default

def log_reprint(db, token_id, token_serial, clerk_id, clerk_name, reason="Lost or torn receipt"):
    """
    Middleware/audit function to log every token reprint event.
    Records token_id, token_serial, clerk_id, clerk_name, reason, and timestamp.
    """
    now_iso = datetime.now().isoformat()
    try:
        db.execute(
            "INSERT INTO reprint_logs (token_id, token_serial, clerk_id, clerk_name, reason, reprinted_at, timestamp) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (token_id, token_serial, clerk_id, clerk_name, reason, now_iso, now_iso)
        )
    except Exception as e:
        print(f"[WARN] Failed to write reprint log: {e}")


def log_audit(db, action, user_id, status, details="", ip_address=""):
    """
    Records security-critical administrative actions (wipe attempts, backups, etc.)
    into an immutable audit log table.
    """
    now_iso = datetime.now().isoformat()
    try:
        db.execute(
            "INSERT INTO audit_logs (action, user_id, ip_address, status, details, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (action, user_id, ip_address, status, details, now_iso)
        )
    except Exception as e:
        print(f"[WARN] Failed to write audit log: {e}")


def upload_to_cloud(filepath, filename):
    """
    Uploads compressed backup to cloud storage (Google Drive or Webhook).
    """
    # 1. Custom Webhook / Cloud Storage Endpoint
    webhook_url = os.environ.get("CLOUD_BACKUP_WEBHOOK_URL")
    if webhook_url:
        try:
            with open(filepath, "rb") as f:
                data = f.read()
            req = urllib.request.Request(
                webhook_url,
                data=data,
                headers={
                    "Content-Type": "application/gzip",
                    "X-Filename": filename,
                    "User-Agent": "SVARA-Backup/1.0"
                }
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                print(f"[CLOUD BACKUP] Uploaded to cloud webhook (HTTP {resp.status})")
                return {"uploaded": True, "provider": "Webhook", "status": resp.status}
        except Exception as e:
            print(f"[WARN Cloud Backup] Webhook upload failed: {e}")
            return {"uploaded": False, "provider": "Webhook", "error": str(e)}

    # 2. Google Drive API Hook
    gdrive_folder = os.environ.get("GDRIVE_FOLDER_ID")
    if gdrive_folder:
        try:
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaFileUpload
            from google.oauth2 import service_account
            sa_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
            if sa_path and os.path.exists(sa_path):
                creds = service_account.Credentials.from_service_account_file(
                    sa_path,
                    scopes=['https://www.googleapis.com/auth/drive.file']
                )
                service = build('drive', 'v3', credentials=creds)
                file_metadata = {'name': filename, 'parents': [gdrive_folder]}
                media = MediaFileUpload(filepath, mimetype='application/gzip')
                drive_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                print(f"[GDRIVE] Backup uploaded successfully: {drive_file.get('id')}")
                return {"uploaded": True, "provider": "GoogleDrive", "file_id": drive_file.get('id')}
        except Exception as e:
            print(f"[WARN GDrive] Cloud upload failed: {e}")
            return {"uploaded": False, "provider": "GoogleDrive", "error": str(e)}

    return {"uploaded": False, "provider": "LocalOnly", "message": "Saved securely to local backup directory."}


def create_compressed_backup(tag="manual"):
    """
    Creates an atomic compressed SQL dump (.sql.gz) of all tables.
    Rotates existing backups keeping the newest 30.
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"backup_svara_{timestamp_str}_{tag}.sql.gz"
    backup_filepath = os.path.join(BACKUP_DIR, backup_filename)

    db = connect()
    try:
        sql_lines = [
            f"-- SVARA 2026 Database Backup ({timestamp_str})",
            "-- Automated Compressed Database Export\n"
        ]

        # Dump counters
        sql_lines.append("-- TABLE: counters")
        c_rows = db.execute("SELECT * FROM counters").fetchall()
        for r in c_rows:
            tk = r["type_key"]
            lbl = str(r["label"]).replace("'", "''")
            pfx = r["prefix"]
            lno = r["last_no"]
            sql_lines.append(f"INSERT INTO counters (type_key, label, prefix, last_no) VALUES ('{tk}', '{lbl}', '{pfx}', {lno}) ON CONFLICT (type_key) DO UPDATE SET last_no = {lno};")

        # Dump tokens
        sql_lines.append("\n-- TABLE: tokens")
        t_rows = db.execute("SELECT * FROM tokens ORDER BY id ASC").fetchall()
        for r in t_rows:
            ser = str(r["serial"]).replace("'", "''")
            tk = str(r["type_key"]).replace("'", "''")
            tt = str(r["token_type"]).replace("'", "''")
            pr = int(r["price"] or 0)
            st = str(r["status"] or "ACTIVE").replace("'", "''")
            vat = f"'{r['voided_at']}'" if row_val(r, "voided_at") else "NULL"
            vby = f"'{r['voided_by']}'" if row_val(r, "voided_by") else "NULL"
            nm = str(r["name"] or "").replace("'", "''")
            mob = str(r["mobile"] or "").replace("'", "''")
            pay = str(r["payment"] or "").replace("'", "''")
            cby = str(r["created_by"] or "admin").replace("'", "''")
            cnm = str(r["counter_name"] or "Main Counter").replace("'", "''")
            cdt = str(r["created_date"] or "").replace("'", "''")
            ctm = str(r["created_time"] or "").replace("'", "''")
            cat = str(r["created_at"] or "").replace("'", "''")
            sql_lines.append(
                f"INSERT INTO tokens (serial, type_key, token_type, price, status, voided_at, voided_by, name, mobile, payment, created_by, counter_name, created_date, created_time, created_at) "
                f"VALUES ('{ser}', '{tk}', '{tt}', {pr}, '{st}', {vat}, {vby}, '{nm}', '{mob}', '{pay}', '{cby}', '{cnm}', '{cdt}', '{ctm}', '{cat}') "
                f"ON CONFLICT (serial) DO NOTHING;"
            )

        # Dump reprint_logs
        try:
            rp_rows = db.execute("SELECT * FROM reprint_logs ORDER BY id ASC").fetchall()
            if rp_rows:
                sql_lines.append("\n-- TABLE: reprint_logs")
                for r in rp_rows:
                    tid = r["token_id"]
                    tser = str(r["token_serial"]).replace("'", "''")
                    cid = str(r["clerk_id"]).replace("'", "''")
                    cnm = str(r["clerk_name"]).replace("'", "''")
                    rsn = str(row_val(r, "reason", "")).replace("'", "''")
                    rat = str(r["reprinted_at"]).replace("'", "''")
                    ts = str(row_val(r, "timestamp", rat)).replace("'", "''")
                    sql_lines.append(
                        f"INSERT INTO reprint_logs (token_id, token_serial, clerk_id, clerk_name, reason, reprinted_at, timestamp) "
                        f"VALUES ({tid}, '{tser}', '{cid}', '{cnm}', '{rsn}', '{rat}', '{ts}');"
                    )
        except Exception:
            pass

        # Dump audit_logs
        try:
            al_rows = db.execute("SELECT * FROM audit_logs ORDER BY id ASC").fetchall()
            if al_rows:
                sql_lines.append("\n-- TABLE: audit_logs")
                for r in al_rows:
                    act = str(r["action"]).replace("'", "''")
                    uid = str(r["user_id"]).replace("'", "''")
                    ip = str(row_val(r, "ip_address", "")).replace("'", "''")
                    st = str(r["status"]).replace("'", "''")
                    dtl = str(row_val(r, "details", "")).replace("'", "''")
                    cat = str(r["created_at"]).replace("'", "''")
                    sql_lines.append(
                        f"INSERT INTO audit_logs (action, user_id, ip_address, status, details, created_at) "
                        f"VALUES ('{act}', '{uid}', '{ip}', '{st}', '{dtl}', '{cat}');"
                    )
        except Exception:
            pass

        sql_content = "\n".join(sql_lines).encode("utf-8")

        # Compress to .sql.gz
        with gzip.open(backup_filepath, "wb") as gz_file:
            gz_file.write(sql_content)

        file_size_kb = round(os.path.getsize(backup_filepath) / 1024, 2)
        print(f"[BACKUP] Created compressed backup: {backup_filename} ({file_size_kb} KB)")

        # Rotate old backups keeping newest 30
        try:
            existing = sorted(
                glob.glob(os.path.join(BACKUP_DIR, "backup_svara_*.sql.gz")),
                key=os.path.getmtime
            )
            while len(existing) > 30:
                oldest = existing.pop(0)
                try:
                    os.remove(oldest)
                except Exception:
                    pass
        except Exception:
            pass

        cloud_status = upload_to_cloud(backup_filepath, backup_filename)

        return {
            "success": True,
            "filename": backup_filename,
            "filepath": backup_filepath,
            "size_kb": file_size_kb,
            "token_count": len(t_rows),
            "cloud_status": cloud_status
        }
    finally:
        db.close()


_scheduler_started = False
_scheduler_lock = threading.Lock()

def start_backup_scheduler():
    global _scheduler_started
    with _scheduler_lock:
        if _scheduler_started:
            return
        _scheduler_started = True

    def backup_loop():
        # Wait 45 seconds after process start
        time.sleep(45)
        while True:
            try:
                create_compressed_backup(tag="scheduled")
            except Exception as e:
                print(f"[WARN] Scheduled periodic backup error: {e}")
            # Run every 6 hours (21600 seconds)
            time.sleep(21600)

    t = threading.Thread(target=backup_loop, daemon=True, name="BackupScheduler")
    t.start()


def mask_mobile(m):
    digits = re.sub(r"\D", "", str(m))
    if len(digits) >= 10:
        return f"+91 {digits[:2]}****{digits[-4:]}"
    return str(m)


def send_sms(mobile, message):
    """
    Dispatches SMS to mobile number.
    Supports Fast2SMS (FAST2SMS_API_KEY), custom HTTP SMS gateway (SMS_GATEWAY_URL),
    and always logs to server console for audit trail.
    """
    clean_mobile = re.sub(r"\D", "", str(mobile))
    if len(clean_mobile) == 12 and clean_mobile.startswith("91"):
        clean_mobile = clean_mobile[2:]

    try:
        print(f"\n[SMS DISPATCH] Destination: +91 {clean_mobile}\nMessage:\n{message}\n")
    except Exception:
        safe_msg = message.encode("ascii", "replace").decode("ascii")
        print(f"\n[SMS DISPATCH] Destination: +91 {clean_mobile}\nMessage:\n{safe_msg}\n")

    # 1. Fast2SMS Provider
    fast2sms_key = os.environ.get("FAST2SMS_API_KEY")
    if fast2sms_key:
        try:
            req_data = json.dumps({
                "route": "q",
                "message": message,
                "language": "english",
                "flash": 0,
                "numbers": clean_mobile
            }).encode("utf-8")
            req = urllib.request.Request(
                "https://www.fast2sms.com/dev/bulkV2",
                data=req_data,
                headers={
                    "authorization": fast2sms_key.strip(),
                    "Content-Type": "application/json",
                    "User-Agent": "SVARA-SMS/1.0"
                }
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                result = json.loads(resp.read().decode("utf-8"))
                return {"success": True, "provider": "Fast2SMS", "response": result}
        except Exception as e:
            print(f"[WARN Fast2SMS] Dispatch error: {e}")

    # 2. Generic Custom SMS Gateway (e.g., https://gateway/send?apikey={api_key}&to={mobile}&msg={message})
    gateway_url = os.environ.get("SMS_GATEWAY_URL")
    if gateway_url:
        try:
            formatted_url = gateway_url.format(
                mobile=clean_mobile,
                message=urllib.parse.quote_plus(message),
                api_key=os.environ.get("SMS_API_KEY", "")
            )
            req = urllib.request.Request(formatted_url, headers={"User-Agent": "SVARA-SMS/1.0"})
            with urllib.request.urlopen(req, timeout=8) as resp:
                return {"success": True, "provider": "CustomGateway", "status": resp.status}
        except Exception as e:
            print(f"[WARN CustomGateway] Dispatch error: {e}")

    return {"success": True, "provider": "ConsoleAudit", "message": "Dispatched to console log"}


def build_receipt_message(devotee_name, tokens_list, payment, total_amount, counter_name, date_str, time_str):
    serials = [t["serial"] for t in tokens_list]
    if len(serials) == 1:
        ser_text = serials[0]
    elif len(serials) <= 3:
        ser_text = ", ".join(serials)
    else:
        ser_text = f"{serials[0]} to {serials[-1]} ({len(serials)} tokens)"

    cat_name = tokens_list[0]["type"]
    price_val = tokens_list[0].get("price", 0)

    lines = [
        "🌸 SVARA 2026 LUCKY DRAW 🌸",
        "Official Token Receipt",
        "-----------------------------",
        f"Devotee: {devotee_name}",
        f"Token No: {ser_text}",
        f"Category: {cat_name}"
    ]
    if price_val and price_val > 0:
        lines.append(f"Amount: Rs. {total_amount} ({payment})")

    lines.extend([
        f"Date: {date_str} | {time_str}",
        f"Counter: {counter_name}",
        "-----------------------------",
        THANK_YOU_MESSAGE,
        "Contact: +91 9848433020, +91 9885897093"
    ])
    return "\n".join(lines)


def register_user_session(user):
    token = secrets.token_hex(24)
    now_iso = datetime.now().isoformat()
    db = connect()
    try:
        if db.is_pg:
            db.execute(
                "INSERT INTO active_sessions (user_id, session_token, logged_in_at) VALUES (?, ?, ?) "
                "ON CONFLICT (user_id) DO UPDATE SET session_token = EXCLUDED.session_token, logged_in_at = EXCLUDED.logged_in_at",
                (user["username"], token, now_iso)
            )
        else:
            db.execute(
                "INSERT INTO active_sessions (user_id, session_token, logged_in_at) VALUES (?, ?, ?) "
                "ON CONFLICT (user_id) DO UPDATE SET session_token = excluded.session_token, logged_in_at = excluded.logged_in_at",
                (user["username"], token, now_iso)
            )
        db.commit()
    finally:
        db.close()

    session["token"] = token
    session["user"] = user["username"]
    return token


COLUMNS = ["Token No", "Type", "Price (₹)", "Status", "Name", "Mobile", "Payment", "Date", "Time", "Issued By", "Counter"]


def row_val(r, key, default=None):
    if r is None:
        return default
    if isinstance(r, dict):
        return r.get(key, default)
    try:
        val = r[key]
        return val if val is not None else default
    except (IndexError, KeyError):
        return default


def build_workbook(rows, title="Tokens"):
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(COLUMNS)
    for c in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7A1424")
        cell.alignment = Alignment(horizontal="center")

    total_amount = 0
    active_count = 0
    void_count = 0

    void_fill = PatternFill("solid", fgColor="FDE8E8")
    void_font = Font(color="B00020", bold=True)

    for r in rows:
        price_val = row_val(r, "price")
        if price_val is None:
            price_val = PRICES.get(row_val(r, "type_key"), 0)
        status_val = row_val(r, "status", "ACTIVE")

        price_int = int(price_val)
        if status_val == "ACTIVE":
            total_amount += price_int
            active_count += 1
        else:
            void_count += 1

        price_cell = price_int if price_int > 0 else "-"

        row_cells = [
            row_val(r, "serial", ""),
            row_val(r, "token_type", ""),
            price_cell,
            status_val,
            row_val(r, "name", ""),
            row_val(r, "mobile", ""),
            row_val(r, "payment", ""),
            row_val(r, "created_date", ""),
            row_val(r, "created_time", ""),
            row_val(r, "created_by", "admin"),
            row_val(r, "counter_name", "Counter")
        ]
        ws.append(row_cells)


        if status_val == "VOID":
            current_row_idx = ws.max_row
            for col_idx in range(1, len(COLUMNS) + 1):
                c = ws.cell(row=current_row_idx, column=col_idx)
                c.fill = void_fill
            ws.cell(row=current_row_idx, column=4).font = void_font

    # Summary row (active collection only)
    summary_row = len(rows) + 2
    ws.cell(row=summary_row, column=1, value=f"{active_count} Active ({void_count} Void)").font = Font(bold=True)
    ws.cell(row=summary_row, column=2, value="TOTAL (ACTIVE):").font = Font(bold=True)
    ws.cell(row=summary_row, column=3, value=total_amount).font = Font(bold=True)

    col_widths = [12, 16, 11, 10, 26, 14, 10, 12, 13, 14, 16]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].number_format = "@"
    ws.freeze_panes = "A2"
    return wb


def all_rows(db, user=None, counter_filter=None, newest_first=False, category=None):
    order = "DESC" if newest_first else "ASC"
    conditions = []
    params = []

    if user and user["role"] != "admin":
        u_clean = user["username"].lower().replace(" ", "")
        conditions.append("(REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?)")
        params.extend([u_clean, u_clean])
    elif counter_filter and counter_filter != "all":
        cf = counter_filter.strip().lower()
        if cf in ("pending", "payment pending"):
            conditions.append("payment IN ('Payment Pending', 'Pending')")
        elif cf in ("archana", "kunkuma archana", "ka"):
            conditions.append("type_key = 'KA'")
        elif cf in ("luckydraw", "lucky draw", "ld"):
            conditions.append("type_key IN ('RE', 'SI', 'SA')")
        else:
            cf_clean = cf.replace(" ", "")
            conditions.append("(REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?)")
            params.extend([cf_clean, cf_clean])

    if category:
        cat_lower = str(category).lower().strip()
        if cat_lower in ("archana", "kunkuma archana", "ka"):
            conditions.append("type_key = 'KA'")
        elif cat_lower in ("luckydraw", "lucky draw"):
            conditions.append("type_key IN ('RE', 'SI', 'SA')")
        elif cat_lower.upper() in ("RE", "SI", "SA", "KA"):
            conditions.append("type_key = ?")
            params.append(cat_lower.upper())

    where_clause = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    cur = db.execute(f"SELECT * FROM tokens {where_clause} ORDER BY id {order}", tuple(params))
    return cur.fetchall()


def refresh_excel_file(db):
    try:
        ld_rows = all_rows(db, category="luckydraw")
        build_workbook(ld_rows, title="Lucky Draw Tokens").save(EXPORT_FILE)
    except OSError:
        pass
    try:
        ka_rows = all_rows(db, category="archana")
        build_workbook(ka_rows, title="Kunkuma Archana").save(EXPORT_ARCHANA_FILE)
    except OSError:
        pass


def token_json(r):
    price_val = row_val(r, "price")
    if price_val is None:
        price_val = PRICES.get(row_val(r, "type_key"), 0)
    status_val = row_val(r, "status", "ACTIVE")
    return {
        "id": row_val(r, "id"),
        "serial": row_val(r, "serial", ""),
        "type_key": row_val(r, "type_key", ""),
        "type": row_val(r, "token_type", ""),
        "price": int(price_val),
        "status": status_val,
        "is_void": (status_val == "VOID"),
        "voided_at": row_val(r, "voided_at"),
        "voided_by": row_val(r, "voided_by"),
        "name": row_val(r, "name", ""),
        "mobile": row_val(r, "mobile", ""),
        "payment": row_val(r, "payment", ""),
        "created_by": row_val(r, "created_by", "admin"),
        "counter_name": row_val(r, "counter_name", "Counter"),
        "date": row_val(r, "created_date", ""),
        "time": row_val(r, "created_time", "")
    }


@app.post("/api/auth/login")
def login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    user = authenticate_user(username, password)
    if not user:
        return jsonify(error="Invalid User ID or Password."), 401

    register_user_session(user)
    return jsonify(
        success=True,
        username=user["username"],
        role=user["role"],
        name=user["name"],
        counter_name=user["counter_name"]
    )


@app.post("/api/auth/logout")
def logout():
    user = get_current_user()
    if user:
        db = connect()
        try:
            db.execute("DELETE FROM active_sessions WHERE user_id = ?", (user["username"],))
            db.commit()
        finally:
            db.close()
    session.clear()
    return jsonify(success=True)


@app.get("/api/auth/status")
def auth_status():
    user = get_current_user()
    if user:
        return jsonify(
            authenticated=True,
            username=user["username"],
            role=user["role"],
            name=user["name"],
            counter_name=user["counter_name"]
        )
    return jsonify(authenticated=False)


@app.get("/api/counters")
@require_auth
def counters(user):
    db = connect()
    try:
        categories = {}
        for c in db.execute("SELECT * FROM counters").fetchall():
            tk = c["type_key"]
            # Active count
            count_cur = db.execute("SELECT COUNT(*) FROM tokens WHERE type_key = ? AND status = 'ACTIVE'", (tk,))
            count_row = count_cur.fetchone()
            tot_count = list(count_row.values())[0] if isinstance(count_row, dict) else count_row[0]

            # Void count
            void_cur = db.execute("SELECT COUNT(*) FROM tokens WHERE type_key = ? AND status = 'VOID'", (tk,))
            void_row = void_cur.fetchone()
            void_count = list(void_row.values())[0] if isinstance(void_row, dict) else void_row[0]

            # Current user active count
            u_clean = user["username"].lower().replace(" ", "")
            user_count_cur = db.execute(
                "SELECT COUNT(*) FROM tokens WHERE type_key = ? AND status = 'ACTIVE' "
                "AND (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?)",
                (tk, u_clean, u_clean)
            )
            user_count_row = user_count_cur.fetchone()
            user_count = list(user_count_row.values())[0] if isinstance(user_count_row, dict) else user_count_row[0]

            categories[tk] = {
                "label": c["label"],
                "price": PRICES.get(tk, 0),
                "issued": tot_count,
                "voided": void_count,
                "user_issued": user_count,
                "next": fmt_serial(c["prefix"], c["last_no"] + 1)
            }

        # Lucky Draw combined stats (RE, SI, SA strictly isolated)
        ld_cur = db.execute(
            "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
            "WHERE type_key IN ('RE', 'SI', 'SA') AND status = 'ACTIVE'"
        )
        ld_row = ld_cur.fetchone()
        ld_vals = list(ld_row.values()) if isinstance(ld_row, dict) else ld_row
        ld_active_count, ld_active_amount = ld_vals[0], int(ld_vals[1])

        ld_vcur = db.execute("SELECT COUNT(*) FROM tokens WHERE type_key IN ('RE', 'SI', 'SA') AND status = 'VOID'")
        ld_vrow = ld_vcur.fetchone()
        ld_void_count = list(ld_vrow.values())[0] if isinstance(ld_vrow, dict) else ld_vrow[0]

        # Kunkuma Archana combined stats (KA strictly isolated)
        ka_cur = db.execute(
            "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
            "WHERE type_key = 'KA' AND status = 'ACTIVE'"
        )
        ka_row = ka_cur.fetchone()
        ka_vals = list(ka_row.values()) if isinstance(ka_row, dict) else ka_row
        ka_active_count, ka_active_amount = ka_vals[0], int(ka_vals[1])

        ka_vcur = db.execute("SELECT COUNT(*) FROM tokens WHERE type_key = 'KA' AND status = 'VOID'")
        ka_vrow = ka_vcur.fetchone()
        ka_void_count = list(ka_vrow.values())[0] if isinstance(ka_vrow, dict) else ka_vrow[0]

        # Payment Pending stats
        pend_cur = db.execute(
            "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
            "WHERE payment IN ('Payment Pending', 'Pending') AND status = 'ACTIVE'"
        )
        pend_row = pend_cur.fetchone()
        p_vals = list(pend_row.values()) if isinstance(pend_row, dict) else pend_row
        pending_count, pending_amount = p_vals[0], int(p_vals[1])

        # Counter Breakdown for Admin (handles spaces, counter1, counter 1, Counter 1)
        counter_breakdown = {}
        if user["role"] == "admin":
            for ukey in USERS:
                uk_clean = ukey.lower().replace(" ", "")
                # Lucky draw for this counter
                ld_c_cur = db.execute(
                    "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
                    "WHERE (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?) "
                    "AND type_key IN ('RE', 'SI', 'SA') AND status = 'ACTIVE'",
                    (uk_clean, uk_clean)
                )
                ld_c_row = ld_c_cur.fetchone()
                ld_c_vals = list(ld_c_row.values()) if isinstance(ld_c_row, dict) else ld_c_row
                c_ld_cnt, c_ld_amt = ld_c_vals[0], int(ld_c_vals[1])

                ld_vc_cur = db.execute(
                    "SELECT COUNT(*) FROM tokens "
                    "WHERE (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?) "
                    "AND type_key IN ('RE', 'SI', 'SA') AND status = 'VOID'",
                    (uk_clean, uk_clean)
                )
                ld_vc_row = ld_vc_cur.fetchone()
                c_ld_vcnt = list(ld_vc_row.values())[0] if isinstance(ld_vc_row, dict) else ld_vc_row[0]

                # Archana for this counter
                ka_c_cur = db.execute(
                    "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
                    "WHERE (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?) "
                    "AND type_key = 'KA' AND status = 'ACTIVE'",
                    (uk_clean, uk_clean)
                )
                ka_c_row = ka_c_cur.fetchone()
                ka_c_vals = list(ka_c_row.values()) if isinstance(ka_c_row, dict) else ka_c_row
                c_ka_cnt, c_ka_amt = ka_c_vals[0], int(ka_c_vals[1])

                # Pending for this counter
                p_c_cur = db.execute(
                    "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
                    "WHERE (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?) "
                    "AND payment IN ('Payment Pending', 'Pending') AND status = 'ACTIVE'",
                    (uk_clean, uk_clean)
                )
                p_c_row = p_c_cur.fetchone()
                p_c_vals = list(p_c_row.values()) if isinstance(p_c_row, dict) else p_c_row
                c_p_cnt, c_p_amt = p_c_vals[0], int(p_c_vals[1])

                counter_breakdown[ukey] = {
                    "name": USERS[ukey]["name"],
                    "count": c_ld_cnt,
                    "amount": c_ld_amt,
                    "void_count": c_ld_vcnt,
                    "archana_count": c_ka_cnt,
                    "archana_amount": c_ka_amt,
                    "pending_count": c_p_cnt,
                    "pending_amount": c_p_amt,
                    "total_count": c_ld_cnt + c_ka_cnt,
                    "total_amount": c_ld_amt + c_ka_amt
                }

        # Current user stats
        my_clean = user["username"].lower().replace(" ", "")
        user_summary_cur = db.execute(
            "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens "
            "WHERE (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?) "
            "AND status = 'ACTIVE'",
            (my_clean, my_clean)
        )
        user_sum_row = user_summary_cur.fetchone()
        u_vals = list(user_sum_row.values()) if isinstance(user_sum_row, dict) else user_sum_row
        my_cnt, my_amt = u_vals[0], int(u_vals[1])

    finally:
        db.close()

    return jsonify({
        "categories": categories,
        "currentUser": user,
        "myStats": {"count": my_cnt, "amount": my_amt},
        "luckyDrawStats": {
            "count": ld_active_count,
            "amount": ld_active_amount,
            "void_count": ld_void_count
        },
        "archanaStats": {
            "count": ka_active_count,
            "amount": ka_active_amount,
            "void_count": ka_void_count
        },
        "pendingStats": {
            "count": pending_count,
            "amount": pending_amount
        },
        "counterBreakdown": counter_breakdown if user["role"] == "admin" else None
    })


@app.post("/api/tokens")
@require_auth
def create_token(user):
    """
    Atomically generates sequential tokens.
    Thread-safe and process-safe with row locking (PostgreSQL FOR UPDATE / SQLite BEGIN IMMEDIATE).
    Guarantees no race condition or duplicate serial numbers when multiple counters issue simultaneously.
    """
    data = request.get_json(silent=True) or {}
    type_key = str(data.get("type", "")).strip().upper()
    name = re.sub(r"\s+", " ", str(data.get("name", ""))).strip()
    mobile = str(data.get("mobile", "")).strip()
    payment = str(data.get("payment", "")).strip()

    try:
        quantity = int(data.get("quantity", 1))
        if quantity < 1 or quantity > 100:
            return jsonify(error="Number of tokens must be between 1 and 100."), 400
    except (TypeError, ValueError):
        return jsonify(error="Invalid number of tokens entered."), 400

    if not name or len(name) > 60:
        return jsonify(error="Enter the customer name (max 60 characters)."), 400
    if not re.fullmatch(r"\d{10}", mobile):
        return jsonify(error="Enter a valid 10-digit mobile number."), 400
    if payment in ("Payment Pending", "Pending"):
        payment = "Payment Pending"
    elif payment not in ("Cash", "UPI"):
        return jsonify(error="Select a valid payment type: Cash, UPI, or Payment Pending."), 400

    # Acquire threading lock for local multi-thread serialization
    with db_lock:
        db = connect()
        try:
            if not db.is_pg:
                db.execute("BEGIN IMMEDIATE")
                cur = db.execute("SELECT * FROM counters WHERE type_key = ?", (type_key,))
            else:
                cur = db.execute("SELECT * FROM counters WHERE type_key = ? FOR UPDATE", (type_key,))

            c = cur.fetchone()
            if c is None:
                db.rollback()
                return jsonify(error="Unknown token type."), 400

            unit_price = PRICES.get(type_key, 0)
            start_no = c["last_no"]
            end_no = start_no + quantity
            date, time_, iso = now_parts()
            created_tokens = []

            for n in range(start_no + 1, end_no + 1):
                serial = fmt_serial(c["prefix"], n)
                db.execute(
                    "INSERT INTO tokens (serial, type_key, token_type, price, status, name, mobile, payment,"
                    " created_by, counter_name, created_date, created_time, created_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (serial, type_key, c["label"], unit_price, "ACTIVE", name, mobile, payment,
                     user["username"], user["counter_name"], date, time_, iso)
                )
                created_tokens.append({
                    "serial": serial,
                    "type_key": type_key,
                    "type": c["label"],
                    "price": unit_price,
                    "status": "ACTIVE",
                    "name": name,
                    "mobile": mobile,
                    "payment": payment,
                    "created_by": user["username"],
                    "counter_name": user["counter_name"],
                    "date": date,
                    "time": time_
                })

            db.execute("UPDATE counters SET last_no = ? WHERE type_key = ?", (end_no, type_key))
            db.commit()
            refresh_excel_file(db)
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    send_digital = bool(data.get("send_digital", True))
    digital_receipt_data = None
    if send_digital and created_tokens:
        receipt_text = build_receipt_message(
            devotee_name=name,
            tokens_list=created_tokens,
            payment=payment,
            total_amount=unit_price * len(created_tokens),
            counter_name=user["counter_name"],
            date_str=created_tokens[0]["date"],
            time_str=created_tokens[0]["time"]
        )
        sms_res = send_sms(mobile, receipt_text)
        wa_url = f"https://api.whatsapp.com/send?phone=91{mobile}&text={urllib.parse.quote_plus(receipt_text)}"
        digital_receipt_data = {
            "sms_sent": sms_res.get("success", False),
            "sms_provider": sms_res.get("provider", "ConsoleAudit"),
            "whatsapp_url": wa_url,
            "receipt_text": receipt_text,
            "mobile": mobile
        }

    # Generate individual ESC/POS byte streams with Auto-Cutter for each token (Customer & Office copies)
    raw_escpos_all = bytearray()
    tokens_escpos = []
    for t in created_tokens:
        t_stream = escpos.generate_escpos_stream_for_token(t, include_office_copy=True, is_reprint=False)
        raw_escpos_all.extend(t_stream)
        tokens_escpos.append(base64.b64encode(t_stream).decode("ascii"))
    escpos_bytes = bytes(raw_escpos_all)
    escpos_b64 = base64.b64encode(escpos_bytes).decode("ascii")
    escpos_text = "\n\n".join([escpos.generate_receipt_text(t) for t in created_tokens])

    res = {
        "tokens": created_tokens,
        "tokens_escpos": tokens_escpos,
        "count": len(created_tokens),
        "total_amount": unit_price * len(created_tokens),
        "first": created_tokens[0]["serial"],
        "last": created_tokens[-1]["serial"],
        "serial": created_tokens[0]["serial"],
        "type_key": type_key,
        "type": created_tokens[0]["type"],
        "price": unit_price,
        "status": "ACTIVE",
        "name": created_tokens[0]["name"],
        "mobile": created_tokens[0]["mobile"],
        "payment": created_tokens[0]["payment"],
        "created_by": user["username"],
        "counter_name": user["counter_name"],
        "date": created_tokens[0]["date"],
        "time": created_tokens[0]["time"],
        "digital_receipt": digital_receipt_data,
        "thank_you_message": THANK_YOU_MESSAGE,
        "escpos_base64": escpos_b64,
        "escpos_text": escpos_text
    }
    return jsonify(res), 201


@app.post("/api/tokens/<serial>/receipt")
@require_auth
def resend_token_receipt(user, serial):
    """
    Generates and pushes digital receipt (SMS and WhatsApp link) for an existing token.
    """
    serial = str(serial).strip().upper()
    db = connect()
    try:
        row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        if not row:
            return jsonify(error=f"Token {serial} was not found."), 404
        t = token_json(row)
        receipt_text = build_receipt_message(
            devotee_name=t["name"],
            tokens_list=[t],
            payment=t["payment"],
            total_amount=t["price"],
            counter_name=t["counter_name"],
            date_str=t["date"],
            time_str=t["time"]
        )
        sms_res = send_sms(t["mobile"], receipt_text)
        wa_url = f"https://api.whatsapp.com/send?phone=91{t['mobile']}&text={urllib.parse.quote_plus(receipt_text)}"
        return jsonify(
            success=True,
            sms_sent=sms_res.get("success", False),
            whatsapp_url=wa_url,
            receipt_text=receipt_text,
            mobile=t["mobile"],
            thank_you_message=THANK_YOU_MESSAGE
        )
    finally:
        db.close()


@app.get("/api/tokens")
@require_auth
def list_tokens(user):
    counter_filter = request.args.get("counter", "").strip().lower()
    db = connect()
    try:
        rows = all_rows(db, user=user, counter_filter=counter_filter, newest_first=True)
        return jsonify([token_json(r) for r in rows])
    finally:
        db.close()


@app.put("/api/tokens/<serial>")
@require_auth
def edit_token(user, serial):
    """
    Admin-only token details update.
    Allows correcting customer name, mobile, and payment mode.
    """
    if user["role"] != "admin":
        return jsonify(error="Permission denied: Only Administrator can edit token details."), 403

    serial = str(serial).strip().upper()
    data = request.get_json(silent=True) or {}
    name = re.sub(r"\s+", " ", str(data.get("name", ""))).strip()
    mobile = str(data.get("mobile", "")).strip()
    payment = str(data.get("payment", "")).strip()

    if not name or len(name) > 60:
        return jsonify(error="Enter customer name (max 60 characters)."), 400
    if not re.fullmatch(r"\d{10}", mobile):
        return jsonify(error="Enter a valid 10-digit mobile number."), 400
    if payment in ("Payment Pending", "Pending"):
        payment = "Payment Pending"
    elif payment not in ("Cash", "UPI"):
        return jsonify(error="Select a valid payment type: Cash, UPI, or Payment Pending."), 400

    db = connect()
    try:
        row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        if not row:
            return jsonify(error=f"Token {serial} was not found."), 404

        current_status = row_val(row, "status", "ACTIVE")
        if current_status == "VOID":
            return jsonify(error=f"Cannot edit token {serial} because it has been VOIDED."), 400

        db.execute(
            "UPDATE tokens SET name = ?, mobile = ?, payment = ? WHERE UPPER(serial) = ?",
            (name, mobile, payment, serial)
        )
        db.commit()
        refresh_excel_file(db)

        updated_row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        return jsonify(success=True, token=token_json(updated_row), message=f"Token {serial} updated successfully.")
    finally:
        db.close()


@app.post("/api/tokens/<serial>/mark-paid")
@require_auth
def mark_token_paid(user, serial):
    """
    Staff / Admin action to update a 'Payment Pending' token to 'Cash' or 'UPI'
    once the devotee has completed payment.
    """
    serial = str(serial).strip().upper()
    data = request.get_json(silent=True) or {}
    new_payment = str(data.get("payment", "Cash")).strip()
    if new_payment not in ("Cash", "UPI"):
        return jsonify(error="Payment mode must be Cash or UPI."), 400

    db = connect()
    try:
        row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        if not row:
            return jsonify(error=f"Token {serial} was not found."), 404

        status_val = row_val(row, "status", "ACTIVE")
        if status_val == "VOID":
            return jsonify(error=f"Cannot update payment for VOID token {serial}."), 400

        db.execute("UPDATE tokens SET payment = ? WHERE UPPER(serial) = ?", (new_payment, serial))
        db.commit()
        refresh_excel_file(db)

        updated_row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        return jsonify(
            success=True,
            token=token_json(updated_row),
            message=f"Token {serial} payment updated to {new_payment} successfully."
        )
    finally:
        db.close()



@app.post("/api/tokens/reprint")
@require_auth
def reprint_token(user):
    """
    1. Reprint Token Option:
       Allows counter staff to search by Reference ID (serial) or Mobile Number
       to reprint a clean copy of the token. Records every reprint action
       in the reprint_logs audit table (token_id, clerk_id, timestamp).
    """
    data = request.get_json(silent=True) or {}
    query = str(data.get("query", "")).strip()
    reason = str(data.get("reason", "Lost or torn receipt")).strip()
    token_id = data.get("token_id")

    if not query and not token_id:
        return jsonify(error="Please provide a Token Serial / Reference ID or Devotee Mobile Number."), 400

    db = connect()
    try:
        matched_token = None
        if token_id:
            row = db.execute("SELECT * FROM tokens WHERE id = ?", (token_id,)).fetchone()
            if row:
                matched_token = token_json(row)

        if not matched_token and query:
            clean_q = query.upper()
            # 1. Try exact serial match
            row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (clean_q,)).fetchone()
            if row:
                matched_token = token_json(row)
            else:
                # 2. Try mobile number match
                clean_mobile = re.sub(r"\D", "", query)
                if len(clean_mobile) >= 10:
                    clean_mobile = clean_mobile[-10:]
                    rows = db.execute("SELECT * FROM tokens WHERE mobile = ? ORDER BY id DESC", (clean_mobile,)).fetchall()
                    if len(rows) == 1:
                        matched_token = token_json(rows[0])
                    elif len(rows) > 1:
                        tokens_list = [token_json(r) for r in rows]
                        return jsonify(
                            multiple_matches=True,
                            count=len(tokens_list),
                            tokens=tokens_list,
                            message=f"Found {len(tokens_list)} tokens for mobile {clean_mobile}. Please choose which token to reprint."
                        )

        if not matched_token:
            return jsonify(error=f"No token record found matching '{query}'."), 404

        # Audit Logging: Record reprint event (token_id, clerk_id, timestamp)
        now_dt = datetime.now()
        now_iso = now_dt.isoformat()
        reprinted_human = now_dt.strftime("%d/%m/%Y %I:%M:%S %p")
        log_reprint(
            db=db,
            token_id=matched_token["id"],
            token_serial=matched_token["serial"],
            clerk_id=user["username"],
            clerk_name=user["name"],
            reason=reason
        )
        db.commit()

        # Add reprint flag and metadata
        matched_token["is_reprint"] = True
        matched_token["reprint_reason"] = reason
        matched_token["reprinted_at"] = reprinted_human
        matched_token["timestamp"] = now_iso

        # Format receipt using raw ESC/POS commands with Auto-Cutter for 80mm thermal printer
        escpos_bytes = escpos.generate_escpos_stream_for_token(matched_token, include_office_copy=True, is_reprint=True)
        escpos_b64 = base64.b64encode(escpos_bytes).decode("ascii")
        escpos_text = escpos.generate_receipt_text(matched_token, is_reprint=True)

        return jsonify(
            success=True,
            reprint=True,
            token=matched_token,
            reprinted_at=reprinted_human,
            timestamp=now_iso,
            clerk_id=user["username"],
            clerk_name=user["name"],
            escpos_base64=escpos_b64,
            escpos_text=escpos_text,
            message=f"Clean reprint generated for Token {matched_token['serial']}."
        )
    finally:
        db.close()


@app.get("/api/tokens/reprint-logs")
@require_auth
def get_reprint_logs(user):
    """
    Returns recent reprint audit logs including token_id, clerk_id, and timestamp.
    """
    limit = int(request.args.get("limit", 50))
    db = connect()
    try:
        rows = db.execute(
            "SELECT * FROM reprint_logs ORDER BY id DESC LIMIT ?",
            (limit,)
        ).fetchall()
        logs = []
        for r in rows:
            rat = row_val(r, "reprinted_at", "")
            ts = row_val(r, "timestamp", rat)
            logs.append({
                "id": r["id"],
                "token_id": r["token_id"],
                "token_serial": r["token_serial"],
                "clerk_id": r["clerk_id"],
                "clerk_name": r["clerk_name"],
                "reason": row_val(r, "reason", ""),
                "timestamp": ts,
                "reprinted_at": rat
            })
        return jsonify(logs=logs, count=len(logs))
    finally:
        db.close()


@app.get("/api/tokens/<id_or_serial>/escpos")
@require_auth
def get_token_escpos(user, id_or_serial):
    """
    Generates and returns raw ESC/POS commands formatted for an 80mm thermal receipt printer (48 chars/line).
    Appends the standard paper cut command at the end of the stream to trigger the auto-cutter.
    Supports format=download (.bin), format=text (48-char plain text), or JSON (base64).
    """
    target = str(id_or_serial).strip()
    is_reprint = request.args.get("reprint", "false").lower() in ("true", "1", "yes")
    reason = request.args.get("reason", "Lost or torn receipt")
    fmt = request.args.get("format", "json").lower()

    db = connect()
    try:
        if target.isdigit():
            row = db.execute("SELECT * FROM tokens WHERE id = ?", (int(target),)).fetchone()
        else:
            row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (target.upper(),)).fetchone()
        if not row:
            return jsonify(error=f"Token '{target}' not found."), 404

        t = token_json(row)
        if is_reprint:
            t["is_reprint"] = True
            t["reprint_reason"] = reason
            t["reprinted_at"] = datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")

        raw_bytes = escpos.generate_escpos_stream_for_token(t, include_office_copy=True, is_reprint=is_reprint)
        text_preview = escpos.generate_receipt_text(t, is_reprint=is_reprint)

        if fmt in ("download", "bin", "raw"):
            return send_file(
                io.BytesIO(raw_bytes),
                mimetype="application/octet-stream",
                as_attachment=True,
                download_name=f"token_{t['serial']}.bin"
            )
        elif fmt == "text":
            return text_preview, 200, {"Content-Type": "text/plain; charset=utf-8"}
        else:
            return jsonify({
                "success": True,
                "serial": t["serial"],
                "is_reprint": is_reprint,
                "escpos_base64": base64.b64encode(raw_bytes).decode("ascii"),
                "escpos_text": text_preview,
                "byte_count": len(raw_bytes)
            })
    finally:
        db.close()


@app.post("/api/tokens/print-escpos")
@require_auth
def print_token_escpos(user):
    """
    Directly dispatches raw ESC/POS commands with auto-cutter to a connected
    80mm thermal receipt printer (e.g. ATPOS AT-301 via USB spooler or TCP network).
    """
    data = request.get_json(silent=True) or {}
    query = str(data.get("query", data.get("serial", ""))).strip()
    token_id = data.get("token_id")
    serials = data.get("serials")
    printer_name = data.get("printer_name")
    is_reprint = bool(data.get("is_reprint", False))
    reason = str(data.get("reason", "Lost or torn receipt")).strip()

    if not query and not token_id and not serials:
        return jsonify(error="Please provide token_id, serial, or serials list."), 400

    db = connect()
    try:
        matched_tokens = []
        if serials and isinstance(serials, list):
            for s in serials:
                r = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (str(s).strip().upper(),)).fetchone()
                if r:
                    matched_tokens.append(token_json(r))
        elif token_id:
            r = db.execute("SELECT * FROM tokens WHERE id = ?", (token_id,)).fetchone()
            if r:
                matched_tokens.append(token_json(r))
        elif query:
            clean_q = query.upper()
            r = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (clean_q,)).fetchone()
            if r:
                matched_tokens.append(token_json(r))

        if not matched_tokens:
            return jsonify(error=f"No matching tokens found to print."), 404

        if is_reprint:
            for matched in matched_tokens:
                log_reprint(
                    db=db,
                    token_id=matched["id"],
                    token_serial=matched["serial"],
                    clerk_id=user["username"],
                    clerk_name=user["name"],
                    reason=reason
                )
                matched["is_reprint"] = True
                matched["reprint_reason"] = reason
                matched["reprinted_at"] = datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")
            db.commit()

        # Execute unified reusable print-and-cut engine
        ok, msg = escpos.print_tokens_with_autocut(
            tokens_list=matched_tokens,
            printer_name=printer_name,
            include_office_copy=True,
            is_reprint=is_reprint
        )

        raw_bytes_all = bytearray()
        tokens_escpos = []
        for m in matched_tokens:
            s_bytes = escpos.generate_escpos_stream_for_token(m, include_office_copy=True, is_reprint=is_reprint)
            raw_bytes_all.extend(s_bytes)
            tokens_escpos.append(base64.b64encode(s_bytes).decode("ascii"))

        return jsonify({
            "success": ok,
            "message": msg,
            "token": matched_tokens[0],
            "tokens": matched_tokens,
            "tokens_escpos": tokens_escpos,
            "escpos_base64": base64.b64encode(bytes(raw_bytes_all)).decode("ascii"),
            "byte_count": len(raw_bytes_all)
        }), (200 if ok else 500)
    finally:
        db.close()


@app.get("/api/printers")
@require_auth
def list_printers(user):
    """
    Returns list of local and network printers detected by the system.
    """
    installed = escpos.get_installed_printers()
    default_p = None
    if escpos.win32print is not None:
        try:
            default_p = escpos.win32print.GetDefaultPrinter()
        except Exception:
            pass
    thermal_candidate = None
    for name in installed:
        if any(k in name.lower() for k in ["atpos", "pos", "thermal", "receipt", "at-301", "80mm", "tm-t", "tvs", "rp"]):
            thermal_candidate = name
            break
    return jsonify({
        "printers": installed,
        "default_printer": default_p,
        "recommended_thermal_printer": thermal_candidate or default_p
    })


@app.get("/api/devotees/search")
@require_auth
def search_devotees(user):
    """
    2. Quick Lookup / Search by Mobile Number:
       Fast, indexed search endpoint that queries devotee records and recent
       bookings by mobile number (or serial) to pull up details instantly without re-typing names.
    """
    q = str(request.args.get("q", "")).strip()
    if not q:
        return jsonify(results=[], devotee=None, count=0)

    clean_digits = re.sub(r"\D", "", q)
    db = connect()
    try:
        if re.search(r"[A-Za-z]", q):
            # Contains letters -> search by Serial or Name
            search_param = f"%{q.upper()}%"
            rows = db.execute(
                "SELECT * FROM tokens WHERE UPPER(serial) LIKE ? OR UPPER(name) LIKE ? ORDER BY id DESC LIMIT 15",
                (search_param, search_param)
            ).fetchall()
        elif clean_digits:
            # Pure digits -> search by Mobile Number
            search_param = f"%{clean_digits}%"
            rows = db.execute(
                "SELECT * FROM tokens WHERE mobile LIKE ? ORDER BY id DESC LIMIT 15",
                (search_param,)
            ).fetchall()
        else:
            search_param = f"%{q.upper()}%"
            rows = db.execute(
                "SELECT * FROM tokens WHERE UPPER(serial) LIKE ? OR UPPER(name) LIKE ? ORDER BY id DESC LIMIT 15",
                (search_param, search_param)
            ).fetchall()

        results = [token_json(r) for r in rows]
        devotee_info = None
        if results:
            latest = results[0]
            devotee_info = {
                "name": latest["name"],
                "mobile": latest["mobile"],
                "last_serial": latest["serial"],
                "last_type": latest["type"],
                "total_bookings": len(results)
            }

        return jsonify(
            query=q,
            count=len(results),
            results=results,
            devotee=devotee_info
        )
    finally:
        db.close()


@app.get("/api/reports/shift")
@require_auth
def shift_inventory_report(user):
    """
    3. Inventory Breakdown & Grand Totals Report:
       Generates an end-of-shift / daily report showing individual token counts
       and total amounts for each specific category (Royal Enfield, Silver, Saree),
       with a summary footer displaying the full combined grand totals.
    """
    date_filter = request.args.get("date", "today").strip()
    counter_filter = request.args.get("counter", "all").strip()

    if date_filter.lower() == "today":
        today_date, _, _ = now_parts()
        target_date = today_date
    elif date_filter.lower() == "all":
        target_date = None
    else:
        target_date = date_filter

    db = connect()
    try:
        query = "SELECT * FROM tokens WHERE 1=1"
        params = []

        if target_date:
            query += " AND created_date = ?"
            params.append(target_date)

        if user["role"] != "admin":
            u_clean = user["username"].lower().replace(" ", "")
            query += " AND (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?)"
            params.extend([u_clean, u_clean])
        elif counter_filter and counter_filter != "all":
            cf_clean = counter_filter.lower().replace(" ", "")
            query += " AND (REPLACE(LOWER(created_by), ' ', '') = ? OR REPLACE(LOWER(counter_name), ' ', '') = ?)"
            params.extend([cf_clean, cf_clean])

        query += " ORDER BY id ASC"
        rows = db.execute(query, tuple(params)).fetchall()

        categories = {
            "RE": {
                "key": "RE",
                "label": "Royal Enfield",
                "prefix": "B",
                "unit_price": PRICES.get("RE", 301),
                "active_count": 0,
                "void_count": 0,
                "total_count": 0,
                "cash_count": 0,
                "cash_amount": 0,
                "upi_count": 0,
                "upi_amount": 0,
                "pending_count": 0,
                "pending_amount": 0,
                "total_amount": 0,
                "first_serial": None,
                "last_serial": None
            },
            "SI": {
                "key": "SI",
                "label": "Silver",
                "prefix": "S",
                "unit_price": PRICES.get("SI", 201),
                "active_count": 0,
                "void_count": 0,
                "total_count": 0,
                "cash_count": 0,
                "cash_amount": 0,
                "upi_count": 0,
                "upi_amount": 0,
                "pending_count": 0,
                "pending_amount": 0,
                "total_amount": 0,
                "first_serial": None,
                "last_serial": None
            },
            "SA": {
                "key": "SA",
                "label": "Saree",
                "prefix": "SA",
                "unit_price": PRICES.get("SA", 0),
                "active_count": 0,
                "void_count": 0,
                "total_count": 0,
                "cash_count": 0,
                "cash_amount": 0,
                "upi_count": 0,
                "upi_amount": 0,
                "pending_count": 0,
                "pending_amount": 0,
                "total_amount": 0,
                "first_serial": None,
                "last_serial": None
            },
            "KA": {
                "key": "KA",
                "label": "Kunkuma Archana",
                "prefix": "KA",
                "unit_price": PRICES.get("KA", 251),
                "active_count": 0,
                "void_count": 0,
                "total_count": 0,
                "cash_count": 0,
                "cash_amount": 0,
                "upi_count": 0,
                "upi_amount": 0,
                "pending_count": 0,
                "pending_amount": 0,
                "total_amount": 0,
                "first_serial": None,
                "last_serial": None
            }
        }

        # Lucky Draw Grand Totals (RE, SI, SA strictly isolated)
        lucky_draw_totals = {
            "active_count": 0,
            "void_count": 0,
            "total_count": 0,
            "cash_count": 0,
            "cash_amount": 0,
            "upi_count": 0,
            "upi_amount": 0,
            "pending_count": 0,
            "pending_amount": 0,
            "total_amount": 0
        }

        for r in rows:
            tk = r["type_key"]
            if tk not in categories:
                continue
            cat = categories[tk]
            is_void = (r["status"] == "VOID")
            price = int(r["price"] or cat["unit_price"])
            pay = r["payment"]
            ser = r["serial"]

            cat["total_count"] += 1
            if not cat["first_serial"]:
                cat["first_serial"] = ser
            cat["last_serial"] = ser

            is_lucky = (tk in LUCKY_DRAW_KEYS)

            if is_void:
                cat["void_count"] += 1
                if is_lucky:
                    lucky_draw_totals["void_count"] += 1
            else:
                cat["active_count"] += 1
                cat["total_amount"] += price
                if is_lucky:
                    lucky_draw_totals["active_count"] += 1
                    lucky_draw_totals["total_amount"] += price

                if pay == "Cash":
                    cat["cash_count"] += 1
                    cat["cash_amount"] += price
                    if is_lucky:
                        lucky_draw_totals["cash_count"] += 1
                        lucky_draw_totals["cash_amount"] += price
                elif pay == "UPI":
                    cat["upi_count"] += 1
                    cat["upi_amount"] += price
                    if is_lucky:
                        lucky_draw_totals["upi_count"] += 1
                        lucky_draw_totals["upi_amount"] += price
                elif pay in ("Payment Pending", "Pending"):
                    cat["pending_count"] += 1
                    cat["pending_amount"] += price
                    if is_lucky:
                        lucky_draw_totals["pending_count"] += 1
                        lucky_draw_totals["pending_amount"] += price

            if is_lucky:
                lucky_draw_totals["total_count"] += 1

        today_d, today_t, _ = now_parts()
        return jsonify({
            "report_date": target_date or "All Time",
            "counter": counter_filter if user["role"] == "admin" else user["name"],
            "generated_by": user["name"],
            "generated_at": f"{today_d} {today_t}",
            "categories": [categories["RE"], categories["SI"], categories["SA"]],
            "grand_totals": lucky_draw_totals,
            "lucky_draw_categories": [categories["RE"], categories["SI"], categories["SA"]],
            "lucky_draw_totals": lucky_draw_totals,
            "kunkuma_archana": categories["KA"],
            "archana_totals": categories["KA"],
            "combined_all_totals": {
                "active_count": lucky_draw_totals["active_count"] + categories["KA"]["active_count"],
                "total_amount": lucky_draw_totals["total_amount"] + categories["KA"]["total_amount"]
            }
        })
    finally:
        db.close()


@app.post("/api/backup/create")
@require_auth
def manual_backup(user):
    """
    4. Backup & Admin Safeguards:
       Creates an on-demand compressed database backup (.sql.gz)
       and triggers cloud upload if configured.
    """
    res = create_compressed_backup(tag=f"manual_{user['username']}")
    db = connect()
    try:
        log_audit(
            db=db,
            action="BACKUP_MANUAL",
            user_id=user["username"],
            status="SUCCESS",
            details=f"Backup {res['filename']} ({res['size_kb']} KB) created by {user['name']}."
        )
        db.commit()
    finally:
        db.close()
    return jsonify(res)


@app.get("/api/backup/download-latest")
@require_auth
def download_latest_backup(user):
    """
    Downloads the most recent compressed database backup (.sql.gz).
    """
    existing = sorted(
        glob.glob(os.path.join(BACKUP_DIR, "backup_svara_*.sql.gz")),
        key=os.path.getmtime,
        reverse=True
    )
    if not existing:
        # Create one immediately if none exists
        res = create_compressed_backup(tag="on_demand")
        filepath = res["filepath"]
    else:
        filepath = existing[0]

    return send_file(
        filepath,
        mimetype="application/gzip",
        as_attachment=True,
        download_name=os.path.basename(filepath)
    )


@app.get("/api/backup/list")
@require_auth
def list_backups(user):
    """
    Returns list of local backups available.
    """
    files = []
    for fp in sorted(glob.glob(os.path.join(BACKUP_DIR, "backup_svara_*.sql.gz")), key=os.path.getmtime, reverse=True):
        fn = os.path.basename(fp)
        size_kb = round(os.path.getsize(fp) / 1024, 2)
        mtime = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%Y %I:%M:%S %p")
        files.append({"filename": fn, "size_kb": size_kb, "created_at": mtime})
    return jsonify(backups=files, count=len(files))


@app.post("/api/tokens/<serial>/void")
@app.delete("/api/tokens/<serial>")
@require_auth
def void_token(user, serial):
    """
    Admin-only token voiding.
    Marks the token as VOID without rolling back or decrementing the serial counter.
    Future token serials will strictly continue forward.
    """
    if user["role"] != "admin":
        return jsonify(error="Permission denied: Only Administrator can void tokens."), 403

    serial = str(serial).strip().upper()
    db = connect()
    try:
        row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        if not row:
            return jsonify(error=f"Token {serial} was not found."), 404

        current_status = row_val(row, "status", "ACTIVE")
        if current_status == "VOID":
            return jsonify(error=f"Token {serial} is already marked as VOID."), 400

        now_iso = datetime.now().isoformat()
        db.execute(
            "UPDATE tokens SET status = 'VOID', voided_at = ?, voided_by = ? WHERE UPPER(serial) = ?",
            (now_iso, user["username"], serial)
        )
        db.commit()
        # NOTE: counters.last_no is intentionally NEVER decremented!
        refresh_excel_file(db)
    finally:
        db.close()

    return jsonify(
        success=True,
        serial=serial,
        status="VOID",
        message=f"Token {serial} has been VOIDED. The serial sequence remains at its current position."
    )


@app.post("/api/admin/wipe")
def wipe_database():
    """
    4. Admin-Gated 'Wipe Data':
       Protects data reset/wipe behind an Admin Authentication modal
       requiring valid administrator credentials. All wipe attempts (success or failure)
       are securely written to an immutable audit log.
    """
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()
    confirmation = str(data.get("confirm", "")).strip().upper()
    reason = str(data.get("reason", "Administrative Data Reset")).strip()
    client_ip = request.headers.get("X-Forwarded-For", request.remote_addr or "127.0.0.1").split(",")[0].strip()

    db = connect()
    try:
        # Authenticate admin credentials
        user = authenticate_user(username, password)
        if not user or user["role"] != "admin":
            log_audit(
                db=db,
                action="WIPE_DATA_ATTEMPT",
                user_id=username or "unknown",
                status="FAILED",
                details=f"Invalid administrator credentials from IP {client_ip}. Reason given: {reason}",
                ip_address=client_ip
            )
            db.commit()
            return jsonify(error="Administrator authentication failed: Invalid admin User ID or Password."), 401

        if confirmation != "WIPE":
            log_audit(
                db=db,
                action="WIPE_DATA_ATTEMPT",
                user_id=user["username"],
                status="FAILED",
                details=f"Confirmation mismatch ('{confirmation}' != 'WIPE') from IP {client_ip}.",
                ip_address=client_ip
            )
            db.commit()
            return jsonify(error="Confirmation failed: You must type 'WIPE' in capital letters."), 400

        # Safety: Take an automatic pre-wipe compressed backup before deleting data
        try:
            create_compressed_backup(tag="pre_wipe")
        except Exception as bkp_err:
            print(f"[WARN] Pre-wipe automated backup warning: {bkp_err}")

        # Delete all tokens and reset counters
        db.execute("DELETE FROM tokens")
        db.execute("UPDATE counters SET last_no = 0")

        # Write immutable audit log of successful wipe
        log_audit(
            db=db,
            action="WIPE_DATA_SUCCESS",
            user_id=user["username"],
            status="SUCCESS",
            details=f"Database wiped by Admin '{user['name']}'. Reason: {reason}. IP: {client_ip}.",
            ip_address=client_ip
        )
        db.commit()
        refresh_excel_file(db)

        return jsonify(
            success=True,
            message="Database wiped successfully. All token records have been deleted and serial numbers reset to B00001, S00001, SA00001. A secure audit record has been logged."
        )
    except Exception as e:
        db.rollback()
        return jsonify(error=f"Failed to wipe database: {str(e)}"), 500
    finally:
        db.close()


@app.get("/api/system/status")
def system_status():
    db = connect()
    try:
        is_pg = db.is_pg
    finally:
        db.close()

    is_render = bool(os.environ.get("RENDER"))
    if is_pg:
        storage_name = "PostgreSQL Database (Permanent across all deploys)"
        is_persistent = True
    elif is_render:
        storage_name = "Render Container (Ephemeral - click 'Backup Database' to save to laptop)"
        is_persistent = False
    else:
        storage_name = "Permanent Local Laptop Storage (database/svara.db)"
        is_persistent = True

    return jsonify({
        "persistent": is_persistent,
        "storage": storage_name
    })


@app.post("/api/import")
@require_auth
def import_backup(user):
    if user["role"] != "admin":
        return jsonify(error="Only Administrator can restore data."), 403

    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="No file uploaded."), 400

    fn = file.filename.lower()
    if not (fn.endswith(".xlsx") or fn.endswith(".db") or fn.endswith(".sqlite") or fn.endswith(".sqlite3")):
        return jsonify(error="Please upload a valid .xlsx Excel file or .db SQLite backup file."), 400

    db = connect()
    try:
        imported = 0
        max_nums = {"RE": 0, "SI": 0, "SA": 0}
        labels = {"RE": "Royal Enfield", "SI": "Silver", "SA": "Saree"}

        if fn.endswith(".xlsx"):
            try:
                wb = openpyxl.load_workbook(file)
                ws = wb.active
            except Exception as e:
                return jsonify(error=f"Cannot read Excel file: {str(e)}"), 400

            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                return jsonify(error="The uploaded Excel file has no token data rows."), 400

            for row in rows[1:]:
                if not row or not row[0]:
                    continue
                serial = str(row[0]).strip().upper()
                if not (serial.startswith("B") or serial.startswith("S") or serial.startswith("SA")):
                    continue

                token_type = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                status = "ACTIVE"
                if len(row) >= 11:
                    try:
                        price = int(row[2]) if row[2] else 0
                    except (ValueError, TypeError):
                        price = 0
                    status = str(row[3]).strip().upper() if len(row) > 3 and row[3] else "ACTIVE"
                    if status != "VOID":
                        status = "ACTIVE"
                    name = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                    mobile = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                    payment = str(row[6]).strip() if len(row) > 6 and row[6] else "Cash"
                    c_date = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                    c_time = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                    c_user = str(row[9]).strip() if len(row) > 9 and row[9] else "admin"
                    c_name = str(row[10]).strip() if len(row) > 10 and row[10] else "Counter"
                elif len(row) == 10:
                    try:
                        price = int(row[2]) if row[2] else 0
                    except (ValueError, TypeError):
                        price = 0
                    name = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    mobile = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                    payment = str(row[5]).strip() if len(row) > 5 and row[5] else "Cash"
                    c_date = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                    c_time = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                    c_user = str(row[8]).strip() if len(row) > 8 and row[8] else "admin"
                    c_name = str(row[9]).strip() if len(row) > 9 and row[9] else "Counter"
                else:
                    name = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                    mobile = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    payment = str(row[4]).strip() if len(row) > 4 and row[4] else "Cash"
                    c_date = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                    c_time = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                    c_user = "admin"
                    c_name = "Main Counter"
                    price = 0

                if serial.startswith("SA"):
                    type_key = "SA"
                elif serial.startswith("S"):
                    type_key = "SI"
                elif serial.startswith("B"):
                    type_key = "RE"
                else:
                    continue

                if not price:
                    price = PRICES.get(type_key, 0)

                m = re.search(r"\d+", serial)
                if m:
                    n = int(m.group(0))
                    if n > max_nums[type_key]:
                        max_nums[type_key] = n

                if not token_type:
                    token_type = labels.get(type_key, "Token")

                now_iso = datetime.now().isoformat()
                if db.is_pg:
                    db.execute("""
                        INSERT INTO tokens (serial, type_key, token_type, price, status, name, mobile, payment, created_by, counter_name, created_date, created_time, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT (serial) DO NOTHING
                    """, (serial, type_key, token_type, price, status, name, mobile, payment, c_user, c_name, c_date, c_time, now_iso))
                else:
                    db.execute("""
                        INSERT OR IGNORE INTO tokens (serial, type_key, token_type, price, status, name, mobile, payment, created_by, counter_name, created_date, created_time, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (serial, type_key, token_type, price, status, name, mobile, payment, c_user, c_name, c_date, c_time, now_iso))
                imported += 1

        else:
            # Handle uploaded SQLite .db file
            temp_db_path = os.path.join(EXPORT_DIR, f"temp_upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
            file.save(temp_db_path)
            try:
                up_conn = sqlite3.connect(temp_db_path)
                up_conn.row_factory = sqlite3.Row

                up_cur = up_conn.execute("SELECT * FROM tokens")
                for tr in up_cur.fetchall():
                    s = tr["serial"]
                    tk = tr["type_key"]
                    tt = tr["token_type"]
                    pr = row_val(tr, "price", 0)
                    st = row_val(tr, "status", "ACTIVE")
                    va = row_val(tr, "voided_at")
                    vb = row_val(tr, "voided_by")
                    nm = tr["name"]
                    mb = tr["mobile"]
                    pm = tr["payment"]
                    cb = row_val(tr, "created_by", "admin")
                    cn = row_val(tr, "counter_name", "Counter")
                    cd = tr["created_date"]
                    ct = tr["created_time"]
                    ca = tr["created_at"]

                    m = re.search(r"\d+", s)
                    if m and tk in max_nums:
                        n = int(m.group(0))
                        if n > max_nums[tk]:
                            max_nums[tk] = n

                    if db.is_pg:
                        db.execute("""
                            INSERT INTO tokens (serial, type_key, token_type, price, status, voided_at, voided_by, name, mobile, payment, created_by, counter_name, created_date, created_time, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ON CONFLICT (serial) DO NOTHING
                        """, (s, tk, tt, pr, st, va, vb, nm, mb, pm, cb, cn, cd, ct, ca))
                    else:
                        db.execute("""
                            INSERT OR IGNORE INTO tokens (serial, type_key, token_type, price, status, voided_at, voided_by, name, mobile, payment, created_by, counter_name, created_date, created_time, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (s, tk, tt, pr, st, va, vb, nm, mb, pm, cb, cn, cd, ct, ca))
                    imported += 1

                up_cnt = up_conn.execute("SELECT type_key, last_no FROM counters").fetchall()
                for cr in up_cnt:
                    tk = cr["type_key"]
                    ln = cr["last_no"]
                    if tk in max_nums and ln > max_nums[tk]:
                        max_nums[tk] = ln
                up_conn.close()
            finally:
                if os.path.exists(temp_db_path):
                    try:
                        os.remove(temp_db_path)
                    except OSError:
                        pass

        for tk, max_n in max_nums.items():
            if max_n > 0:
                db.execute("UPDATE counters SET last_no = ? WHERE type_key = ? AND last_no < ?", (max_n, tk, max_n))

        db.commit()
        refresh_excel_file(db)
        return jsonify(success=True, imported=imported, message=f"Successfully restored {imported} records! Counters updated to highest sequence numbers.")
    except Exception as e:
        db.rollback()
        return jsonify(error=f"Restore failed: {str(e)}"), 500
    finally:
        db.close()


@app.get("/api/export")
@require_auth
def export_excel(user):
    counter_filter = request.args.get("counter", "").strip().lower()
    category = request.args.get("category", "").strip().lower()
    db = connect()
    try:
        if category in ("archana", "kunkuma archana", "ka"):
            rows = all_rows(db, user=user, category="archana", newest_first=False)
            file_name = "SVARA_2026_Kunkuma_Archana.xlsx"
            sheet_title = "Kunkuma Archana"
        elif category == "all":
            rows = all_rows(db, user=user, counter_filter=counter_filter, newest_first=False)
            file_name = "SVARA_2026_All_Tokens.xlsx"
            sheet_title = "All Tokens"
        else:
            # Default: strictly Lucky Draw tokens (RE, SI, SA)
            rows = all_rows(db, user=user, counter_filter=counter_filter, category="luckydraw", newest_first=False)
            if user["role"] == "admin":
                if counter_filter and counter_filter not in ("all", "pending", "archana"):
                    file_name = f"SVARA_2026_Tokens_{counter_filter.capitalize()}.xlsx"
                    sheet_title = f"{counter_filter.capitalize()} Tokens"
                else:
                    file_name = "SVARA_2026_Tokens_All.xlsx"
                    sheet_title = "Lucky Draw Tokens"
            else:
                file_name = f"SVARA_2026_Tokens_{user['username'].capitalize()}.xlsx"
                sheet_title = f"{user['name']} Tokens"

        wb = build_workbook(rows, title=sheet_title)
    finally:
        db.close()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=file_name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/export/kunkuma-archana")
@require_auth
def export_kunkuma_archana(user):
    db = connect()
    try:
        rows = all_rows(db, user=user, category="archana", newest_first=False)
        wb = build_workbook(rows, title="Kunkuma Archana")
    finally:
        db.close()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="SVARA_2026_Kunkuma_Archana.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/backup/db")
@require_auth
def backup_db(user):
    """
    Downloads the active database as a SQLite .db file directly to the user's laptop.
    Ensures data is permanently stored on their computer for future use and safekeeping.
    """
    db = connect()
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"svara_database_{timestamp}.db"

        if not db.is_pg:
            if os.path.exists(DB_PATH):
                return send_file(DB_PATH, as_attachment=True, download_name=filename, mimetype="application/x-sqlite3")

        # When running on PostgreSQL (Render cloud), create a SQLite backup file to stream to laptop
        backup_path = os.path.join(EXPORT_DIR, filename)
        backup_conn = sqlite3.connect(backup_path)
        with open(os.path.join(ROOT, "database", "schema.sql"), encoding="utf-8") as f:
            backup_conn.executescript(f.read())

        counters_rows = db.execute("SELECT * FROM counters").fetchall()
        for c in counters_rows:
            backup_conn.execute(
                "INSERT OR REPLACE INTO counters (type_key, label, prefix, last_no) VALUES (?, ?, ?, ?)",
                (c["type_key"], c["label"], c["prefix"], c["last_no"])
            )

        tokens_rows = db.execute("SELECT * FROM tokens ORDER BY id ASC").fetchall()
        for t in tokens_rows:
            backup_conn.execute("""
                INSERT OR REPLACE INTO tokens (
                    id, serial, type_key, token_type, price, status, voided_at, voided_by,
                    name, mobile, payment, created_by, counter_name, created_date, created_time, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                t["id"], t["serial"], t["type_key"], t["token_type"], row_val(t, "price", 0),
                row_val(t, "status", "ACTIVE"), row_val(t, "voided_at"), row_val(t, "voided_by"),
                t["name"], t["mobile"], t["payment"], row_val(t, "created_by", "admin"),
                row_val(t, "counter_name", "Main Counter"), t["created_date"], t["created_time"], t["created_at"]
            ))
        backup_conn.commit()
        backup_conn.close()

        return send_file(backup_path, as_attachment=True, download_name=filename, mimetype="application/x-sqlite3")
    finally:
        db.close()


@app.get("/logo.jpg")

def serve_logo():
    return send_from_directory(FRONTEND, "logo.jpg", mimetype="image/jpeg")


@app.get("/")
def index():
    return send_from_directory(FRONTEND, "index.html")


init_db()
start_backup_scheduler()

if __name__ == "__main__":
    host = os.environ.get("SVARA_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("SVARA_PORT", "5000")))
    print(f"\n  SVARA token system running:  http://localhost:{port}\n")
    app.run(host=host, port=port, threaded=True)
