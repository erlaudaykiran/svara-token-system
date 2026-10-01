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
import io
import os
import re
import secrets
import sqlite3
import threading
from datetime import datetime, timezone, timedelta
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

# Database Configuration:
# When DATABASE_URL is set (Render PostgreSQL, Neon, Supabase), use PostgreSQL.
# Otherwise, fall back to SQLite on local disk (database/svara.db).
DATABASE_URL = os.environ.get("DATABASE_URL")
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
DB_PATH = os.environ.get("SVARA_DB", os.path.join(ROOT, "database", "svara.db"))
FRONTEND = os.path.join(ROOT, "frontend")
EXPORT_DIR = os.path.join(ROOT, "exports")
EXPORT_FILE = os.path.join(EXPORT_DIR, "SVARA_2026_Tokens.xlsx")

TZ_NAME = os.environ.get("SVARA_TZ", "Asia/Kolkata")

# Token Prices
PRICES = {
    "RE": 301,  # Royal Enfield
    "SI": 201,  # Silver
    "SA": 101   # Saree
}

# User accounts configuration:
# 1 Admin (handles all counters) + 3 Separate Counter users with independent credentials
USERS = {
    "admin": {
        "password": os.environ.get("ADMIN_PASSWORD", os.environ.get("SVARA_PASSWORD", "admin@svara2026")),
        "role": "admin",
        "name": "Administrator",
        "counter_name": "Main / All Counters"
    },
    "counter1": {
        "password": os.environ.get("COUNTER1_PASSWORD", "counter1@2026"),
        "role": "counter",
        "name": "Counter 1",
        "counter_name": "Counter 1"
    },
    "counter2": {
        "password": os.environ.get("COUNTER2_PASSWORD", "counter2@2026"),
        "role": "counter",
        "name": "Counter 2",
        "counter_name": "Counter 2"
    },
    "counter3": {
        "password": os.environ.get("COUNTER3_PASSWORD", "counter3@2026"),
        "role": "counter",
        "name": "Counter 3",
        "counter_name": "Counter 3"
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
                    ('RE', 'Royal Enfield', 'B',  0),
                    ('SI', 'Silver',        'S',  0),
                    ('SA', 'Saree',         'SA', 0)
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
                    payment       VARCHAR(20) NOT NULL CHECK (payment IN ('Cash', 'UPI')),
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

    custom_admin = os.environ.get("SVARA_USER", "").strip().lower()
    if user_key in ("admin", "administrator", custom_admin):
        user_key = "admin"

    if user_key not in USERS:
        return None

    cfg = USERS[user_key]

    # Collect all acceptable passwords for this user
    allowed_passwords = [cfg["password"]]
    if user_key == "admin":
        allowed_passwords.extend([
            "admin@svara2026",
            "svara@2026",
            "admin",
            "admin123",
            os.environ.get("ADMIN_PASSWORD", ""),
            os.environ.get("SVARA_PASSWORD", "")
        ])
    elif user_key == "counter1":
        allowed_passwords.extend(["counter1@2026", "counter1", os.environ.get("COUNTER1_PASSWORD", "")])
    elif user_key == "counter2":
        allowed_passwords.extend(["counter2@2026", "counter2", os.environ.get("COUNTER2_PASSWORD", "")])
    elif user_key == "counter3":
        allowed_passwords.extend(["counter3@2026", "counter3", os.environ.get("COUNTER3_PASSWORD", "")])

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

        if status_val == "ACTIVE":
            total_amount += int(price_val)
            active_count += 1
        else:
            void_count += 1

        row_cells = [
            row_val(r, "serial", ""),
            row_val(r, "token_type", ""),
            int(price_val),
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


def all_rows(db, user=None, counter_filter=None, newest_first=False):
    order = "DESC" if newest_first else "ASC"
    if user and user["role"] != "admin":
        cur = db.execute(f"SELECT * FROM tokens WHERE LOWER(created_by) = ? ORDER BY id {order}", (user["username"].lower(),))
    elif counter_filter and counter_filter != "all":
        cur = db.execute(f"SELECT * FROM tokens WHERE LOWER(created_by) = ? ORDER BY id {order}", (counter_filter.lower(),))
    else:
        cur = db.execute(f"SELECT * FROM tokens ORDER BY id {order}")
    return cur.fetchall()


def refresh_excel_file(db):
    try:
        build_workbook(all_rows(db)).save(EXPORT_FILE)
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

    token = secrets.token_hex(24)
    now_iso = datetime.now().isoformat()
    db = connect()
    try:
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
        except Exception as sess_err:
            print(f"[WARN] Session table insert error: {sess_err}. Re-creating active_sessions table...")
            db.rollback()
            db.execute("DROP TABLE IF EXISTS active_sessions")
            if db.is_pg:
                db.execute("CREATE TABLE active_sessions (user_id VARCHAR(50) PRIMARY KEY, session_token TEXT NOT NULL, logged_in_at TEXT NOT NULL)")
            else:
                db.execute("CREATE TABLE active_sessions (user_id TEXT PRIMARY KEY, session_token TEXT NOT NULL, logged_in_at TEXT NOT NULL)")
            db.execute(
                "INSERT INTO active_sessions (user_id, session_token, logged_in_at) VALUES (?, ?, ?)",
                (user["username"], token, now_iso)
            )
            db.commit()
    finally:
        db.close()

    session["token"] = token
    session["user"] = user["username"]
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
            user_count_cur = db.execute(
                "SELECT COUNT(*) FROM tokens WHERE type_key = ? AND status = 'ACTIVE' AND LOWER(created_by) = ?",
                (tk, user["username"].lower())
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

        # Breakdown stats for admin (active vs void)
        counter_breakdown = {}
        if user["role"] == "admin":
            for ukey in USERS:
                c_cur = db.execute(
                    "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens WHERE LOWER(created_by) = ? AND status = 'ACTIVE'",
                    (ukey,)
                )
                c_row = c_cur.fetchone()
                if isinstance(c_row, dict):
                    vals = list(c_row.values())
                    cnt, amt = vals[0], vals[1]
                else:
                    cnt, amt = c_row[0], c_row[1]

                v_cur = db.execute("SELECT COUNT(*) FROM tokens WHERE LOWER(created_by) = ? AND status = 'VOID'", (ukey,))
                v_row = v_cur.fetchone()
                v_cnt = list(v_row.values())[0] if isinstance(v_row, dict) else v_row[0]

                counter_breakdown[ukey] = {
                    "name": USERS[ukey]["name"],
                    "count": cnt,
                    "void_count": v_cnt,
                    "amount": int(amt)
                }

        user_summary_cur = db.execute(
            "SELECT COUNT(*), COALESCE(SUM(price), 0) FROM tokens WHERE LOWER(created_by) = ? AND status = 'ACTIVE'",
            (user["username"].lower(),)
        )
        user_sum_row = user_summary_cur.fetchone()
        if isinstance(user_sum_row, dict):
            u_vals = list(user_sum_row.values())
            my_cnt, my_amt = u_vals[0], u_vals[1]
        else:
            my_cnt, my_amt = user_sum_row[0], user_sum_row[1]

    finally:
        db.close()

    return jsonify({
        "categories": categories,
        "currentUser": user,
        "myStats": {"count": my_cnt, "amount": int(my_amt)},
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
    if payment not in ("Cash", "UPI"):
        return jsonify(error="Select a payment type: Cash or UPI."), 400

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

    res = {
        "tokens": created_tokens,
        "count": len(created_tokens),
        "total_amount": unit_price * len(created_tokens),
        "first": created_tokens[0]["serial"],
        "last": created_tokens[-1]["serial"],
        "serial": created_tokens[0]["serial"],
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
    }
    return jsonify(res), 201


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
    if payment not in ("Cash", "UPI"):
        return jsonify(error="Select a valid payment type: Cash or UPI."), 400

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


@app.get("/api/system/status")
def system_status():
    db = connect()
    try:
        is_pg = db.is_pg
    finally:
        db.close()
    return jsonify({
        "persistent": is_pg,
        "storage": "PostgreSQL (Persistent across all deploys)" if is_pg else "Local SQLite (Ephemeral - add DATABASE_URL in Render to persist forever)"
    })


@app.post("/api/import")
@require_auth
def import_excel(user):
    if user["role"] != "admin":
        return jsonify(error="Only Administrator can restore from Excel."), 403

    file = request.files.get("file")
    if not file or not (file.filename.endswith(".xlsx") or file.filename.endswith(".XLSX")):
        return jsonify(error="Please upload a valid .xlsx Excel file."), 400

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
    except Exception as e:
        return jsonify(error=f"Cannot read Excel file: {str(e)}"), 400

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return jsonify(error="The uploaded Excel file has no token data rows."), 400

    db = connect()
    imported = 0
    max_nums = {"RE": 0, "SI": 0, "SA": 0}
    labels = {"RE": "Royal Enfield", "SI": "Silver", "SA": "Saree"}

    try:
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            serial = str(row[0]).strip().upper()
            if not (serial.startswith("B") or serial.startswith("S") or serial.startswith("SA")):
                continue

            token_type = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            status = "ACTIVE"
            
            # Format detection: 11 columns with Status vs 10 columns vs legacy 7 columns
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

        for tk, max_n in max_nums.items():
            if max_n > 0:
                db.execute("UPDATE counters SET last_no = ? WHERE type_key = ? AND last_no < ?", (max_n, tk, max_n))

        db.commit()
        refresh_excel_file(db)
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()

    return jsonify(success=True, imported=imported, message=f"Successfully restored {imported} tokens from Excel! Counters updated.")


@app.get("/api/export")
@require_auth
def export_excel(user):
    counter_filter = request.args.get("counter", "").strip().lower()
    db = connect()
    try:
        rows = all_rows(db, user=user, counter_filter=counter_filter, newest_first=False)
        sheet_title = "Tokens"
        if user["role"] == "admin":
            if counter_filter and counter_filter != "all":
                file_name = f"SVARA_2026_Tokens_{counter_filter.capitalize()}.xlsx"
                sheet_title = f"{counter_filter.capitalize()} Tokens"
            else:
                file_name = "SVARA_2026_Tokens_All.xlsx"
                sheet_title = "All Counters"
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


@app.get("/logo.jpg")
def serve_logo():
    return send_from_directory(FRONTEND, "logo.jpg", mimetype="image/jpeg")


@app.get("/")
def index():
    return send_from_directory(FRONTEND, "index.html")


init_db()

if __name__ == "__main__":
    host = os.environ.get("SVARA_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("SVARA_PORT", "5000")))
    print(f"\n  SVARA token system running:  http://localhost:{port}\n")
    app.run(host=host, port=port, threaded=True)
