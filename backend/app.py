"""
SVARA 2026 Lucky Draw - token printing backend.

Flask + SQLite / PostgreSQL. Serves the frontend and a JSON API:
  POST   /api/auth/login    authenticate staff (single user & session)
  POST   /api/auth/logout   end active session
  GET    /api/auth/status   check if current user is logged in
  GET    /api/counters      next serial + issued count per category
  POST   /api/tokens        create a batch of tokens in serial sequence
  DELETE /api/tokens/<id>   cancel and delete a token
  GET    /api/tokens        list all tokens (newest first)
  GET    /api/export        download all tokens as an .xlsx file
"""
import io
import os
import re
import secrets
import sqlite3
from datetime import datetime, timezone, timedelta
try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo
from functools import wraps

from flask import Flask, jsonify, request, send_file, send_from_directory, session
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

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

AUTH_USER = os.environ.get("SVARA_USER", "admin")
AUTH_PASS = os.environ.get("SVARA_PASSWORD", "svara@2026")
TZ_NAME = os.environ.get("SVARA_TZ", "Asia/Kolkata")

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
    if DATABASE_URL:
        import psycopg2
        import psycopg2.extras
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
        conn.autocommit = False
        return DBWrapper(conn, is_pg=True)
    else:
        conn = sqlite3.connect(DB_PATH, timeout=10, isolation_level=None)
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
                    name          VARCHAR(100) NOT NULL,
                    mobile        VARCHAR(20) NOT NULL,
                    payment       VARCHAR(20) NOT NULL CHECK (payment IN ('Cash', 'UPI')),
                    created_date  VARCHAR(20) NOT NULL,
                    created_time  VARCHAR(20) NOT NULL,
                    created_at    VARCHAR(40) NOT NULL
                );
            """)
            db.execute("CREATE INDEX IF NOT EXISTS idx_tokens_type ON tokens(type_key);")
            db.execute("""
                CREATE TABLE IF NOT EXISTS active_sessions (
                    id            INTEGER PRIMARY KEY CHECK (id = 1),
                    session_token TEXT NOT NULL,
                    logged_in_at  TEXT NOT NULL
                );
            """)
            db.commit()
        else:
            with open(os.path.join(ROOT, "database", "schema.sql"), encoding="utf-8") as f:
                db.conn.executescript(f.read())
    finally:
        db.close()


def get_tz():
    try:
        return ZoneInfo(TZ_NAME)
    except Exception:
        return timezone(timedelta(hours=5, minutes=30))


def now_parts():
    d = datetime.now(get_tz())
    return d.strftime("%d/%m/%Y"), d.strftime("%I:%M:%S %p"), d.isoformat(timespec="seconds")


def is_authenticated():
    token = session.get("token")
    if not token:
        return False
    db = connect()
    try:
        cur = db.execute("SELECT session_token FROM active_sessions WHERE id = 1")
        row = cur.fetchone()
        return bool(row and secrets.compare_digest(row["session_token"], token))
    finally:
        db.close()


def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_authenticated():
            return jsonify(error="Authentication required. Please sign in.", auth=False), 401
        return f(*args, **kwargs)
    return decorated


def fmt_serial(prefix, n):
    return f"{prefix}{n:03d}"


COLUMNS = ["Token No", "Type", "Name", "Mobile", "Payment", "Date", "Time"]


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tokens"
    ws.append(COLUMNS)
    for c in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7A1424")
    for r in rows:
        ws.append([r["serial"], r["token_type"], r["name"], r["mobile"],
                   r["payment"], r["created_date"], r["created_time"]])
    for i, w in enumerate([11, 16, 26, 14, 10, 12, 13], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].number_format = "@"
    ws.freeze_panes = "A2"
    return wb


def all_rows(db, newest_first=False):
    order = "DESC" if newest_first else "ASC"
    cur = db.execute(f"SELECT * FROM tokens ORDER BY id {order}")
    return cur.fetchall()


def refresh_excel_file(db):
    try:
        build_workbook(all_rows(db)).save(EXPORT_FILE)
    except OSError:
        pass


def token_json(r):
    return {"serial": r["serial"], "type": r["token_type"], "name": r["name"],
            "mobile": r["mobile"], "payment": r["payment"],
            "date": r["created_date"], "time": r["created_time"]}


@app.post("/api/auth/login")
def login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not (secrets.compare_digest(username, AUTH_USER) and secrets.compare_digest(password, AUTH_PASS)):
        return jsonify(error="Invalid User ID or Password."), 401

    token = secrets.token_hex(24)
    now_iso = datetime.now().isoformat()
    db = connect()
    try:
        db.execute(
            "INSERT INTO active_sessions (id, session_token, logged_in_at) VALUES (1, ?, ?) "
            "ON CONFLICT (id) DO UPDATE SET session_token = EXCLUDED.session_token, logged_in_at = EXCLUDED.logged_in_at",
            (token, now_iso)
        )
        db.commit()
    finally:
        db.close()

    session["token"] = token
    return jsonify(success=True, user=AUTH_USER)


@app.post("/api/auth/logout")
def logout():
    token = session.get("token")
    if token:
        db = connect()
        try:
            db.execute("DELETE FROM active_sessions WHERE id = 1 AND session_token = ?", (token,))
            db.commit()
        finally:
            db.close()
    session.clear()
    return jsonify(success=True)


@app.get("/api/auth/status")
def auth_status():
    if is_authenticated():
        return jsonify(authenticated=True, user=AUTH_USER)
    return jsonify(authenticated=False)


@app.get("/api/counters")
@require_auth
def counters():
    db = connect()
    try:
        out = {}
        for c in db.execute("SELECT * FROM counters").fetchall():
            count_cur = db.execute("SELECT COUNT(*) FROM tokens WHERE type_key = ?", (c["type_key"],))
            count_row = count_cur.fetchone()
            if isinstance(count_row, dict):
                count = list(count_row.values())[0]
            else:
                count = count_row[0]
            out[c["type_key"]] = {
                "label": c["label"],
                "issued": count,
                "next": fmt_serial(c["prefix"], c["last_no"] + 1)
            }
    finally:
        db.close()
    return jsonify(out)


@app.post("/api/tokens")
@require_auth
def create_token():
    data = request.get_json(silent=True) or {}
    type_key = str(data.get("type", "")).strip()
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

        start_no = c["last_no"]
        end_no = start_no + quantity
        date, time_, iso = now_parts()
        created_tokens = []

        for n in range(start_no + 1, end_no + 1):
            serial = fmt_serial(c["prefix"], n)
            db.execute(
                "INSERT INTO tokens (serial, type_key, token_type, name, mobile, payment,"
                " created_date, created_time, created_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (serial, type_key, c["label"], name, mobile, payment, date, time_, iso))
            created_tokens.append({
                "serial": serial, "type": c["label"], "name": name,
                "mobile": mobile, "payment": payment,
                "date": date, "time": time_
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
        "first": created_tokens[0]["serial"],
        "last": created_tokens[-1]["serial"],
        "serial": created_tokens[0]["serial"],
        "type": created_tokens[0]["type"],
        "name": created_tokens[0]["name"],
        "mobile": created_tokens[0]["mobile"],
        "payment": created_tokens[0]["payment"],
        "date": created_tokens[0]["date"],
        "time": created_tokens[0]["time"],
    }
    return jsonify(res), 201


@app.get("/api/tokens")
@require_auth
def list_tokens():
    db = connect()
    try:
        return jsonify([token_json(r) for r in all_rows(db, newest_first=True)])
    finally:
        db.close()


@app.delete("/api/tokens/<serial>")
@require_auth
def delete_token(serial):
    serial = str(serial).strip().upper()
    db = connect()
    try:
        row = db.execute("SELECT * FROM tokens WHERE UPPER(serial) = ?", (serial,)).fetchone()
        if not row:
            return jsonify(error=f"Token {serial} was not found."), 404
        db.execute("DELETE FROM tokens WHERE UPPER(serial) = ?", (serial,))
        db.commit()
        refresh_excel_file(db)
    finally:
        db.close()
    return jsonify(success=True, deleted=serial, message=f"Token {serial} cancelled and removed from database and Excel.")


@app.get("/api/export")
@require_auth
def export_excel():
    db = connect()
    try:
        wb = build_workbook(all_rows(db))
    finally:
        db.close()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="SVARA_2026_Tokens.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/")
def index():
    return send_from_directory(FRONTEND, "index.html")


init_db()

if __name__ == "__main__":
    host = os.environ.get("SVARA_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("SVARA_PORT", "5000")))
    print(f"\n  SVARA token system running:  http://localhost:{port}\n")
    app.run(host=host, port=port, threaded=True)
