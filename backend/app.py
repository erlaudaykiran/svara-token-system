"""
SVARA 2026 Lucky Draw - token printing backend.

Flask + SQLite. Serves the frontend and a small JSON API:
  GET  /api/counters      next serial + issued count per category
  POST /api/tokens        create a token (serial is assigned atomically)
  GET  /api/tokens        list all tokens (newest first)
  GET  /api/export        download all tokens as an .xlsx file
Every saved token also refreshes exports/SVARA_2026_Tokens.xlsx.
"""
import io
import os
import re
import sqlite3
from datetime import datetime

from flask import Flask, jsonify, request, send_file, send_from_directory
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
DB_PATH = os.environ.get("SVARA_DB", os.path.join(ROOT, "database", "svara.db"))
SCHEMA = os.path.join(ROOT, "database", "schema.sql")
FRONTEND = os.path.join(ROOT, "frontend")
EXPORT_DIR = os.path.join(ROOT, "exports")
EXPORT_FILE = os.path.join(EXPORT_DIR, "SVARA_2026_Tokens.xlsx")

app = Flask(__name__, static_folder=None)


def connect():
    conn = sqlite3.connect(DB_PATH, timeout=10, isolation_level=None)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    conn = connect()
    try:
        with open(SCHEMA, encoding="utf-8") as f:
            conn.executescript(f.read())
    finally:
        conn.close()


def fmt_serial(prefix, n):
    return f"{prefix}{n:03d}"


def now_parts():
    d = datetime.now()
    return d.strftime("%d/%m/%Y"), d.strftime("%I:%M:%S %p"), d.isoformat(timespec="seconds")


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
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):  # keep mobile as text
        row[0].number_format = "@"
    ws.freeze_panes = "A2"
    return wb


def all_rows(conn, newest_first=False):
    order = "DESC" if newest_first else "ASC"
    return conn.execute(f"SELECT * FROM tokens ORDER BY id {order}").fetchall()


def refresh_excel_file(conn):
    """Best effort: fails quietly if the file is open in Excel."""
    try:
        build_workbook(all_rows(conn)).save(EXPORT_FILE)
    except OSError:
        pass


def token_json(r):
    return {"serial": r["serial"], "type": r["token_type"], "name": r["name"],
            "mobile": r["mobile"], "payment": r["payment"],
            "date": r["created_date"], "time": r["created_time"]}


@app.get("/api/counters")
def counters():
    conn = connect()
    try:
        out = {}
        for c in conn.execute("SELECT * FROM counters"):
            out[c["type_key"]] = {"label": c["label"], "issued": c["last_no"],
                                  "next": fmt_serial(c["prefix"], c["last_no"] + 1)}
    finally:
        conn.close()
    return jsonify(out)


@app.post("/api/tokens")
def create_token():
    data = request.get_json(silent=True) or {}
    type_key = str(data.get("type", "")).strip()
    name = re.sub(r"\s+", " ", str(data.get("name", ""))).strip()
    mobile = str(data.get("mobile", "")).strip()
    payment = str(data.get("payment", "")).strip()

    if not name or len(name) > 60:
        return jsonify(error="Enter the customer name (max 60 characters)."), 400
    if not re.fullmatch(r"\d{10}", mobile):
        return jsonify(error="Enter a valid 10-digit mobile number."), 400
    if payment not in ("Cash", "UPI"):
        return jsonify(error="Select a payment type: Cash or UPI."), 400

    conn = connect()
    try:
        conn.execute("BEGIN IMMEDIATE")  # lock so two counters never get the same serial
        try:
            c = conn.execute("SELECT * FROM counters WHERE type_key = ?", (type_key,)).fetchone()
            if c is None:
                conn.execute("ROLLBACK")
                return jsonify(error="Unknown token type."), 400
            n = c["last_no"] + 1
            serial = fmt_serial(c["prefix"], n)
            date, time_, iso = now_parts()
            conn.execute("UPDATE counters SET last_no = ? WHERE type_key = ?", (n, type_key))
            conn.execute(
                "INSERT INTO tokens (serial, type_key, token_type, name, mobile, payment,"
                " created_date, created_time, created_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (serial, type_key, c["label"], name, mobile, payment, date, time_, iso))
            conn.execute("COMMIT")
        except Exception:
            if conn.in_transaction:
                conn.execute("ROLLBACK")
            raise
        row = conn.execute("SELECT * FROM tokens WHERE serial = ?", (serial,)).fetchone()
        refresh_excel_file(conn)
    finally:
        conn.close()
    return jsonify(token_json(row)), 201


@app.get("/api/tokens")
def list_tokens():
    conn = connect()
    try:
        return jsonify([token_json(r) for r in all_rows(conn, newest_first=True)])
    finally:
        conn.close()


@app.get("/api/export")
def export_excel():
    conn = connect()
    try:
        wb = build_workbook(all_rows(conn))
    finally:
        conn.close()
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
    host = os.environ.get("SVARA_HOST", "0.0.0.0")  # 0.0.0.0 lets other PCs on the network open it
    port = int(os.environ.get("PORT", os.environ.get("SVARA_PORT", "5000")))
    print(f"\n  SVARA token system running:  http://localhost:{port}\n")
    app.run(host=host, port=port, threaded=True)
