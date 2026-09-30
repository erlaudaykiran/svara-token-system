# SVARA 2026 Lucky Draw - Token Printing System

Frontend + backend + database. Tokens for Royal Enfield (B001...), Silver (S001...)
and Saree (SA001...). Each entry stores name, mobile and payment type (Cash / UPI),

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/erlaudaykiran/svara-token-system)


## Folders
    frontend/   index.html          the counter screen (HTML/CSS/JS, no build step)
    backend/    app.py              Flask API + serves the frontend
                requirements.txt    Flask, openpyxl
    database/   schema.sql          table definitions
                svara.db            SQLite file, created automatically on first run
    exports/    SVARA_2026_Tokens.xlsx   refreshed automatically after every token
    run.bat / run.sh                one-click start (Windows / Linux, Mac)

## Run it
1. Install Python 3.9+ (Windows: tick "Add Python to PATH").
2. Windows: double-click run.bat.   Linux/Mac: ./run.sh
   The first run installs Flask and openpyxl (needs internet once), then opens http://localhost:5000
3. Use it: pick a token type -> enter name, mobile, payment -> Save & print.

To use it from other computers on the same Wi-Fi/LAN, open
http://<server-PC-IP>:5000 on them. All counters then share one database,
and serial numbers never repeat (each save is one locked database transaction).

Optional environment variables: SVARA_PORT (default 5000),
SVARA_HOST (default 0.0.0.0, use 127.0.0.1 to allow this PC only), SVARA_DB (database path).

## Excel
- exports/SVARA_2026_Tokens.xlsx is rewritten after every saved token.
  If the file is open in Excel on Windows it may be locked and skip that refresh,
  so use the button below or close the file.
- "Download Excel" on the page always gives a fresh file from the database.
- Columns: Token No, Type, Name, Mobile, Payment, Date, Time.

## Printing
Tokens are 76 mm wide (80 mm thermal printer). For A4, choose "Fit to page" in the print dialog.
For silent one-click printing, start Chrome/Edge with the --kiosk-printing flag.

## Backup / reset
- Backup: copy database/svara.db while the server is stopped.
- New series from B001/S001/SA001: stop the server and delete database/svara.db
  (and svara.db-wal / svara.db-shm if present). Download the Excel first.

## API
    GET  /api/counters   next serial and issued count per type
    POST /api/tokens     {"type":"RE|SI|SA","name":"..","mobile":"10 digits","payment":"Cash|UPI"}
    GET  /api/tokens     all tokens, newest first
    GET  /api/export     Excel download
