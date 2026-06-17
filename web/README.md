# Kareem T-01 — Web version (Frontend + Backend)

```
web/
  backend/api.py        ← FastAPI server (wraps the engine process_contracts.py)
  frontend/index.html   ← the UI (gold/green, sidebar, cards, table, progress, footer)
```

The backend reuses the SAME engine as the desktop app: `process_contracts.py`, `refs.json`,
`recipients.json`, `config.json` (keep them in the `_tool` folder, or copy next to `api.py`).

## Run (Windows, on a machine where running is allowed)
1. Install Python 3.13 + Tesseract OCR (with the `ara` language pack).
2. Install packages:
   ```
   py -m pip install fastapi uvicorn PyMuPDF pytesseract Pillow openpyxl numpy pywin32
   ```
3. Start the server from the backend folder:
   ```
   cd web\backend
   py -m uvicorn api:app --port 8000
   ```
4. Open `http://localhost:8000` in the browser → the UI loads and talks to the API.

## API endpoints
- `GET  /api/inbox`        → files in Inbox + count + sync folder
- `POST /api/process`      → start processing (background)
- `GET  /api/status`       → live progress + per-file results
- `GET  /api/recipients`   / `POST /api/recipients`  → read / save project To/CC
- `GET  /api/pending`      → files needing a manual [s-e]
- `GET  /api/history`      → separated contracts in Output

## Deploy elsewhere (Lovable / Cursor / any platform)
Use `Kareem T-01 - Build Prompt.md` as the full spec, or host this FastAPI backend and
serve `index.html` as the frontend. Keep contract PDFs local/private — do not upload them
to third-party servers without IT approval.

Footer on every screen: **Created By : DPC Department / By Kareem Talaat**
