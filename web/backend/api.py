# -*- coding: utf-8 -*-
"""
Kareem T-01 — EDECS Contract Processor : Backend API (FastAPI, local/offline).

Run:
    pip install fastapi uvicorn python-multipart
    uvicorn api:app --port 8000
Then open http://127.0.0.1:8000/  (the frontend is served from /).

Wraps the processing engine (process_contracts.py) and the OCR module (ocr_engine.py).
Everything is local — no contract data leaves the machine.
"""
import os, sys, glob, json, shutil, threading, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("KT01_ROOT") or HERE
# allow importing the engine from this folder or the parent app folder
for p in (HERE, os.path.abspath(os.path.join(HERE, "..", "..")), os.path.abspath(os.path.join(HERE, ".."))):
    if os.path.exists(os.path.join(p, "process_contracts.py")) and p not in sys.path:
        sys.path.insert(0, p)

import process_contracts as pc
import ocr_engine as ocr

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Kareem T-01 — EDECS Contract Processor")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

STATE = {"running": False, "rows": [], "done": 0, "total": 0}
CFG_PATH = os.path.join(pc.HERE, "config.json")

def _ensure_dirs():
    os.makedirs(pc.INBOX, exist_ok=True)
    os.makedirs(pc.OUTPUT, exist_ok=True)

# --------------------------------------------------------------------------
# Inbox / upload
# --------------------------------------------------------------------------
@app.get("/api/inbox")
def inbox():
    _ensure_dirs()
    files = [os.path.basename(f) for f in sorted(glob.glob(os.path.join(pc.INBOX, "*.pdf")))]
    return {"count": len(files), "files": files, "sync": pc.SYNC, "inbox": pc.INBOX,
            "output": pc.OUTPUT, "languages": list(ocr.LANGUAGE_OPTIONS.keys()),
            "default_language": pc.CFG.get("ocr_default_language", ocr.DEFAULT_LANGUAGE)}

@app.get("/api/parse_preview")
def parse_preview():
    """Deterministic filename pre-processing for every PDF in the Inbox (no OCR, no AI)."""
    _ensure_dirs()
    rows = []
    for f in sorted(glob.glob(os.path.join(pc.INBOX, "*.pdf"))):
        base = os.path.basename(f)
        meta = pc.parse_bundle_name(os.path.splitext(base)[0])
        if meta:
            rows.append({"original": base, "code": meta["code"], "no": meta["no"],
                         "annex": meta["annex"],
                         "type": ("Addendum %d" % meta["annex"] if meta["annex"] else "Contract"),
                         "annex_label": pc.annex_label(meta["annex"]),
                         "status": "Parsed"})
        else:
            rows.append({"original": base, "code": "", "no": "", "annex": None,
                         "type": "", "annex_label": "", "status": "Needs Manual Review"})
    return {"rows": rows, "examples": pc.PARSE_EXAMPLES}


@app.post("/api/upload")
async def upload(files: list[UploadFile] = File(...)):
    _ensure_dirs()
    saved = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            continue
        dest = os.path.join(pc.INBOX, os.path.basename(f.filename))
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(os.path.basename(f.filename))
    return {"saved": saved, "count": len(saved)}

# --------------------------------------------------------------------------
# Processing  (batch + per-file)
# --------------------------------------------------------------------------
@app.post("/api/process")
async def process(req: Request):
    if STATE["running"]:
        return {"running": True}
    try:
        body = await req.json()
    except Exception:
        body = {}
    lang = ocr.LANGUAGE_OPTIONS.get(body.get("language", ""), body.get("language")) \
           or pc.CFG.get("ocr_default_language")
    man_date = body.get("agreement_date", "") or ""
    man_sr = body.get("sr_number", "") or ""

    STATE.update(running=True, rows=[], done=0, total=0)
    files = sorted(glob.glob(os.path.join(pc.INBOX, "*.pdf")))
    STATE["total"] = len(files)

    def worker():
        try:
            for i, fp in enumerate(files, 1):
                try:
                    r = pc.process_one(fp, lang=lang,
                                       manual_agreement_date=man_date, manual_sr_number=man_sr)
                except Exception as e:
                    r = {"file": os.path.basename(fp), "status": "failed", "msg": str(e),
                         "confidence": None, "missing": [], "warnings": []}
                STATE["rows"].append({
                    "file": r.get("file"), "status": r.get("status"), "msg": r.get("msg", ""),
                    "confidence": r.get("confidence"), "missing": r.get("missing", []),
                    "warnings": r.get("warnings", []), "review_reason": r.get("review_reason", ""),
                    "output_pdf": r.get("output_pdf", ""),
                    "missing_fields": r.get("missing_fields", []),
                    "low_confidence_fields": r.get("low_confidence_fields", []),
                    "manual_review_required": r.get("manual_review_required", False),
                    "manual_review_reason": r.get("manual_review_reason", ""),
                    "page_range": r.get("page_range"),
                    "extracted_fields": r.get("extracted_fields", {}),
                })
                STATE["done"] = i
        finally:
            STATE["running"] = False
    threading.Thread(target=worker, daemon=True).start()
    return {"started": True, "total": len(files), "language": lang}

@app.post("/api/process_one")
async def process_one_ep(req: Request):
    """Process a single named file (used by Manual Review reprocess)."""
    body = await req.json()
    name = body.get("file", "")
    fp = os.path.join(pc.INBOX, name)
    if not os.path.exists(fp):
        return JSONResponse({"error": "file not found in Inbox"}, status_code=404)
    lang = ocr.LANGUAGE_OPTIONS.get(body.get("language", ""), body.get("language")) \
           or pc.CFG.get("ocr_default_language")
    rng = None
    if body.get("range_start") and body.get("range_end"):
        rng = (int(body["range_start"]), int(body["range_end"]))
    r = pc.process_one(fp, lang=lang,
                       manual_agreement_date=body.get("agreement_date", "") or "",
                       manual_sr_number=body.get("sr_number", "") or "",
                       manual_range=rng)
    return r

@app.get("/api/status")
def status():
    return STATE

# --------------------------------------------------------------------------
# OCR : health check + confidence test  (spec 5 & 6)
# --------------------------------------------------------------------------
@app.get("/api/ocr/health")
def ocr_health():
    return ocr.health_check(pc.CFG)

@app.get("/api/ocr/resolve")
def ocr_resolve():
    return ocr.resolve_ocr(pc.CFG)

@app.post("/api/ocr/confidence")
async def ocr_conf(req: Request):
    try:
        body = await req.json()
    except Exception:
        body = {}
    lang = ocr.LANGUAGE_OPTIONS.get(body.get("language", ""), body.get("language")) \
           or pc.CFG.get("ocr_default_language")
    return ocr.confidence_test(pc.CFG, lang=lang)

# --------------------------------------------------------------------------
# Recipients  (spec 16.3)
# --------------------------------------------------------------------------
@app.get("/api/recipients")
def get_recipients():
    f = os.path.join(pc.HERE, "recipients.json")
    return json.load(open(f, encoding="utf-8")) if os.path.exists(f) else {}

@app.post("/api/recipients")
async def save_recipients(req: Request):
    data = await req.json()
    json.dump(data, open(os.path.join(pc.HERE, "recipients.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=0)
    return {"ok": True}

# --------------------------------------------------------------------------
# Settings  (spec 16.5)
# --------------------------------------------------------------------------
@app.get("/api/settings")
def get_settings():
    cfg = json.load(open(CFG_PATH, encoding="utf-8")) if os.path.exists(CFG_PATH) else {}
    res = ocr.resolve_ocr(cfg)
    cfg["_resolved_tesseract"] = res["exe"]
    cfg["_resolved_tessdata"] = res["tessdata"]
    cfg["_ocr_source"] = res["source"]
    cfg["_languages"] = list(ocr.LANGUAGE_OPTIONS.keys())
    return cfg

@app.post("/api/settings")
async def save_settings(req: Request):
    data = await req.json()
    cur = json.load(open(CFG_PATH, encoding="utf-8")) if os.path.exists(CFG_PATH) else {}
    for k in ("sync_dir", "output_dir", "name_separator", "make_outlook_draft",
              "ocr_default_language", "ocr_dpi", "ocr_mode", "ocr_force_bundled"):
        if k in data:
            cur[k] = data[k]
    json.dump(cur, open(CFG_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return {"ok": True, "note": "Restart the app to apply folder/path changes."}

# --------------------------------------------------------------------------
# Pending / Manual review, History, Audit, file download
# --------------------------------------------------------------------------
@app.get("/api/pending")
def pending():
    """Back-compat: surface the manual-review queue as a simple item list."""
    items = [{"file": r.get("file"), "reason": r.get("reason", "")}
             for r in pc.mrq_list() if r.get("status") != "Corrected"]
    return {"items": items}


# ---- Manual Review (spec ISSUE 1) ----------------------------------------
MR_FIELDS = ["agreement_date", "sr_number", "po_no", "vendor_account",
             "company", "scope", "amount", "project", "range_start", "range_end"]

@app.get("/api/manual_review")
def manual_review():
    """All files needing correction (missing/weak fields, or undetected range)."""
    return {"items": pc.mrq_list(), "fields": MR_FIELDS}

@app.post("/api/manual_review/{file_id}/save")
async def manual_review_save(file_id: str, req: Request):
    body = await req.json()
    rec = pc.mrq_get(file_id) or {"file": file_id}
    rec.setdefault("manual", {})
    for k in MR_FIELDS:
        if k in body and body[k] != "":
            rec["manual"][k] = body[k]
    rec["status"] = "Saved (not yet reprocessed)"
    pc.mrq_upsert(rec)
    return {"ok": True, "record": pc.mrq_get(file_id)}

@app.post("/api/manual_review/{file_id}/reprocess")
async def manual_review_reprocess(file_id: str, req: Request):
    try:
        body = await req.json()
    except Exception:
        body = {}
    fp = os.path.join(pc.INBOX, file_id + ".pdf")
    if not os.path.exists(fp):
        # also accept an exact filename passed in
        cand = os.path.join(pc.INBOX, file_id)
        fp = cand if os.path.exists(cand) else fp
    if not os.path.exists(fp):
        return JSONResponse({"error": "source PDF not found in Inbox: %s.pdf" % file_id}, status_code=404)
    rec = pc.mrq_get(file_id) or {}
    manual = dict(rec.get("manual", {}))
    for k in MR_FIELDS:                      # body overrides saved values
        if k in body and body[k] != "":
            manual[k] = body[k]
    lang = ocr.LANGUAGE_OPTIONS.get(body.get("language", ""), body.get("language")) \
           or pc.CFG.get("ocr_default_language")
    r = pc.process_one(fp, lang=lang, manual=manual)
    return {"ok": True, "result": r, "record": pc.mrq_get(file_id)}

@app.get("/api/history")
def history():
    return {"audit": pc.read_audit(),
            "outputs": [os.path.basename(f) for f in sorted(glob.glob(os.path.join(pc.OUTPUT, "*.pdf")))]}

@app.get("/api/download")
def download(name: str):
    fp = os.path.join(pc.OUTPUT, os.path.basename(name))
    if os.path.exists(fp):
        return FileResponse(fp, filename=os.path.basename(fp))
    return JSONResponse({"error": "not found"}, status_code=404)

# serve the frontend (index.html) at "/"
FRONT = os.path.join(ROOT, "web", "frontend")
if not os.path.isdir(FRONT):
    FRONT = os.path.abspath(os.path.join(HERE, "..", "frontend"))
if os.path.isdir(FRONT):
    app.mount("/", StaticFiles(directory=FRONT, html=True), name="static")
