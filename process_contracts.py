# -*- coding: utf-8 -*-
"""
Kareem T-01 — EDECS Contract Processor  (local Windows app — backend engine)

For every bundle PDF (uploaded in the app or placed in Inbox):
  1. Reads project code / contract no / annex from the file name.
  2. OCRs the pages (bundled Tesseract), detects the Purchase-Order report and the
     contract copy that follows it.  Detection order: visual-repeat -> FRM footer ->
     contract-number token -> manual [s-e] range.
  3. Saves ONE contract copy as a correctly-named PDF in Output.
  4. Builds a fully-mapped 26-column SR Log row and writes it to a SEPARATE file
     ("SR Log - NEW ROWS.xlsx") — the shared master is read read-only for de-dup only.
  5. Builds an editable .eml draft (X-Unsent:1) or a real Outlook draft.

This version adds (per the build spec):
  * Bundled OCR via ocr_engine (system Tesseract is fallback only).
  * Selectable OCR language per run.
  * Per-page OCR confidence + <70% -> Manual Review routing.
  * Optional manual Agreement Date & SR Number (override OCR; never block on missing).
  * Missing-field highlighting flags + SR Log notes.
  * Structured audit/history log (audit_log.jsonl).
The original detection & enrichment logic is preserved.
"""
import os, re, json, glob, getpass, datetime, traceback

import ocr_engine as OCR

HERE = os.environ.get("KT01_ROOT") or os.path.dirname(os.path.abspath(__file__))

def load_config():
    p = os.path.join(HERE, "config.json")
    cfg = json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
    if not cfg.get("sync_dir"):
        for base in [os.environ.get("OneDrive", ""), os.environ.get("OneDriveCommercial", ""),
                     os.path.join(os.path.expanduser("~"), "OneDrive")]:
            cand = os.path.join(base, "Procurement-Sync") if base else ""
            if cand and os.path.isdir(cand):
                cfg["sync_dir"] = cand; break
    cfg.setdefault("sync_dir", HERE)                       # default: app folder (local, offline)
    cfg.setdefault("ocr_mode", "bundled")
    cfg.setdefault("ocr_default_language", "ara+eng")
    cfg.setdefault("ocr_dpi", 300)
    cfg.setdefault("ocr_confidence_test", True)
    cfg.setdefault("ocr_force_bundled", False)             # if True, never fall back to system
    cfg.setdefault("name_separator", "_")
    cfg.setdefault("make_outlook_draft", True)
    cfg.setdefault("output_dir", "")                       # optional explicit output folder
    # robust: if SR Log isn't in sync_dir, look beside the app / its parent
    if not os.path.exists(os.path.join(cfg["sync_dir"], "SR Log 2026.xlsx")):
        for cand in (HERE, os.path.dirname(HERE)):
            if os.path.exists(os.path.join(cand, "SR Log 2026.xlsx")):
                cfg["sync_dir"] = cand; break
    return cfg

CFG    = load_config()
SYNC   = CFG["sync_dir"]
INBOX  = os.path.join(SYNC, "Inbox")
OUTPUT = CFG.get("output_dir") or os.path.join(SYNC, "Output")
SRLOG  = os.path.join(SYNC, "SR Log 2026.xlsx")
AUDIT  = os.path.join(OUTPUT, "audit_log.jsonl")

REFS   = json.load(open(os.path.join(HERE, "refs.json"), encoding="utf-8"))
RECIPS = json.load(open(os.path.join(HERE, "recipients.json"), encoding="utf-8"))
KB, LABELS, AR_ENG = REFS["kb"], REFS["label"], REFS["ar_eng"]

MON_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
CONF_THRESHOLD = 70.0     # below this -> Manual Review (spec section 6)

WORK_MAP = [
    (r"equipment|equ\b", ("Equ. Rent", "Logistic - Equ Rent", "معدات")),
    (r"consult",         ("Consultant", "Consultant agreement", "خدمات استشارية")),
    (r"transport|waste", ("Logistic", "Logistic - Inland transportation", "نقل")),
    (r"survey",          ("Survey", "Subcontractor agreement", "أعمال مساحة")),
    (r"ready ?mix",      ("Ready Mix", "Subcontractor agreement", "خرسانة جاهزة")),
    (r"steel|formwork|metal", ("Steel", "Subcontractor agreement", "أعمال حدادة ونجارة")),
    (r"asphalt|road",    ("Asphalt", "Subcontractor agreement", "أعمال أسفلت وطرق")),
    (r"paint",           ("Painting", "Subcontractor agreement", "أعمال دهانات")),
    (r"plaster|masonry|gypsum", ("Finishing", "Subcontractor agreement", "أعمال تشطيبات")),
    (r"marble|floor|interlock", ("Finishing", "Subcontractor agreement", "أعمال أرضيات")),
    (r"pile|soil|earth|concrete", ("Civil", "Subcontractor agreement", "أعمال خرسانة وأساسات")),
    (r"dolomite|sand|aggregate|rock", ("Materials", "Subcontractor agreement", "توريد مواد")),
]
def map_work(w):
    if not w: return ("", "Subcontractor agreement", "")
    for pat, val in WORK_MAP:
        if re.search(pat, w.lower()): return val
    return (w, "Subcontractor agreement", w)

def ncode(c):
    """Normalise project code: strip leading zeros from base, zero-pad sub to 2, keep letters."""
    c = c.replace("_", "-"); ps = c.split("-")
    b = ps[0].lstrip("0") or "0"
    sub = [p.zfill(2) if p.isdigit() else p for p in ps[1:]]
    return "-".join([b] + sub) if sub else b

# Accepted filename examples shown to the user when parsing fails (no AI; deterministic).
PARSE_EXAMPLES = (
    "107-59659.pdf | 98-60033.pdf | "
    "72-49835 Annex(1).pdf | 78-48033 Annex (2).pdf | "
    "099-02-26700 Annex (1).pdf | CON-01-57697.pdf | "
    "099-02-26700 [3-4].pdf | 099-02-26700 Annex(1) [3-4].pdf"
)

def annex_label(n):
    """Normalise annex number to the canonical display form 'Annex (N)'."""
    return "Annex (%d)" % n if n else ""

def parse_bundle_name(fn):
    """
    Deterministic, AI-free bundle-name parser. Tolerant of spacing/case.

    Layout:  <projectCode>-<contractNo>[ Annex(N) | Annex (N)] [ [s-e] ]
      * projectCode = digits, optional '-NN' sub-project, or letter codes (CON-01, N-Store)
      * contractNo  = the LAST run of 3-6 digits that belongs to the code
      * Annex / addendum number optional, any case, with or without a space before '('
      * manual page range [s-e] optional
    Returns dict(code, no, annex, manual) or None if it cannot be parsed.
    """
    if not fn:
        return None
    name = os.path.splitext(str(fn))[0]          # ignore extension
    name = re.sub(r"\s+", " ", name).strip()      # trim / collapse spaces

    # 1) manual page range [s-e]  (also tolerates [s - e])
    manual = None
    mr = re.search(r"\[\s*(\d+)\s*-\s*(\d+)\s*\]", name)
    if mr:
        manual = (int(mr.group(1)), int(mr.group(2)))
    core = re.sub(r"\[\s*\d+\s*-\s*\d+\s*\]", " ", name)   # strip range token

    # 2) annex / addendum number  (Annex(1), Annex (1), ANNEX1, addendum 2, ...)
    annex = None
    ma = re.search(r"\b(?:annex|addendum)\b\s*\(?\s*(\d+)\s*\)?", core, re.I)
    if ma:
        annex = int(ma.group(1))
    core = re.sub(r"\b(?:annex|addendum)\b\s*\(?\s*\d*\s*\)?", " ", core, flags=re.I)
    core = re.sub(r"\s+", " ", core).strip(" -_")

    # 3) split <code>-<contractNo>: contractNo is the LAST 3-6 digit group.
    #    Everything before the final '-<digits>' is the project code.
    m = re.match(r"^(?P<code>.+?)[-_ ]+(?P<no>\d{3,6})$", core)
    if not m:
        # also accept a bare 'code no' with a space separator
        m = re.match(r"^(?P<code>.+?)\s+(?P<no>\d{3,6})$", core)
    if not m:
        return None
    code_raw = m.group("code").strip(" -_")
    no = m.group("no")
    if not code_raw:
        return None
    return dict(code=ncode(code_raw), no=no, annex=annex, manual=manual)

_fitz = None
def get_fitz():
    global _fitz
    if _fitz is None:
        import fitz; _fitz = fitz
    return _fitz

def _prep_ocr():
    OCR.configure_pytesseract(CFG)

def ocr_pages(pdf_path, lang=None, want_conf=False):
    """OCR the first <=8 pages (PO report is always near the front). Optionally collect conf."""
    import pytesseract
    from PIL import Image
    _prep_ocr()
    lang = lang or CFG["ocr_default_language"]
    cfgstr = OCR.tess_config(CFG, extra="--psm 6")
    doc = get_fitz().open(pdf_path)
    zoom = CFG["ocr_dpi"] / 72.0
    mat = get_fitz().Matrix(zoom, zoom)
    out = []; confs = []
    for i in range(min(len(doc), 8)):
        pix = doc[i].get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        # PO report parsing relies on English layout; keep eng for those tokens but honor lang otherwise
        out.append((i + 1, pytesseract.image_to_string(img, lang="eng", config=cfgstr)))
        if want_conf:
            confs.append(OCR.page_confidence(img, CFG, lang))
    doc.close()
    if want_conf:
        avg = round(sum(confs) / len(confs), 1) if confs else 0.0
        return out, avg
    return out

def parse_po(text):
    d = {}
    m = re.search(r"ED[EC]{1,2}G?-?0*([0-9]{3,})", text); d["po_no"] = m.group(1) if m else None
    m = re.search(r"(\d{4}-\d{2}-\d{2})", text); d["po_date"] = m.group(1) if m else None
    m = re.search(r"\b(SUB|SER|CU)-?\s?0*([0-9]{2,})", text)
    d["vendor_account"] = f"{m.group(1)}-{m.group(2).zfill(5)}" if m else None
    amts = [(float(a.replace(",", "")), s) for s, a in re.findall(r"(-?)\s*([0-9][0-9,]*\.\d{2})\s*EGP", text)]
    pos = [a for a, s in amts if s == "" and a > 0]; negs = [a for a, s in amts if s == "-"]
    if pos:
        gross = max(pos); net = gross - sum(abs(x) for x in negs)
        exp = [a for a in pos if abs(a - net) < 1 and a != gross]
        d["net"] = round(exp[0] if exp else net, 2)
    else:
        d["net"] = None
    m = re.search(r"Report generated:\s*(\d{2}/\d{2}/\d{4})", text); d["report_date"] = m.group(1) if m else None
    m = re.search(r"APPROVED BY[^\n]*\n\s*((?:Dr\.?\s*|Eng\.?\s*)?[A-Za-z]+\s+[A-Za-z]+)", text)
    d["approved_by"] = m.group(1).strip() if m else None
    m = re.search(r"ISSUED BY[^\n]*\n(.+)", text)
    d["issued_raw"] = m.group(1) if m else ""
    d["issued_by"] = " ".join(m.group(1).split()[-2:]) if m else None
    return d

# ---- SR Number extraction (spec section 8) -------------------------------
def extract_sr_number(pages):
    """Try to find an SR number token in the OCR'd pages, e.g. SR-2026-00125 / SR 125."""
    text = "\n".join(t for _, t in pages)
    for pat in (r"SR[-\s]?No\.?\s*[:.]?\s*(SR[-\s]?\d{2,4}[-\s]?\d{2,6})",
                r"\b(SR[-\s]?\d{4}[-\s]?\d{2,6})\b",
                r"SR[-\s]?No\.?\s*[:.]?\s*([0-9]{2,8})"):
        m = re.search(pat, text, re.I)
        if m:
            return re.sub(r"\s+", "-", m.group(1).strip())
    return ""

def page_vectors(pdf_path):
    import numpy as np
    from PIL import Image
    doc = get_fitz().open(pdf_path)
    mat = get_fitz().Matrix(70 / 72.0, 70 / 72.0)
    V = []
    for i in range(len(doc)):
        pix = doc[i].get_pixmap(matrix=mat)
        im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("L").resize((36, 50))
        a = np.asarray(im, dtype=float); a = (a - a.mean()) / (a.std() + 1e-6)
        V.append(a.flatten())
    doc.close()
    return V

def best_repeat(V, thr=0.6):
    """First contract copy = a contiguous page block that repeats later (the 2nd copy)."""
    import numpy as np
    n = len(V)
    if n < 2: return None
    Lv = len(V[0]); best = None
    for s in range(n):
        for L in range(1, (n - s) // 2 + 1):
            avg = sum(float(np.dot(V[s + k], V[s + L + k]) / Lv) for k in range(L)) / L
            if avg > thr:
                score = L * avg
                if best is None or score > best[0]:
                    best = (score, s, L, avg)
    if not best: return None
    _, s, L, avg = best
    copies = 2; k = 2
    while s + (k + 1) * L <= n and \
          sum(float(np.dot(V[s + m], V[s + k * L + m]) / Lv) for m in range(L)) / L > thr:
        copies += 1; k += 1
    return dict(start=s, length=L, copies=copies, sim=avg)

def frm_flags(pdf_path):
    """True per page if the contract form footer (FRM-...) is on its bottom ~14% strip."""
    import pytesseract
    from PIL import Image
    _prep_ocr()
    cfgstr = OCR.tess_config(CFG, extra="--psm 6")
    doc = get_fitz().open(pdf_path)
    mat = get_fitz().Matrix(150 / 72.0, 150 / 72.0)
    flags = []
    for i in range(len(doc)):
        pix = doc[i].get_pixmap(matrix=mat)
        im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        w, h = im.size
        strip = im.crop((0, int(h * 0.86), w, h))
        t = pytesseract.image_to_string(strip, lang="eng", config=cfgstr)
        flags.append(bool(re.search(r"FRM", t)))
    doc.close()
    return flags

def longest_frm_run(flags):
    n = len(flags); runs = []; i = 0
    while i < n:
        if flags[i]:
            j = i
            while j < n and (flags[j] or (j + 1 < n and flags[j + 1])):
                j += 1
            runs.append((i, j - 1)); i = j
        else:
            i += 1
    if not runs: return None
    return max(runs, key=lambda r: r[1] - r[0])

def token_run(pages, code, no):
    """Short annexes: the contract page shows 'no/code' (e.g. 26700/099-02)."""
    base = code.split("-")[0]
    tok = re.compile(rf"{no}\s*/\s*0*{base}|0*{base}\s*/\s*{no}")
    hits = [i for i, (_, t) in enumerate(pages) if tok.search(t)]
    if not hits:
        return None
    s = e = hits[0]
    for i in hits[1:]:
        if i == e + 1: e = i
        else: break
    return (s, e)

def contract_date(pdf_path, page_idx, lang=None):
    """Read the contract date from the contract's first page (Arabic+English OCR)."""
    import pytesseract
    from PIL import Image
    _prep_ocr()
    lang = lang or CFG["ocr_default_language"] or "ara+eng"
    cfgstr = OCR.tess_config(CFG, extra="--psm 6")
    try:
        doc = get_fitz().open(pdf_path)
        pix = doc[page_idx].get_pixmap(matrix=get_fitz().Matrix(150 / 72.0, 150 / 72.0))
        im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        t = pytesseract.image_to_string(im, lang=lang, config=cfgstr)
    except Exception:
        return ("", "")
    m = re.search(r"(\d{4})\s*[/\-]\s*(\d{1,2})\s*[/\-]\s*(\d{1,2})", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return (f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}", f"{MON_ABBR[int(m.group(2))]} {m.group(1)}")
    m = re.search(r"(\d{1,2})\s*[/\-]\s*(\d{1,2})\s*[/\-]\s*(\d{4})", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return (f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}", f"{MON_ABBR[int(m.group(2))]} {m.group(3)}")
    return ("", "")

def analyse(pdf_path, pages, code, no):
    po_text = "\n".join(t for _, t in pages if re.search(r"purch.{0,4}order", t, re.I))
    rep = best_repeat(page_vectors(pdf_path), thr=0.92)  # old bundles: 2 identical copies
    if rep:
        s, L = rep["start"], rep["length"]
        return dict(copy_pages=list(range(s, s + L)), copies=rep["copies"], po_text=po_text, confident=True)
    blk = longest_frm_run(frm_flags(pdf_path))           # modern multi-page: one copy, FRM footer
    if blk and blk[1] > blk[0]:
        return dict(copy_pages=list(range(blk[0], blk[1] + 1)), copies=1, po_text=po_text, confident=True)
    tr = token_run(pages, code, no)                      # short annex: contract no on the page
    if tr:
        return dict(copy_pages=list(range(tr[0], tr[1] + 1)), copies=1, po_text=po_text, confident=True)
    return None

_SR = None
def sr_refs():
    global _SR
    if _SR is not None: return _SR
    import openpyxl
    try:
        wb = openpyxl.load_workbook(SRLOG, data_only=True)
    except Exception:
        _SR = dict(vend={}, scope={}, div={}, cat={}, names=set()); return _SR
    vend = {}
    if "Vendors" in wb.sheetnames:
        vs = wb["Vendors"]
        for r in range(2, vs.max_row + 1):
            a = vs.cell(r, 1).value
            if a: vend[str(a).strip().replace("_", "-")] = vs.cell(r, 2).value
    scope = {}; div = {}; cat = {}
    ws = wb["SR Log"]
    for r in range(2, ws.max_row + 1):
        a = str(ws.cell(r, 11).value or "").replace("_", "-")
        if not a: continue
        if a not in scope and ws.cell(r, 9).value: scope[a] = ws.cell(r, 9).value
        if a not in div and ws.cell(r, 13).value: div[a] = ws.cell(r, 13).value
        if a not in cat and ws.cell(r, 10).value: cat[a] = ws.cell(r, 10).value
    names = set()
    if "ERP User" in wb.sheetnames:
        eu = wb["ERP User"]
        for r in range(2, eu.max_row + 1):
            v = eu.cell(r, 2).value
            if v and isinstance(v, str) and len(v.split()) >= 2: names.add(v.strip())
    for d in RECIPS.values():
        for fld in ("to", "cc"):
            for nm in re.findall(r"([A-Za-z][A-Za-z]+(?:\s+[A-Za-z]+){1,3})\s*<", d.get(fld, "")):
                names.add(nm.strip())
    wb.close()
    _SR = dict(vend=vend, scope=scope, div=div, cat=cat, names=names)
    return _SR

def match_person(text, names):
    t = (text or "").lower(); best = ""; key = (-1, -1)
    for nm in names:
        i = t.rfind(nm.lower())
        if i >= 0 and (i, len(nm)) > key: key = (i, len(nm)); best = nm
    return best

def norm_ar(s):
    s = re.sub(r"[ًٌٍَُِّْـ‏‎]", "", s or "")
    for a, b in [("أ","ا"),("إ","ا"),("آ","ا"),("ة","ه"),("ى","ي")]:
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s).strip()

def ar_to_eng(arabic):
    if not arabic: return ""
    na = norm_ar(arabic)
    for k, v in AR_ENG.items():
        if norm_ar(k) == na: return v
    a2 = " ".join(na.split()[:2])
    for k, v in AR_ENG.items():
        if " ".join(norm_ar(k).split()[:2]) == a2: return v
    return ""

def enrich(meta, po):
    code, no = meta["code"], meta["no"]
    ref = KB.get(f"{code}|{no}", {}); sr = sr_refs(); acct = po.get("vendor_account") or ""
    label = ref.get("label") or LABELS.get(code) or code
    arabic_excel = sr["vend"].get(acct) or ref.get("arabic") or ""
    arabic_file = ref.get("arabic") or arabic_excel
    scope = (ref.get("work") or sr["scope"].get(acct) or "").strip()
    veng = ref.get("veng") or ar_to_eng(arabic_file) or arabic_file
    cat = sr["cat"].get(acct) or map_work(scope)[1]
    div = sr["div"].get(acct) or map_work(scope)[2]
    person = match_person(po.get("issued_raw", ""), sr["names"]) or (po.get("issued_by") or "")
    return dict(code=code, no=no, annex=meta["annex"], label=label, veng=veng, work=scope,
                arabic_excel=arabic_excel, arabic_file=arabic_file, scope=scope, cat=cat, div=div, person=person)

def build_name(meta, po, en, my):
    head = "Agreement" + (f" Annex ({en['annex']})" if en["annex"] else "")
    mid = f"PJ{en['code']}_{en['no']} {en['label']}-{en['veng']}-{en['work']}-{my}".rstrip("-")
    sep = CFG["name_separator"]
    return (f"{head}-{mid}{sep}{en['arabic_file']}").rstrip(sep).strip()

def _safe(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()

def write_contract_pdf(src, copy_pages, out_name):
    doc = get_fitz().open(src); new = get_fitz().open()
    for p in copy_pages:
        new.insert_pdf(doc, from_page=p, to_page=p)
    dest = os.path.join(OUTPUT, _safe(out_name) + ".pdf")
    new.save(dest); new.close(); doc.close()
    return dest

SR_COLS = 26
def build_row(meta, po, copies, en, date_iso, sr_number, note_extra=""):
    proj = LABELS.get(en["code"], en["label"])
    typ = f"Addendum {en['annex']}" if en["annex"] else "Contract"
    rd = po.get("report_date") or ""
    note = (po.get("approved_by") or "")
    if note_extra:
        note = (note + " | " + note_extra).strip(" |")
    row = [""] * SR_COLS
    row[0]=en["code"]; row[1]=proj; row[3]=sr_number or ""; row[4]=date_iso; row[7]=en["person"]
    row[8]=en["scope"]; row[9]=en["cat"]; row[10]=po.get("vendor_account") or ""; row[11]=en["arabic_excel"]
    row[12]=en["div"]; row[13]=en["no"]; row[14]=typ; row[15]=date_iso
    row[16]=po.get("net") if po.get("net") is not None else ""
    row[19]=rd; row[20]=rd; row[21]="Open"; row[22]=note; row[23]=copies
    return row

NEW_ROWS_NAME = "SR Log - NEW ROWS.xlsx"
def _master_keys():
    """Read the (possibly server-locked) master read-only for de-dup; safe if open elsewhere."""
    import openpyxl
    seen = set()
    try:
        m = openpyxl.load_workbook(SRLOG, read_only=True); ws = m["SR Log"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0]: seen.add((str(r[0] or ""), str(r[13] or ""), str(r[14] or "")))
        m.close()
    except Exception:
        pass
    return seen

def append_sr_row(row, highlight_cols=None):
    """Write new rows to a SEPARATE file in Output (never locks/touches the shared master).
    highlight_cols: list of 0-based column indexes to shade yellow (missing fields)."""
    import openpyxl
    from openpyxl.styles import PatternFill
    out = os.path.join(OUTPUT, NEW_ROWS_NAME)
    seen = _master_keys()
    if os.path.exists(out):
        nb = openpyxl.load_workbook(out); ns = nb["New Rows"]
        for r in ns.iter_rows(min_row=2, values_only=True):
            if r and r[0]: seen.add((str(r[0] or ""), str(r[13] or ""), str(r[14] or "")))
    else:
        nb = openpyxl.Workbook(); ns = nb.active; ns.title = "New Rows"
        hdr = None
        try:
            m = openpyxl.load_workbook(SRLOG, read_only=True); ws = m["SR Log"]
            hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]; m.close()
        except Exception:
            pass
        ns.append(hdr or DEFAULT_HEADERS)
    if (str(row[0]), str(row[13]), str(row[14])) in seen:
        nb.close(); return False
    ns.append(row)
    if highlight_cols:
        fill = PatternFill(start_color="FFF3B0", end_color="FFF3B0", fill_type="solid")
        r = ns.max_row
        for c in highlight_cols:
            ns.cell(row=r, column=c + 1).fill = fill
    nb.save(out); nb.close(); return True

DEFAULT_HEADERS = ["Project Code","Project Name","Request Type","No.","Create Date",
    "Cost Control Sign Date","Procurement Sign Date","Person","Scope","Category",
    "Vendor account","Vendor","Division","PO / Contract No.","Type","Date","Amount",
    "Procurement Sign Date2","Cost Control Sign Date2","Chairman Sign Date","Distribution Date",
    "Status","Note","Archive Copies No.","Archive Vendor","Archive Date"]

def recipients_for(code, work):
    r = RECIPS.get(code) or RECIPS.get(code.split("-")[0]) or {"to": "", "cc": ""}
    cc = [p.strip() for p in re.split(r";\s*", r.get("cc", "")) if p.strip()]
    if "equipment" not in (work or "").lower():
        cc = [p for p in cc if "logistics" not in p.lower()]
    return r.get("to", ""), "; ".join(cc)

def email_body(proj, code, veng, work, copies):
    cp = {1:"One",2:"Two",3:"Three",4:"Four"}.get(copies, str(copies))
    return ("Dear All,\r\n\r\nKindly find attached Subcontractor's Contract detailed as follows:-\r\n\r\n"
            f"Project          : {proj} (PJ {code}).\r\n"
            f"Contractor       : {veng}.\r\n"
            f"Contract Subject : {work}.\r\n"
            f"Copies           : {cp} Original(s).\r\n\r\n"
            "This is for your information and Record.\r\n\r\nBest Regards,\r\n")

def make_draft(name, to, cc, subject, body, pdf_path):
    name = _safe(name)
    if CFG["make_outlook_draft"]:
        try:
            import win32com.client
            ol = win32com.client.Dispatch("Outlook.Application")
            mail = ol.CreateItem(0)
            mail.To = to; mail.CC = cc; mail.Subject = subject; mail.Body = body
            mail.Attachments.Add(os.path.abspath(pdf_path)); mail.Save()
            return "Outlook draft", ""
        except Exception:
            pass
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["X-Unsent"] = "1"; msg["To"] = to; msg["Cc"] = cc; msg["Subject"] = subject
    msg.set_content(body)
    with open(pdf_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=name + ".pdf")
    eml = os.path.join(OUTPUT, name + ".eml")
    open(eml, "wb").write(bytes(msg))
    return ".eml file", eml

# ---- Audit log (spec section 17) -----------------------------------------
def audit(record):
    os.makedirs(OUTPUT, exist_ok=True)
    record = dict(record)
    record.setdefault("timestamp", datetime.datetime.now().isoformat(timespec="seconds"))
    record.setdefault("user", getpass.getuser())
    with open(AUDIT, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")

def read_audit():
    if not os.path.exists(AUDIT): return []
    out = []
    for line in open(AUDIT, encoding="utf-8"):
        line = line.strip()
        if line:
            try: out.append(json.loads(line))
            except Exception: pass
    return out

# ---------------------------------------------------------------------------
# MAIN: process a single bundle. Returns a rich dict for the API/UI.
#   status: "done" | "warning" | "manual" | "failed"
# ---------------------------------------------------------------------------
def process_one(pdf_path, lang=None, manual_agreement_date="", manual_sr_number="",
                manual_range=None):
    fn = os.path.splitext(os.path.basename(pdf_path))[0]
    lang = lang or CFG["ocr_default_language"]
    result = {"file": fn, "status": "failed", "msg": "", "confidence": None,
              "missing": [], "warnings": [], "output_pdf": "", "eml": "",
              "sr_added": False, "review_reason": "",
              # filename pre-processing preview (deterministic, no AI):
              "parsed": {"original": os.path.basename(pdf_path), "code": "", "no": "",
                         "annex": None, "type": "", "status": "Needs Manual Review"}}

    meta = parse_bundle_name(fn)
    if not meta:
        result["parsed"]["status"] = "Needs Manual Review"
        result.update(status="manual",
                      review_reason=("Filename not understood. Accepted examples: " + PARSE_EXAMPLES),
                      msg="اسم الملف غير مفهوم. أمثلة مقبولة: " + PARSE_EXAMPLES)
        audit({"file": fn, "status": "manual", "manual_review_reason": result["review_reason"]})
        return result

    # populate the parse preview now that the filename is understood
    result["parsed"].update(
        code=meta["code"], no=meta["no"], annex=meta["annex"],
        type=("Addendum %d" % meta["annex"] if meta["annex"] else "Contract"),
        status="Parsed")

    # OCR + confidence
    try:
        pages, conf = ocr_pages(pdf_path, lang=lang, want_conf=True)
    except Exception as e:
        result.update(status="failed", msg="OCR error: %s" % e)
        audit({"file": fn, "status": "failed", "msg": result["msg"], "ocr_language": lang})
        return result
    result["confidence"] = conf

    # <70% confidence -> Manual Review (do not silently finalise)
    if conf and conf < CONF_THRESHOLD:
        result.update(status="manual",
                      review_reason="OCR confidence %.1f%% < %d%%." % (conf, CONF_THRESHOLD),
                      msg="ثقة OCR منخفضة (%.1f%%) — مراجعة يدوية" % conf)
        audit({"file": fn, "status": "manual", "ocr_confidence": conf,
               "ocr_language": lang, "manual_review_reason": result["review_reason"]})
        return result

    # contract detection
    rng = manual_range or meta.get("manual")
    if rng:
        s, e = rng
        po_text = "\n".join(t for _, t in pages if re.search(r"purch.{0,4}order", t, re.I))
        info = {"copy_pages": list(range(s - 1, e)), "copies": 1, "po_text": po_text, "confident": True}
    else:
        info = analyse(pdf_path, pages, meta["code"], meta["no"])
    if not info or not info["copy_pages"] or not info["confident"]:
        result.update(status="manual",
                      review_reason="Contract page range not detected — add [s-e] to filename or set range.",
                      msg="تعذّر تحديد صفحات العقد — أضف [s-e] لاسم الملف")
        audit({"file": fn, "status": "manual", "ocr_confidence": conf,
               "ocr_language": lang, "manual_review_reason": result["review_reason"]})
        return result

    po = parse_po(info["po_text"]) if info["po_text"] else {}
    en = enrich(meta, po)

    # ---- Agreement Date (manual overrides OCR; never blocks) -------------
    ci, cmy = contract_date(pdf_path, info["copy_pages"][0], lang=lang)
    po_my = ""
    if po.get("po_date"):
        y, m, _ = po["po_date"].split("-"); po_my = f"{MON_ABBR[int(m)]} {y}"
    detected_date = ci
    agreement_date = (manual_agreement_date or "").strip() or detected_date or po.get("po_date") \
                     or KB.get(f"{meta['code']}|{meta['no']}", {}).get("my", "")
    date_iso = agreement_date or ""
    # month-year used in the file name / subject
    my = ""
    if manual_agreement_date:
        mm = re.search(r"(\d{1,2})\D+(\d{4})", manual_agreement_date) or \
             re.search(r"(\d{4})\D+(\d{1,2})", manual_agreement_date)
        if mm:
            g = mm.groups()
            month = int(g[0]) if len(g[0]) <= 2 else int(g[1])
            year = g[1] if len(g[1]) == 4 else g[0]
            if 1 <= month <= 12: my = f"{MON_ABBR[month]} {year}"
    my = my or cmy or po_my or KB.get(f"{meta['code']}|{meta['no']}", {}).get("my", "")

    # ---- SR Number (manual overrides OCR; never blocks) ------------------
    detected_sr = extract_sr_number(pages)
    sr_number = (manual_sr_number or "").strip() or detected_sr

    # ---- missing-field flags & notes -------------------------------------
    notes = []
    highlight = []          # 0-based SR Log columns to shade
    if not date_iso:
        result["missing"].append("agreement_date")
        result["warnings"].append("Agreement Date missing / not detected")
        notes.append("Agreement Date missing / not detected")
        highlight += [4, 15]     # Create Date, Date
    if not sr_number:
        result["missing"].append("sr_number")
        result["warnings"].append("SR Number missing / not detected")
        notes.append("SR Number missing / not detected")
        highlight += [3]         # No.
    if not po.get("vendor_account"):
        result["warnings"].append("Vendor account missing")
        highlight += [10]
    if po.get("net") is None:
        result["warnings"].append("Amount not detected")
        highlight += [16]

    # ---- outputs ----------------------------------------------------------
    name = build_name(meta, po, en, my)
    pdf_out = write_contract_pdf(pdf_path, info["copy_pages"], name)
    added = append_sr_row(build_row(meta, po, info["copies"], en, date_iso, sr_number,
                                    note_extra=" | ".join(notes)),
                          highlight_cols=sorted(set(highlight)))
    proj = LABELS.get(meta["code"], en["label"])
    to, cc = recipients_for(meta["code"], en["scope"])
    body = email_body(proj, meta["code"], en["veng"], en["scope"], info["copies"])
    kind, eml = make_draft(name, to, cc, name, body, pdf_out)

    status = "warning" if result["warnings"] else "done"
    result.update(status=status, output_pdf=os.path.basename(pdf_out), eml=eml,
                  sr_added=added, copies=info["copies"],
                  msg="عقد: %s | نسخ: %d | SR Log: %s | إيميل: %s%s" % (
                      os.path.basename(pdf_out), info["copies"],
                      "أُضيف" if added else "موجود", kind,
                      (" | تنبيهات: " + "؛ ".join(result["warnings"])) if result["warnings"] else ""))

    audit({"file": fn, "output_pdf": result["output_pdf"], "project_code": meta["code"],
           "contract_no": meta["no"], "ocr_language": lang, "ocr_confidence": conf,
           "detected_agreement_date": detected_date, "manual_agreement_date": manual_agreement_date,
           "detected_sr_number": detected_sr, "manual_sr_number": manual_sr_number,
           "missing_fields": result["missing"], "status": status,
           "email_draft": eml or kind, "sr_log_output": os.path.join(OUTPUT, NEW_ROWS_NAME)})

    # record manual overrides explicitly
    if manual_agreement_date and manual_agreement_date != detected_date:
        audit({"field_changed": "agreement_date", "ocr_value": detected_date,
               "manual_value": manual_agreement_date, "file": fn})
    if manual_sr_number and manual_sr_number != detected_sr:
        audit({"field_changed": "sr_number", "ocr_value": detected_sr,
               "manual_value": manual_sr_number, "file": fn})
    return result

def run(log=print, progress=None, on_result=None, lang=None):
    """Batch-process every PDF in Inbox (advanced / legacy folder mode)."""
    os.makedirs(OUTPUT, exist_ok=True)
    log(f"المجلد: {SYNC}")
    if not os.path.exists(SRLOG):
        log(f"تنبيه: لم يتم العثور على {SRLOG} — سأكمل وأكتب الصفوف في Output بدون إثراء.")
    files = sorted(glob.glob(os.path.join(INBOX, "*.pdf")))
    total = len(files)
    if not files:
        log("لا توجد ملفات في Inbox.")
        if progress: progress(0, 0)
        return
    done = warn = manual = failed = 0
    for i, f in enumerate(files, 1):
        try:
            r = process_one(f, lang=lang)
        except Exception:
            r = {"status": "failed", "file": os.path.basename(f),
                 "msg": "خطأ: " + traceback.format_exc().splitlines()[-1]}
        st = r["status"]
        if st == "done": done += 1; log("[OK] " + r["file"] + " — " + r["msg"])
        elif st == "warning": warn += 1; log("[!] " + r["file"] + " — " + r["msg"])
        elif st == "manual":
            manual += 1; log("[?] " + r["file"] + " — " + r.get("review_reason", r["msg"]))
            with open(os.path.join(OUTPUT, "_pending.txt"), "a", encoding="utf-8") as p:
                p.write(f"{r['file']}\t{r.get('review_reason', r['msg'])}\n")
        else:
            failed += 1; log("[X] " + r["file"] + " — " + r["msg"])
        if on_result: on_result(r["file"], st, r.get("msg", ""))
        if progress: progress(i, total)
    log(f"\nانتهى. تمت {done} | تنبيهات {warn} | مراجعة {manual} | فشل {failed}.")

if __name__ == "__main__":
    run()
