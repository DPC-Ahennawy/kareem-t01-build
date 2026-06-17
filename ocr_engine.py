# -*- coding: utf-8 -*-
"""
Kareem T-01 — EDECS Contract Processor
ocr_engine.py  —  Bundled-OCR resolution, health check, confidence test, language map.

Design goals (from build spec sections 3-6):
  * The app must NOT depend on a manual system install of Tesseract.
  * Primary target is the BUNDLED copy under  app_root/ocr/tesseract/ .
  * A system-installed Tesseract may be used only as a FALLBACK.
  * Provide an OCR Health Check and an OCR Confidence Test that the UI can call.

All processing stays local/offline. No network calls here.
"""
import os, re, sys, shutil, subprocess, tempfile

HERE = os.environ.get("KT01_ROOT") or os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------------------
# Bundled OCR folder layout (spec section 3):
#   app_root/ocr/tesseract/tesseract.exe
#   app_root/ocr/tesseract/tessdata/{eng,ara,fra}.traineddata
# ----------------------------------------------------------------------------
OCR_ROOT      = os.path.join(HERE, "ocr", "tesseract")
BUNDLED_EXE   = os.path.join(OCR_ROOT, "tesseract.exe" if os.name == "nt" else "tesseract")
BUNDLED_TDATA = os.path.join(OCR_ROOT, "tessdata")

LANG_FILES = {"eng": "eng.traineddata", "ara": "ara.traineddata", "fra": "fra.traineddata"}

# UI option label -> tesseract -l parameter  (spec section 4)
LANGUAGE_OPTIONS = {
    "Arabic":                    "ara",
    "English":                   "eng",
    "French":                    "fra",
    "Arabic + English":          "ara+eng",
    "Arabic + English + French": "ara+eng+fra",
}
DEFAULT_LANGUAGE = "ara+eng"


def _system_tesseract():
    """Locate a system-installed tesseract as a fallback (not the primary target)."""
    # 1) explicit common Windows path
    win = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(win):
        return win
    # 2) anything on PATH
    found = shutil.which("tesseract") or shutil.which("tesseract.exe")
    return found


def resolve_ocr(cfg=None):
    """
    Resolve which Tesseract executable + tessdata folder to use.

    Returns dict:
      exe        : path to tesseract executable (or None if nothing usable)
      tessdata   : path to the tessdata folder that will be used
      source     : 'bundled' | 'system' | 'none'
      allow_fallback : whether system fallback is permitted by config
    Honors config.json:
      ocr_mode = 'bundled'  -> prefer bundled, fall back to system unless ocr_force_bundled
    """
    cfg = cfg or {}
    allow_fallback = not bool(cfg.get("ocr_force_bundled", False))

    # Primary: bundled copy
    if os.path.exists(BUNDLED_EXE):
        return {"exe": BUNDLED_EXE, "tessdata": BUNDLED_TDATA,
                "source": "bundled", "allow_fallback": allow_fallback}

    # Optional explicit override in config (still local)
    cfg_exe = cfg.get("tesseract_exe") or ""
    if cfg_exe and os.path.exists(cfg_exe) and allow_fallback:
        td = os.path.join(os.path.dirname(cfg_exe), "tessdata")
        return {"exe": cfg_exe, "tessdata": td if os.path.isdir(td) else BUNDLED_TDATA,
                "source": "system", "allow_fallback": allow_fallback}

    # Fallback: system install
    if allow_fallback:
        sysexe = _system_tesseract()
        if sysexe:
            td = os.path.join(os.path.dirname(sysexe), "tessdata")
            return {"exe": sysexe, "tessdata": td if os.path.isdir(td) else BUNDLED_TDATA,
                    "source": "system", "allow_fallback": allow_fallback}

    return {"exe": None, "tessdata": BUNDLED_TDATA, "source": "none",
            "allow_fallback": allow_fallback}


def configure_pytesseract(cfg=None):
    """
    Point pytesseract at the resolved executable and tessdata folder.

    IMPORTANT: pytesseract builds the command line by splitting the `config` string
    on whitespace, so passing  --tessdata-dir "C:\\path with spaces"  there breaks:
    the literal quote characters get glued onto the path. The robust fix is to set
    TESSDATA_PREFIX in the process environment (no quoting issues, spaces handled by
    the OS) and NOT pass a quoted --tessdata-dir in the config string.

    Tesseract 4/5 expect TESSDATA_PREFIX to be the tessdata directory ITSELF.
    """
    res = resolve_ocr(cfg)
    if res["exe"]:
        try:
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = res["exe"]
        except Exception:
            pass
    td = res["tessdata"]
    if td and os.path.isdir(td):
        # Point at the tessdata folder itself (correct for Tesseract 4 & 5).
        os.environ["TESSDATA_PREFIX"] = td
    return res


def tess_config(cfg=None, extra=""):
    """
    Build the pytesseract `config` string. We deliberately DO NOT put --tessdata-dir
    here (it would be whitespace-split and the quotes mangled). The tessdata folder is
    provided via the TESSDATA_PREFIX environment variable set in configure_pytesseract().
    Only safe, space-free flags (e.g. --psm 6) go in the config string.
    """
    return extra or ""


def _td_dir(cfg=None):
    """Return the resolved tessdata directory (the folder itself)."""
    return resolve_ocr(cfg).get("tessdata")


# ----------------------------------------------------------------------------
# OCR HEALTH CHECK  (spec section 5)
# ----------------------------------------------------------------------------
def health_check(cfg=None):
    """
    Verify each bundled OCR component. Returns a list of {check, status, detail}
    plus an overall 'ok' flag. Statuses use Ready/Not Ready and Passed/Failed.
    """
    cfg = cfg or {}
    res = resolve_ocr(cfg)
    items = []

    exe_ok = bool(res["exe"]) and os.path.exists(res["exe"])
    items.append({"check": "Bundled Tesseract executable found",
                  "status": "Ready" if (res["source"] == "bundled") else
                            ("Ready (system fallback)" if exe_ok else "Not Ready"),
                  "detail": res["exe"] or "not found"})

    td = res["tessdata"]
    td_ok = bool(td) and os.path.isdir(td)
    for lang, fname in LANG_FILES.items():
        fp = os.path.join(td, fname) if td else ""
        ok = bool(fp) and os.path.exists(fp)
        items.append({"check": "%s found" % fname,
                      "status": "Ready" if ok else "Not Ready",
                      "detail": fp or "tessdata folder unknown"})

    readable = td_ok and os.access(td, os.R_OK)
    items.append({"check": "Tessdata folder readable",
                  "status": "Ready" if readable else "Not Ready",
                  "detail": td or "not found"})

    # OCR command executable -> Passed/Failed
    cmd_ok = False; cmd_detail = "executable missing"
    if exe_ok:
        try:
            out = subprocess.run([res["exe"], "--version"], capture_output=True,
                                 text=True, timeout=20)
            cmd_ok = (out.returncode == 0) or bool(out.stdout) or bool(out.stderr)
            cmd_detail = (out.stdout or out.stderr or "").splitlines()[0] if (out.stdout or out.stderr) else "ran"
        except Exception as e:
            cmd_detail = "error: %s" % e
    items.append({"check": "OCR command executable",
                  "status": "Passed" if cmd_ok else "Failed",
                  "detail": cmd_detail})

    # Languages actually LOADABLE: run `tesseract --list-langs` with TESSDATA_PREFIX
    # set to the bundled tessdata folder. This catches path/quoting problems that a
    # plain file-existence check would miss.
    langs_ok = False; langs_detail = "not run"
    listed = set()
    if exe_ok and td_ok:
        try:
            env = dict(os.environ); env["TESSDATA_PREFIX"] = td
            out = subprocess.run([res["exe"], "--list-langs"],
                                 capture_output=True, text=True, timeout=30, env=env)
            blob = (out.stdout or "") + "\n" + (out.stderr or "")
            listed = {ln.strip() for ln in blob.splitlines() if ln.strip() and " " not in ln.strip()}
            need = {"eng", "ara", "fra"}
            langs_ok = need.issubset(listed)
            langs_detail = "loadable: " + ", ".join(sorted(listed & need)) + \
                           ("" if langs_ok else " (missing: %s)" % ", ".join(sorted(need - listed)))
        except Exception as e:
            langs_detail = "error: %s" % e
    items.append({"check": "Languages loadable (eng, ara, fra)",
                  "status": "Passed" if langs_ok else "Failed",
                  "detail": langs_detail})

    critical_ok = exe_ok and td_ok and cmd_ok and langs_ok and all(
        os.path.exists(os.path.join(td, f)) for f in LANG_FILES.values()) if td_ok else False
    return {"ok": bool(critical_ok), "source": res["source"], "exe": res["exe"],
            "tessdata": td, "items": items}


# ----------------------------------------------------------------------------
# OCR CONFIDENCE TEST  (spec section 6)
# ----------------------------------------------------------------------------
SAMPLE_TEXT_LINES = [
    "Contract No.: CON-01-57697",
    "SR No.: SR-2026-00125",
    "Agreement Date: 15/06/2026",
    "Vendor: ABC Construction",
    "Date de l'accord: 15/06/2026",
]
SAMPLE_TEXT_AR = "العقد رقم: 57697"


def _make_sample_image(path):
    """Render an internal multilingual test image (no external file needed)."""
    from PIL import Image, ImageDraw, ImageFont
    W, H = 1000, 360
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    font = None
    for cand in [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\tahoma.ttf",
    ]:
        if os.path.exists(cand):
            try:
                font = ImageFont.truetype(cand, 30); break
            except Exception:
                pass
    if font is None:
        font = ImageFont.load_default()
    y = 20
    for line in SAMPLE_TEXT_LINES:
        d.text((30, y), line, fill="black", font=font); y += 50
    # Arabic line (may render unshaped; still exercises ara model)
    try:
        d.text((30, y), SAMPLE_TEXT_AR, fill="black", font=font)
    except Exception:
        pass
    img.save(path)
    return path


def classify(conf):
    """Map an average confidence (0-100) to status + action (spec section 6)."""
    if conf >= 90:  return ("Excellent",  "Continue normally")
    if conf >= 80:  return ("Acceptable", "Continue with minor warning")
    if conf >= 70:  return ("Weak",       "Continue but highlight for manual verification")
    return ("Failed", "Send for manual review")


def confidence_test(cfg=None, lang=None):
    """
    Run OCR on the internal sample and compute average word confidence.
    Returns {ok, confidence, status, action, text, error}.
    """
    cfg = cfg or {}
    lang = lang or cfg.get("ocr_default_language") or DEFAULT_LANGUAGE
    res = configure_pytesseract(cfg)
    if not res["exe"]:
        return {"ok": False, "confidence": 0.0, "status": "Failed",
                "action": "Send for manual review",
                "error": "No Tesseract executable resolved (bundled or system)."}
    try:
        import pytesseract
        tmp = os.path.join(tempfile.gettempdir(), "kt01_ocr_sample.png")
        _make_sample_image(tmp)
        cfgstr = tess_config(cfg, extra="--psm 6")
        data = pytesseract.image_to_data(tmp, lang=lang, config=cfgstr,
                                          output_type=pytesseract.Output.DICT)
        confs = [int(c) for c in data.get("conf", []) if str(c).lstrip("-").isdigit() and int(c) >= 0]
        text = pytesseract.image_to_string(tmp, lang=lang, config=cfgstr)
        avg = round(sum(confs) / len(confs), 1) if confs else 0.0
        status, action = classify(avg)
        return {"ok": True, "confidence": avg, "status": status, "action": action,
                "text": text.strip(), "lang": lang, "error": None}
    except Exception as e:
        return {"ok": False, "confidence": 0.0, "status": "Failed",
                "action": "Send for manual review", "error": str(e)}


def page_confidence(img_or_path, cfg=None, lang=None):
    """Average word confidence for one rendered page image (PIL image or path)."""
    cfg = cfg or {}
    lang = lang or cfg.get("ocr_default_language") or DEFAULT_LANGUAGE
    configure_pytesseract(cfg)
    try:
        import pytesseract
        cfgstr = tess_config(cfg, extra="--psm 6")
        data = pytesseract.image_to_data(img_or_path, lang=lang, config=cfgstr,
                                          output_type=pytesseract.Output.DICT)
        confs = [int(c) for c in data.get("conf", []) if str(c).lstrip("-").isdigit() and int(c) >= 0]
        return round(sum(confs) / len(confs), 1) if confs else 0.0
    except Exception:
        return 0.0


if __name__ == "__main__":
    import json
    print("RESOLVE:", json.dumps(resolve_ocr(), indent=2))
    print("HEALTH :", json.dumps(health_check(), indent=2, ensure_ascii=False))
    print("CONFID :", json.dumps(confidence_test(), indent=2, ensure_ascii=False))
