# Kareem T-01 — EDECS Contract Processor

Local, offline desktop/web application that automates processing of construction‑procurement
contract PDF bundles for the **EDECS Procurement / DPC Department**.

> Developed for the DPC Department in EDECS_EGY
> By Kareem Talaat & Alaaeldien Alhennawy

All processing stays on the machine. **No contract data is uploaded to any external service.**
(The only online step is the optional one‑time `fetch_ocr.bat`, which downloads the OCR engine/data.)

---

## 1. What it does

For each bundle PDF named `<projectCode>-<contractNo>[ Annex(N)]` (e.g. `100-56815.pdf`,
`099-02-26700 Annex (1).pdf`, `CON-01-57697.pdf`) the app:

1. Detects and extracts **one** contract copy into a correctly‑named PDF.
2. Builds a fully‑mapped 26‑column SR Log row.
3. Writes new rows to a **separate** file — `SR Log - NEW ROWS.xlsx` — never touching the shared master.
4. Generates an editable `.eml` draft (`X-Unsent: 1`) with the contract attached (or a real Outlook draft).
5. Routes weak/failed OCR or undetected page ranges to **Manual Review**.
6. Logs everything to a structured audit log.

---

## 2. Run instructions (local)

Requires **Python 3.10+** on Windows.

```bat
1) Setup.bat        :: installs Python libraries (one time)
2) fetch_ocr.bat    :: downloads bundled OCR engine + ara/eng/fra data (one time, online)
3) Run.bat          :: starts the local app and opens http://127.0.0.1:8000/
```

`Run.bat` launches a local FastAPI server (`uvicorn`) bound to `127.0.0.1` only and opens the
browser UI. Nothing is exposed to the network.

Manual start (equivalent to Run.bat):

```bat
cd web\backend
py -m uvicorn api:app --host 127.0.0.1 --port 8000
```

---

## 3. Bundled OCR — how the folder works

The app does **not** require a manual system install of Tesseract. It resolves OCR in this order:

1. **Bundled** (primary target):
   ```
   Kareem T-01/ocr/tesseract/tesseract.exe
   Kareem T-01/ocr/tesseract/tessdata/eng.traineddata
   Kareem T-01/ocr/tesseract/tessdata/ara.traineddata
   Kareem T-01/ocr/tesseract/tessdata/fra.traineddata
   ```
2. **System fallback** (only if bundled is missing and `ocr_force_bundled` is `false`):
   `C:\Program Files\Tesseract-OCR\tesseract.exe` or any `tesseract` on `PATH`.

Set `"ocr_force_bundled": true` in `config.json` to disable the fallback and require the bundled copy
(recommended for the production rollout).

### Getting the OCR binaries

`fetch_ocr.bat` downloads the three `*.traineddata` files automatically into
`ocr/tesseract/tessdata/`. The **engine** (`tesseract.exe` + its DLLs) is not auto‑extracted; pick one:

- **A.** Install the UB‑Mannheim Tesseract build (<https://github.com/UB-Mannheim/tesseract/wiki>),
  then copy its `tesseract.exe` and accompanying DLLs into `ocr/tesseract/`. → fully bundled.
- **B.** Leave it and let the app fall back to a system‑installed Tesseract during transition.

For the final installer (Milestone 5), place the engine + 3 traineddata files into
`ocr/tesseract/` before packaging so the shipped build is self‑contained.

Verify any time in the app: **Settings → OCR Health Check**.

### OCR languages (Settings → Default OCR language, or the Process screen selector)

| UI option | Tesseract `-l` |
|---|---|
| Arabic | `ara` |
| English | `eng` |
| French | `fra` |
| Arabic + English (default) | `ara+eng` |
| Arabic + English + French | `ara+eng+fra` |

`ara+eng` is a runtime combination of `ara.traineddata` + `eng.traineddata`, not a separate file.

---

## 4. OCR Confidence → routing

| Confidence | Status | Action |
|---:|---|---|
| ≥ 90% | Excellent | Continue normally |
| 80–89% | Acceptable | Continue with minor warning |
| 70–79% | Weak | Continue, highlight for manual verification |
| < 70% | Failed | **Send to Manual Review** |

Files below 70% are not silently finalised — they appear in **Pending / Manual Review**, where you
can set a page range `[s-e]`, correct the Agreement Date / SR Number, choose a language, and reprocess.

---

## 5. Optional manual fields

On the Process screen (and per‑file in Manual Review) you can supply **Agreement Date** and
**SR Number**. Manual values override OCR. If either is missing after OCR + manual entry, the app
**continues anyway**, highlights the field, adds a note to the SR Log row, and records a warning in
the audit log. Processing is never blocked solely by a missing Agreement Date or SR Number.

---

## 6. Output files (in the Output folder)

- `Agreement[ Annex (N)]-PJ<code>_<no> <label>-<vendorEng>-<scope>-<Mon YYYY>_<vendorArabic>.pdf`
- `SR Log - NEW ROWS.xlsx` — sheet `New Rows`, 26 columns, missing fields shaded yellow.
- `<output name>.eml` — editable unsent draft (or a real Outlook draft if enabled).
- `audit_log.jsonl` — one JSON record per processed file + per manual override.
- `_pending.txt` — manual‑review queue.

The shared master `SR Log 2026.xlsx` is opened **read‑only** (for de‑duplication on
`Project Code + PO/Contract No. + Type`) and is never written to.

---

## 7. Smoke test (one PDF)

1. `Setup.bat` → `fetch_ocr.bat` → `Run.bat`. Browser opens the app.
2. **App opens** — sidebar shows Process / Pending / Emails / History / Settings; footer shows the
   DPC / Kareem Talaat & Alaaeldien Alhennawy credit.
3. **OCR Health Check** — Settings → OCR Health Check → all rows `Ready` / `Passed`.
4. **OCR Confidence Test** — Settings → OCR Confidence Test → shows a % and status.
5. **Upload works** — drag a bundle (e.g. `100-56815.pdf`) onto the drop zone; "In Inbox" increments.
6. Choose an OCR language (optional Agreement Date / SR Number) → **Process Inbox**.
7. **Processing runs** — the row updates live with status + OCR confidence; progress bar reaches 100%.
8. Open the Output folder and confirm:
   - the extracted contract **PDF** exists with the correct name,
   - **`SR Log - NEW ROWS.xlsx`** was created/appended (26 columns),
   - the **`.eml`** draft was created,
   - **History** lists the run with timestamp, confidence, and links.

A self‑contained engine test (synthetic 2‑copy bundle) was run during the build and produced all four
outputs with detection of 2 copies and 94.8% confidence.

---

## 8. Configuration (`config.json`)

```json
{
  "sync_dir": "",                     // blank = app folder; where Inbox/Output/SR Log live
  "output_dir": "",                   // blank = <sync_dir>/Output
  "ocr_mode": "bundled",
  "ocr_default_language": "ara+eng",
  "ocr_dpi": 300,
  "ocr_confidence_test": true,
  "ocr_force_bundled": false,         // true = require bundled OCR, no system fallback
  "name_separator": "_",              // separator before the Arabic vendor name
  "make_outlook_draft": true          // false = always produce .eml
}
```

All paths are editable from **Settings** (restart to apply folder changes).

---

## 9. Changed / new files

**New**
- `ocr_engine.py` — bundled‑OCR resolution, health check, confidence test, language map.
- `web/frontend/index.html` — rewritten full UI (5 screens, upload/drag‑drop, OCR tools, branding/footer).
- `fetch_ocr.bat` — one‑time OCR download/setup.
- `ocr/tesseract/tessdata/PLACEHOLDER.txt` — bundled‑OCR folder structure.
- `README.md` — this file.

**Modified**
- `process_contracts.py` — added bundled OCR, language selection, per‑file confidence + <70% Manual
  Review routing, manual Agreement Date / SR Number overrides, missing‑field highlighting + notes,
  audit log, richer `process_one` return for the API. Detection & enrichment logic preserved.
- `web/backend/api.py` — added upload, per‑file process, OCR health/confidence, settings GET/POST,
  history (audit), manual‑review reprocess, file download.
- `config.json` — switched to bundled‑OCR mode + new keys.
- `Run.bat` — now launches the local web app (uvicorn) instead of the Tkinter script.
- `Setup.bat` — installs the web stack (FastAPI/uvicorn/python‑multipart) too.

**Unchanged**
- `refs.json`, `recipients.json` — reference data (editable via the UI for recipients).
- `gui_app.py` — kept only as a legacy Tkinter fallback; not the production interface.

---

## 10. Build / packaging (for rollout)

- Place the OCR engine + 3 traineddata files into `ocr/tesseract/` and set
  `"ocr_force_bundled": true`.
- Package with PyInstaller (one‑folder) or ship the folder + an embeddable Python; include the `ocr/`
  tree as data. Include licence notices for Tesseract (Apache‑2.0) and the Python libraries.
- This README covers run/build; full user/admin manuals + release notes are the later Milestone 5 deliverable.

---

## 11. Known limitations / pending items

- **OCR binaries are not committed** to this build (large platform binaries). Run `fetch_ocr.bat` or
  place them manually before first real use; the engine `tesseract.exe` step is manual (option A/B above).
- **French/Arabic confidence** depends on scan quality; the confidence test image is a synthetic sample.
- **Outlook drafts** require Windows + Outlook + `pywin32`; otherwise the app produces `.eml`.
- **De‑duplication** relies on reading the master read‑only; if the master is unreachable, the app still
  writes new rows (no enrichment from the sheet in that case).
- **Manual page range** uses 1‑based inclusive pages (`[s-e]`).
- Detection accuracy for unusual bundle layouts may still need a manual range — by design these route to
  Manual Review rather than producing a wrong extraction.

---

## 12. Production installer (no-Python build)

The supported production deployment is the **Windows installer** `Kareem_T01_Setup.exe`, not the
source build. See `packaging\BUILD_INSTRUCTIONS.md`.

- **The source kit does NOT contain the OCR binaries.** During the build, `build_all.bat` runs
  `fetch_ocr_full.ps1`, which **downloads the real Tesseract engine (`tesseract.exe` + DLLs) and
  `eng/ara/fra.traineddata`** into `ocr\tesseract\` before packaging. The build **aborts** if any
  are missing — no placeholder installer is ever produced.
- **The final `Kareem_T01_Setup.exe` DOES include bundled OCR** (engine + DLLs + 3 languages) inside
  the installed app folder.
- Installs per-user to `%LocalAppData%\EDECS\Kareem T-01\` (writable, no admin, not Program Files).
- Packaged `config.json` sets **`ocr_force_bundled: true`** → the installed app uses bundled OCR only
  and **does not fall back to a system-installed Tesseract**. OCR Health Check reports **Source: bundled**.
- The installed app needs **no Python** on the user PC.

---

## 13. Cloud build via GitHub Actions (no Python on your PC)

The production installer is built automatically in the cloud on a Windows runner. You never run
Python, PyInstaller, or Inno Setup yourself.

**What the workflow does** (`.github/workflows/build-installer.yml`, runs on `windows-latest`):

1. Checks out the repo and installs Python **on the runner** (build-only; the app ships without Python).
2. Installs the Python build/runtime libraries.
3. Installs the **real Tesseract engine + DLLs** (via Chocolatey) and copies them into `ocr\tesseract\`.
4. Downloads **eng / ara / fra** `traineddata` (official `tessdata_best`) into `ocr\tesseract\tessdata\`.
5. **Verifies** `tesseract.exe`, the DLLs, and all three languages exist — **fails the build** if any are missing.
6. Applies `packaging\config.packaged.json` (sets `ocr_force_bundled: true`).
7. Runs **PyInstaller** → `KareemT01.exe`.
8. Runs **Inno Setup** → `Kareem_T01_Setup.exe`.
9. Uploads **`Kareem_T01_Setup.exe`** as a downloadable workflow artifact.

### Steps for you (browser only)

1. **Create a repository** — github.com → New repository → name it e.g. `kareem-t01` → Create.
2. **Upload the files** — on the repo page click **Add file → Upload files**, drag in the whole
   `Kareem T-01` folder contents (so `.github/`, `packaging/`, `process_contracts.py`, etc. sit at the
   repo root), then **Commit changes**.
   *(Keep the `.github` folder — that's the workflow.)*
3. **Open Actions** — click the **Actions** tab. If prompted, enable workflows.
4. **Run the workflow** — select **Build Kareem_T01_Setup.exe** → **Run workflow** → **Run workflow**.
   (It also runs automatically on every push to `main`.)
5. **Download the artifact** — when the run finishes (green check), open it and download
   **`Kareem_T01_Setup`** under *Artifacts*. Inside is **`Kareem_T01_Setup.exe`**.

Then double-click `Kareem_T01_Setup.exe` on any target PC — it installs to
`%LocalAppData%\EDECS\Kareem T-01\`, needs no admin and no Python, bundles OCR, makes Desktop +
Start Menu shortcuts, and opens `http://127.0.0.1:8000/`.
