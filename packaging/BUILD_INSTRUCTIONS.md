# Kareem T-01 — Build the production installer (`Kareem_T01_Setup.exe`)

This kit turns the source into a **no-Python Windows installer** with **bundled OCR**.
You run it **once** on any Windows PC that has internet. The target user PCs need **nothing** installed.

> Route B: the build script downloads the real Tesseract engine + ara/eng/fra data into
> `ocr\tesseract\` **before** packaging. The build **aborts** if those files are missing —
> it will never ship placeholders.

---

## Build machine prerequisites (one time)

1. **Python 3.10+ (64-bit)** — <https://www.python.org/downloads/> (tick "Add to PATH").
   *(Needed only to BUILD. The installed app does not need Python.)*
2. **Inno Setup 6** — <https://jrsoftware.org/isdl.php>
3. **7-Zip** (recommended, for clean engine extraction) — <https://www.7-zip.org/>
4. **Internet access** (to download Tesseract + traineddata).

> Build on the same architecture you ship to: 64-bit Windows → 64-bit app.

---

## Build steps

1. Unzip the source kit so you have the `Kareem T-01\` folder.
2. Open the `Kareem T-01\packaging\` folder.
3. Double-click **`build_all.bat`**.

`build_all.bat` runs, in order:

| Step | What it does |
|---|---|
| 1 | `pip install` PyInstaller + the app's runtime libraries |
| 2 | `fetch_ocr_full.ps1` → downloads **tesseract.exe + DLLs** and **eng/ara/fra.traineddata** into `ocr\tesseract\` |
| 2b | **Hard gate** — aborts if any OCR file is missing (no placeholder installer) |
| 3 | PyInstaller → `dist\KareemT01\KareemT01.exe` (one-folder build) |
| 4 | Inno Setup → `packaging\Output\Kareem_T01_Setup.exe` |

**Result:** `Kareem T-01\packaging\Output\Kareem_T01_Setup.exe` — this is the file you distribute.

---

## What the installer does on the user PC

- Installs **per-user** to `%LocalAppData%\EDECS\Kareem T-01\` (writable, **no admin, not Program Files**).
- Lays down `KareemT01.exe` + `_internal\` and the **`ocr\tesseract\`** folder beside it.
- Writes `config.json` with **`ocr_force_bundled: true`** → the app uses **bundled OCR only** (no system fallback).
- Creates **Desktop** + **Start Menu** shortcuts.
- On finish (and each launch) starts the local server and opens **http://127.0.0.1:8000/**.
- Stays fully **local/offline**.

Installed OCR layout:
```
%LocalAppData%\EDECS\Kareem T-01\ocr\tesseract\
    tesseract.exe
    *.dll
    tessdata\eng.traineddata
    tessdata\ara.traineddata
    tessdata\fra.traineddata
```

---

## OCR download source (Route B)

- **Engine** (`tesseract.exe` + DLLs): UB-Mannheim Tesseract build, pinned in `fetch_ocr_full.ps1`:
  `https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe`
  Extracted with 7-Zip (or silent install to a temp dir), then `tesseract.exe` + all `*.dll` are copied into `ocr\tesseract\`.
- **Language data**: official `tessdata_best`:
  `https://github.com/tesseract-ocr/tessdata_best/raw/main/{eng,ara,fra}.traineddata`

If the UB-Mannheim version URL changes, edit the `$url` line near the top of `fetch_ocr_full.ps1`.

---

## Verify after install (smoke test)

1. Launch from the Desktop shortcut → browser opens `http://127.0.0.1:8000/`.
2. **Settings → OCR Health Check** → all rows **Ready/Passed**, **Source: bundled**.
3. **Settings → OCR Confidence Test** → shows a % + band.
4. Drag a contract PDF onto Process → **Process Inbox** → check Output folder for the PDF, `SR Log - NEW ROWS.xlsx`, and the `.eml`.

---

## Troubleshooting

- **"OCR bundle incomplete / BUILD FAILED"** — no internet, or the engine URL changed. Fix connectivity or update `$url`, re-run `build_all.bat`.
- **"Inno Setup 6 not found"** — install it; the script looks in the default Program Files locations.
- **Engine extraction failed** — install 7-Zip and re-run (the script falls back to a silent install but 7-Zip is cleaner).
- **Antivirus flags the fresh exe** — PyInstaller exes are sometimes false-positived; sign the binary for production rollout.
