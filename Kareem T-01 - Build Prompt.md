# BUILD PROMPT — "Kareem T-01" : Procurement Contract Processor

Build a desktop/web application named **"Kareem T-01"** that automates processing of construction-procurement contract PDFs for the EDECS Procurement Department. Paste this whole document as the spec. Implement BOTH the backend (processing engine) and the frontend (UI). Brand footer on EVERY screen (bottom, centered):

```
Created By : DPC Department
By Kareem Talaat
```

Theme colors: **Gold `#A5872F`** and **Green `#155F65`** (primary buttons & progress = green; logo, active tab & accents = gold; white background, light-gray sidebar `#f3f1ea`).

---

## 1. PURPOSE
Each input is a **bundle PDF** named `<projectCode>-<contractNo>[ Annex(N)]` (e.g. `100-56815`, `099-02-26700 Annex (1)`, `CON-01-57697`). A bundle contains: a **Purchase-Order (PO) report page** + the **contract** (often duplicated as 2 identical copies) + attachments (quotations/comparisons). For each bundle the app must:
1. Detect and extract **ONE** contract copy into a new, correctly-named PDF.
2. Append a fully-mapped row to a log spreadsheet (`SR Log 2026.xlsx`, sheet `SR Log`, 26 columns) — written to a SEPARATE output file, never locking the shared master.
3. Generate an editable email draft (`.eml`, header `X-Unsent:1`) with the contract attached.

## 2. INPUT / REFERENCE DATA
- Folder with `Inbox/` (bundles to process) and `Output/` (results) and `SR Log 2026.xlsx`.
- `SR Log 2026.xlsx` sheets: `SR Log` (the 26-col log), `Vendors` (account→exact Arabic name, 5800+ rows), `Data` (project code→name, category/division vocab), `ERP User` (person names).
- Two bundled JSON reference files built once from historical named files:
  - `refs.json`: `kb` = map `"<code>|<no>" → {label, veng(English vendor), work, arabic(short), my("Mon YYYY")}`; `label` = code→project short label; `ar_eng` = short-Arabic→English vendor.
  - `recipients.json`: project code → `{to, cc}` (latest email recipients per project).

## 3. CONTRACT-PAGE DETECTION (in this order; first hit wins)
1. **Visual repeat:** render each page to a small normalized grayscale vector (e.g. 36×50, z-scored); find the longest contiguous page-block that repeats later with average cosine-similarity **≥ 0.92** → that block is copy #1; copies = number of repeats. (Handles old 2-copy bundles.)
2. **FRM footer:** OCR (English) the bottom ~14% strip of every page; pages containing `FRM` (contract form code, e.g. `FRM-CO-004`) form the contract; take the longest run (bridge single-page gaps). copies=1. (Handles modern single-copy multi-page contracts.)
3. **Contract-number token:** OCR first ~8 pages; the contract page contains `"<no>/<code>"` (e.g. `26700/099-02`); take the contiguous run. (Handles short annexes; the PO page does NOT contain this slashed token.)
4. **Manual override:** if filename contains `[s-e]` (e.g. `... [3-4].pdf`) use pages s..e directly. If none of the above succeed → mark **Pending** and ask the user to add `[s-e]`.

Code normalization (`ncode`): strip leading zeros from the numeric base, zero-pad sub-project to 2 digits, keep letter codes (`CON-01`, `N-Store`). Unify `99-1`≡`99-01`.

## 4. PO-REPORT EXTRACTION (English OCR of the PO page)
- `vendor_account`: regex `(SUB|SER|CU)-0*(\d+)` → e.g. `SUB-00497`.
- `net_amount`: gross (largest `#,###.## EGP`) minus the discount lines.
- `po_date`: `YYYY-MM-DD`; `report_date`: `Report generated: DD/MM/YYYY`.
- `issued_by` (→ Person): take the text after `ISSUED BY` and match it against the known person-name list (ERP User + recipient names); pick the **right-most** matching full name (longer name wins ties).
- `approved_by`: name after `APPROVED BY`.

## 5. ENRICHMENT (combine KB + Vendors + SR Log history + ar_eng)
For `(code,no)`: `ref = kb["code|no"]`. Then:
- **Vendor (Excel cell):** exact name from `Vendors` by account; **Vendor (filename):** short Arabic from ref.
- **Vendor English:** ref.veng, else `ar_eng` lookup on the Arabic with Arabic normalization (remove diacritics; أإآ→ا, ة→ه, ى→ي) matching on the **first two words**.
- **Scope / Division / Category:** from ref.work, else from the same vendor's previous `SR Log` rows (by account); Division = Arabic work; Category from vocab (Equipment→`Logistic - Equ Rent`, Consultancy→`Consultant agreement`, else `Subcontractor agreement`).
- **Date for name/subject:** read a date (`YYYY/MM/DD` or `DD/MM/YYYY`) from the contract's FIRST page (Arabic+English OCR), else po_date, else ref.my → format `Mon YYYY`.

## 6. OUTPUT FILE NAME
`Agreement[ Annex (N)]-PJ<code>_<no> <projectLabel>-<vendorEnglish>-<scope>-<Mon YYYY>_<vendorArabicShort>.pdf`
Separator before the Arabic name configurable (default `_`). Example:
`Agreement-PJ100_56815 New Capital Stabling Yard-A Build-Insulation-Jun 2026_اية بيلد.pdf`

## 7. SR LOG ROW (26 columns, in order)
`Project Code | Project Name | Request Type | No. | Create Date | Cost Control Sign Date | Procurement Sign Date | Person | Scope | Category | Vendor account | Vendor | Division | PO / Contract No. | Type | Date | Amount | Procurement Sign Date2 | Cost Control Sign Date2 | Chairman Sign Date | Distribution Date | Status | Note | Archive Copies No. | Archive Vendor | Archive Date`
Fill: ProjectCode=code, ProjectName=Data lookup, CreateDate=contract date, Person=issued_by, Scope, Category, Vendor account, Vendor(exact Arabic), Division, No.=contractNo, Type=`Contract` or `Addendum N`, Date=contract date, Amount=net, Chairman & Distribution = report_date, Status=`Open`, Note=approved_by, Archive Copies No.=copies. Leave blank: Request Type, No.(SR), the two sign dates. **De-dup** by (ProjectCode, No., Type) against the master read-only; write new rows to a SEPARATE `SR Log - NEW ROWS.xlsx` (master is shared/locked).

## 8. EMAIL DRAFT (.eml, X-Unsent:1)
- To/Cc from `recipients.json` for the project; **exclude "Logistics" from Cc unless** work = Equipment.
- Subject = output file name (without extension).
- Body (NO signature, ends at "Best Regards,"):
```
Dear All,

Kindly find attached Subcontractor's Contract detailed as follows:-

Project          : <Project Name> (PJ <code>).
Contractor       : <vendor English>.
Contract Subject : <scope>.
Copies           : <One/Two...> Original(s).

This is for your information and Record.

Best Regards,
```
On Windows with Outlook installed, optionally create a real draft via Outlook automation instead of `.eml`.

## 9. FRONTEND (5 screens, sidebar navigation)
Left sidebar (gold logo "EDECS" / "Kareem T-01"), items: **Process / Pending / Emails / History / Settings**; active item highlighted gold. Footer on every screen.
- **Process:** Inbox path + Scan button; 3 metric cards (In Inbox / Done / Review); a table of bundles (File | Status) updated live per file; a **green progress bar with %** and `done / total (pct%)`; big green "Process" button. On run, process each Inbox file, update the row status (✅ done / ⚠️ review) and cards live.
- **Pending:** list files that need a manual `[s-e]` range (with instructions).
- **Emails:** choose an existing project **or type a new project code**; edit `To` and `Cc` (one recipient per line; add a line to add someone, delete a line to remove); Save writes `recipients.json`.
- **History:** list separated contracts in Output.
- **Settings:** editable paths (sync folder, Tesseract path), name separator, Outlook-draft toggle; Save writes `config.json` (restart to apply path changes).

## 10. TECH STACK (suggested)
- Desktop: Python 3.13 + Tkinter (current build) OR a web app (React frontend + FastAPI/Node backend).
- PDF render + split: PyMuPDF. OCR: Tesseract (`eng` + `ara` language packs). Spreadsheet: openpyxl. Image vectors: NumPy + Pillow.
- Self-contained; reads the local sync folder; no external upload of contract data (privacy).

## 11. EDGE CASES / RULES
- Re-runs are idempotent (de-dup). If a project has no recipients → leave To/Cc blank. If OCR can't read a date → fall back to PO date. New contracts not in KB still get vendor(Arabic+English), amount, dates, scope from the PO + Vendors + SR-Log history. Names must always be full (match against the known-names list). Never modify or bypass company security; the app runs where the user is allowed to run it.

DELIVER: a working app matching the above, themed gold/green, with the DPC/Kareem-Talaat footer on every screen.
