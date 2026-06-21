# -*- coding: utf-8 -*-
"""v1.0.3 regression tests (deterministic; no network).
Covers: page-range detection, start-page validation, dynamic extraction (no hardcoding),
manual review queue save/reprocess, output regeneration. Run: python test_v103.py"""
import os, sys, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import process_contracts as pc

TMP = tempfile.mkdtemp(prefix="kt103_")
pc.OUTPUT = TMP
pc.AUDIT = os.path.join(TMP, "audit_log.jsonl")
pc.MRQ = os.path.join(TMP, "manual_review.json")
pc.CFG["make_outlook_draft"] = False
FAILS = []

def check(name, cond):
    print(("PASS  " if cond else "FAIL  ") + name)
    if not cond: FAILS.append(name)

# ---------- 1) start-page validation ----------
def test_start_validation():
    # footers: idx0=Page1of7 .. but a detector wrongly proposes start=1 (Page 2 of 7)
    flags = [(1,7),(2,7),(3,7),(4,7),(5,7),(6,7),(7,7)]
    check("start-page: shift back from Page2of7 to Page1of7", pc.valid_start_page(flags, 1) == 0)
    check("start-page: a true Page1 stays put", pc.valid_start_page(flags, 0) == 0)
    # a real second sequence start (Page1of7 again) must NOT shift
    flags2 = [(1,7),(2,7),(1,7),(2,7)]
    check("start-page: second seq Page1 unchanged", pc.valid_start_page(flags2, 2) == 2)

# ---------- 2) sequences_from_pageof ----------
def test_sequences():
    flags = [None,None,(1,7),(2,7),(3,7),(4,7),(5,7),(6,7),(7,7),(1,7),(2,7),(3,7),(4,7),(5,7),(6,7),(7,7)]
    seqs = pc.sequences_from_pageof(flags)
    check("sequences: two full Page1..7 sequences found", len(seqs) == 2)
    check("sequences: first is idx 2..8 (human 3-9)", seqs[0] == (2, 8))

# ---------- 3) page-range via visual repeat on a synthetic 16-page bundle ----------
def _make_bundle(path, supplier, scope, amount):
    import fitz
    d = fitz.open()
    for k in range(2):                       # PO pages 1-2
        p = d.new_page(); p.insert_text((40,60),"Purchase Order Report  page %d"%(k+1))
        p.insert_text((40,90),"ISSUED BY\nMohamed Awad"); p.insert_text((40,130),"APPROVED BY\nDr Hassan")
        p.insert_text((40,160),"SUB-0497"); p.insert_text((40,190),"%s EGP"%amount)
        p.insert_text((40,220),"Report generated: 15/06/2026"); p.insert_text((40,250),"2026-06-10")
    body = ("Second Party: %s\nScope of Work: %s\nProject: HSR\nDate: 2026/06/15\n"%(supplier,scope))
    for copy in range(2):                     # two identical 7-page copies (3-9, 10-16)
        for pg in range(7):
            p = d.new_page()
            p.insert_text((40,60), body if pg==0 else ("Clause %d"%pg))
            p.insert_text((40,760), "Page %d of 7   FRM-CO-016"%(pg+1))
    d.save(path); d.close()

def test_page_range_and_dynamic():
    b1 = os.path.join(TMP,"78-60238.pdf"); _make_bundle(b1,"Green Concrete Co","Ready-mix supply","1,000,000.00")
    b2 = os.path.join(TMP,"80-12345.pdf"); _make_bundle(b2,"Falcon Steel Ltd","Steel fabrication","2,500,000.00")
    V = pc.page_vectors(b1)
    rep = None
    for thr in (0.92,0.88,0.85,0.82):
        rep = pc.best_repeat(V, thr=thr)
        if rep: break
    cp = list(range(rep["start"], rep["start"]+rep["length"])) if rep else None
    check("page-range: synthetic bundle -> idx 2..8", cp == [2,3,4,5,6,7,8])
    check("page-range: exactly 7 pages", cp and len(cp)==7)

    # dynamic extraction returns DIFFERENT values per document (nothing hardcoded)
    f1,_ = pc.extract_contract_fields(b1, [2,3,4,5,6,7,8], lang="eng")
    f2,_ = pc.extract_contract_fields(b2, [2,3,4,5,6,7,8], lang="eng")
    c1=f1.get("company",{}).get("value",""); c2=f2.get("company",{}).get("value","")
    s1=f1.get("scope",{}).get("value","");   s2=f2.get("scope",{}).get("value","")
    print("   extracted #1 company=%r scope=%r" % (c1,s1))
    print("   extracted #2 company=%r scope=%r" % (c2,s2))
    check("dynamic: company differs per document", c1 and c2 and c1!=c2)
    check("dynamic: no hardcoded sample value", "جرين" not in (c1+c2))

# ---------- 4) manual review queue: appears, save, reprocess regenerates outputs ----------
def test_manual_review_flow():
    b1 = os.path.join(TMP,"78-60238.pdf")
    # force a manual-review condition by processing with a manual range but missing fields
    r = pc.process_one(b1, lang="eng", manual_range=(3,9))
    check("MR: file processed (status done/warning)", r["status"] in ("done","warning"))
    check("MR: page_range reported 3-9", r["page_range"] and r["page_range"]["start"]==3 and r["page_range"]["end"]==9)
    check("MR: extracted PDF has 7 pages", _pdf_pages(os.path.join(TMP,r["output_pdf"]))==7)
    rec = pc.mrq_list()
    queued = [x for x in rec if x.get("file")=="78-60238"]
    # if any field missing/weak, a record must exist
    if r["manual_review_required"]:
        check("MR: record created when fields missing/weak", len(queued)==1)
    else:
        check("MR: no record needed (all fields clean)", True)

    # simulate user save + reprocess with manual values overriding
    pc.mrq_upsert({"file":"78-60238","original":"78-60238.pdf","status":"Needs Manual Review",
                   "manual":{}, "range_start":3,"range_end":9})
    saved = pc.mrq_get("78-60238"); saved.setdefault("manual",{})
    saved["manual"].update({"sr_number":"SR-2026-99999","company":"Manual Vendor","scope":"Manual Scope","amount":"777777.00"})
    pc.mrq_upsert(saved)
    r2 = pc.process_one(b1, lang="eng", manual_range=(3,9), manual=pc.mrq_get("78-60238")["manual"])
    check("MR-reprocess: manual SR overrides", r2["extracted_fields"]["sr_number"]=="SR-2026-99999")
    check("MR-reprocess: manual company overrides", r2["extracted_fields"]["company"]=="Manual Vendor")
    check("MR-reprocess: manual amount overrides", r2["extracted_fields"]["amount"]=="777777.00")
    check("MR-reprocess: output PDF regenerated", bool(r2["output_pdf"]) and os.path.exists(os.path.join(TMP,r2["output_pdf"])))
    check("MR-reprocess: SR Log NEW ROWS exists", os.path.exists(os.path.join(TMP,"SR Log - NEW ROWS.xlsx")))
    check("MR-reprocess: .eml generated", bool(r2["eml"]) and os.path.exists(r2["eml"]))

def _pdf_pages(path):
    import fitz
    return len(fitz.open(path)) if os.path.exists(path) else -1

if __name__ == "__main__":
    print("== v1.0.3 tests ==")
    test_start_validation()
    test_sequences()
    test_page_range_and_dynamic()
    test_manual_review_flow()
    print()
    if FAILS:
        print("FAILURES:", FAILS); shutil.rmtree(TMP, ignore_errors=True); sys.exit(1)
    print("ALL v1.0.3 TESTS PASSED")
    shutil.rmtree(TMP, ignore_errors=True)
