# -*- coding: utf-8 -*-
"""Deterministic tests for parse_bundle_name (no AI, no network). Run: python test_filename_parser.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import process_contracts as pc

# (filename) -> expected (code, no, annex, manual)
CASES = [
    # 1) normal
    ("107-59659.pdf",                ("107", "59659", None, None)),
    ("98-60033.pdf",                 ("98",  "60033", None, None)),
    ("104-48113.pdf",                ("104", "48113", None, None)),
    ("78-60238.pdf",                 ("78",  "60238", None, None)),
    ("107-59673.pdf",                ("107", "59673", None, None)),
    # 2) annex without space
    ("72-49835 Annex(1).pdf",        ("72",  "49835", 1,    None)),
    ("97-49659 Annex(8).pdf",        ("97",  "49659", 8,    None)),
    ("104-48113 Annex(10).pdf",      ("104", "48113", 10,   None)),
    ("105-54139 Annex(2).pdf",       ("105", "54139", 2,    None)),
    # 3) annex with space
    ("78-48033 Annex (1).pdf",       ("78",  "48033", 1,    None)),
    ("78-48033 Annex (2).pdf",       ("78",  "48033", 2,    None)),
    ("105-55517 Annex (1).pdf",      ("105", "55517", 1,    None)),
    # 4) sub-project codes
    ("099-02-26700 Annex(1).pdf",    ("99-02", "26700", 1,  None)),
    ("099-02-26700 Annex (1).pdf",   ("99-02", "26700", 1,  None)),
    # 5) letter project codes
    ("CON-01-57697.pdf",             ("CON-01", "57697", None, None)),
    # 6) manual page range
    ("099-02-26700 [3-4].pdf",       ("99-02", "26700", None, (3, 4))),
    ("099-02-26700 Annex(1) [3-4].pdf", ("99-02", "26700", 1, (3, 4))),
    # tolerance extras: case, extra spaces, addendum word, no extension
    ("72-49835  ANNEX (3) .pdf",     ("72",  "49835", 3,    None)),
    ("105-54139 annex(4)",           ("105", "54139", 4,    None)),
    ("99-1-12345 Annex (2).pdf",     ("99-01", "12345", 2,  None)),
]

FAIL_CASES = ["", "random_document.pdf", "notes.txt", "Annex (1).pdf"]

def run():
    ok = 0; bad = []
    for fn, exp in CASES:
        r = pc.parse_bundle_name(fn)
        got = (r["code"], r["no"], r["annex"], r["manual"]) if r else None
        if got == exp:
            ok += 1
            print("PASS  %-34s -> code=%s no=%s annex=%s range=%s" % (fn, *exp))
        else:
            bad.append((fn, exp, got))
            print("FAIL  %-34s -> expected %s got %s" % (fn, exp, got))
    for fn in FAIL_CASES:
        r = pc.parse_bundle_name(fn)
        if r is None:
            ok += 1; print("PASS  (rejected) %r" % fn)
        else:
            bad.append((fn, None, r)); print("FAIL  should reject %r got %s" % (fn, r))
    # annex_label normalisation
    assert pc.annex_label(1) == "Annex (1)" and pc.annex_label(None) == ""
    print("PASS  annex_label normalisation")
    total = len(CASES) + len(FAIL_CASES) + 1
    print("\n%d/%d checks passed." % (ok + 1, total))
    if bad:
        print("FAILURES:", bad); sys.exit(1)
    print("ALL FILENAME PARSER TESTS PASSED")

if __name__ == "__main__":
    run()
