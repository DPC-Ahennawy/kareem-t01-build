# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Kareem T-01 (one-folder build -> KareemT01.exe + _internal).
# Run from the app root:  pyinstaller packaging\KareemT01.spec
import os
APP = os.path.abspath(os.getcwd())

datas = [
    (os.path.join(APP, "refs.json"), "."),
    (os.path.join(APP, "recipients.json"), "."),
    (os.path.join(APP, "config.json"), "."),
    (os.path.join(APP, "web", "frontend", "index.html"), os.path.join("web", "frontend")),
]
hiddenimports = [
    "uvicorn", "uvicorn.logging", "uvicorn.loops", "uvicorn.loops.auto",
    "uvicorn.protocols", "uvicorn.protocols.http", "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets", "uvicorn.protocols.websockets.auto",
    "uvicorn.lifespan", "uvicorn.lifespan.on",
    "fastapi", "starlette", "anyio", "click", "h11",
    "fitz", "pytesseract", "PIL", "PIL.Image", "openpyxl", "numpy",
    "api", "process_contracts", "ocr_engine",
    "multipart", "python_multipart",
]

a = Analysis(
    [os.path.join(APP, "app_server.py")],
    pathex=[APP, os.path.join(APP, "web", "backend")],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[], runtime_hooks=[], excludes=["tkinter"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="KareemT01",
    console=True,          # keep a console window so the server log is visible
    icon=None,
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    name="KareemT01",      # -> dist\KareemT01\KareemT01.exe + _internal\
)
