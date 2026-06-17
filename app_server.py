# -*- coding: utf-8 -*-
"""
Kareem T-01 — EDECS Contract Processor
app_server.py  —  PyInstaller entry point for the packaged (no-Python) build.

When frozen by PyInstaller this becomes KareemT01.exe. It:
  * resolves the app root next to the executable (so bundled ocr/ and refs are found),
  * starts the local FastAPI server on http://127.0.0.1:8000/ (localhost only),
  * opens the browser to the UI.
Everything stays local/offline.
"""
import os, sys, threading, time, webbrowser

def app_root():
    # When frozen, sys.executable is .../KareemT01.exe ; data files sit beside it
    # (Inno installs the whole app folder, incl. ocr/, next to the exe).
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

ROOT = app_root()
os.environ["KT01_ROOT"] = ROOT
os.chdir(ROOT)
# Make the engine + backend importable regardless of frozen layout
for p in (ROOT, os.path.join(ROOT, "web", "backend"),
          getattr(sys, "_MEIPASS", "")):
    if p and p not in sys.path:
        sys.path.insert(0, p)

URL = "http://127.0.0.1:8000/"

def open_browser():
    time.sleep(2.0)
    try:
        webbrowser.open(URL)
    except Exception:
        pass

def main():
    # Import after sys.path/cwd are set so config + ocr resolve against ROOT
    import uvicorn
    import api  # FastAPI app (web/backend/api.py)
    threading.Thread(target=open_browser, daemon=True).start()
    print("Kareem T-01 running locally — open", URL)
    uvicorn.run(api.app, host="127.0.0.1", port=8000, log_level="warning")

if __name__ == "__main__":
    main()
