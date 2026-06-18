"""
watchdog.py — keeps the RDC-OP Flask server alive forever.

Run this instead of app.py:
    python watchdog.py

It starts app.py as a child process. If the server exits for any reason
(crash, restart-button, etc.) it automatically restarts it after 2 seconds.
Windows Startup folder runs this at login so the app is always up.
"""

import os
import subprocess
import sys
import time
from datetime import datetime

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
APP       = os.path.join(BASE_DIR, "app.py")
PYTHON    = sys.executable
CRASH_LOG = os.path.join(BASE_DIR, "server_crash.log")


def _log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(CRASH_LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


print("=" * 55)
print("  RDC-OP Watchdog — server will restart automatically")
print("  Press Ctrl+C to stop completely")
print(f"  Crash log: {CRASH_LOG}")
print("=" * 55)

_log("Watchdog started.")

while True:
    _log("Starting server (python app.py)...")
    try:
        proc = subprocess.Popen([PYTHON, APP], cwd=BASE_DIR)
        proc.wait()
        code = proc.returncode
        if code == 0:
            _log(f"Server exited cleanly (code 0). Restarting in 2 s...")
        else:
            _log(f"Server crashed with exit code {code}. Check server.log for traceback. Restarting in 2 s...")
    except KeyboardInterrupt:
        _log("Watchdog stopped by user (Ctrl+C).")
        try:
            proc.terminate()
        except Exception:
            pass
        sys.exit(0)
    except Exception as e:
        _log(f"Watchdog error: {e}. Retrying in 5 s...")
        time.sleep(5)
        continue
    time.sleep(2)
