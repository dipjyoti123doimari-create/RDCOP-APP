"""
watchdog.py — keeps the RDC-OP Flask server alive forever.

Run this instead of app.py:
    python watchdog.py

It starts app.py as a child process. If the server exits for any reason
(crash, restart-button, etc.) it automatically restarts it after 2 seconds.
Windows Task Scheduler runs this at system startup so the app is always up.
"""

import os
import subprocess
import sys
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP      = os.path.join(BASE_DIR, "app.py")
PYTHON   = sys.executable

print("=" * 55)
print("  RDC-OP Watchdog — server will restart automatically")
print("  Press Ctrl+C to stop completely")
print("=" * 55)

while True:
    print(f"\n[watchdog] Starting server...")
    try:
        proc = subprocess.Popen([PYTHON, APP], cwd=BASE_DIR)
        proc.wait()
        code = proc.returncode
        if code == 0:
            print(f"[watchdog] Server exited cleanly (code 0). Restarting in 2s...")
        else:
            print(f"[watchdog] Server crashed (code {code}). Restarting in 2s...")
    except KeyboardInterrupt:
        print("\n[watchdog] Stopped by user.")
        try:
            proc.terminate()
        except Exception:
            pass
        sys.exit(0)
    except Exception as e:
        print(f"[watchdog] Error: {e}. Retrying in 5s...")
        time.sleep(5)
        continue
    time.sleep(2)
