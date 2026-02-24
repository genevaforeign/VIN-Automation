"""
vin_hotkey.py - Background hotkey listener for VIN automation.

Press Ctrl+Shift+V while a Pinnacle MVR is open to automatically decode
the VIN on VINMatchPro and fill all vehicle fields.

Usage:
    py -3.12 vin_hotkey.py
    (or double-click start_vin_hotkey.bat)

Press Ctrl+C in this window to stop.
"""

import os
import subprocess
import sys
import threading
import time

import keyboard

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON     = sys.executable
SCRIPT     = os.path.join(SCRIPT_DIR, 'vinmatch_to_pinnacle.py')
HOTKEY     = 'ctrl+shift+v'

_lock    = threading.Lock()
_running = False


def _run_fill():
    global _running
    with _lock:
        if _running:
            print('[hotkey] Already running — ignoring trigger.')
            return
        _running = True

    def _worker():
        global _running
        try:
            print(f'\n[hotkey] Ctrl+Shift+V pressed — launching fill script...')
            proc = subprocess.Popen(
                [PYTHON, SCRIPT],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
                cwd=SCRIPT_DIR,
            )
            proc.wait()
            print(f'[hotkey] Done (exit code {proc.returncode}). Ready.')
        except Exception as e:
            print(f'[hotkey] ERROR: {e}')
        finally:
            # Small cooldown so a double-tap doesn't re-fire immediately
            time.sleep(2)
            _running = False

    threading.Thread(target=_worker, daemon=True).start()


def main():
    print('=' * 50)
    print('  VIN Hotkey Listener  ')
    print('=' * 50)
    print(f'  Hotkey : Ctrl+Shift+V')
    print(f'  Script : {SCRIPT}')
    print()
    print('  Open a Pinnacle MVR, then press Ctrl+Shift+V')
    print('  to auto-fill all vehicle fields from VINMatchPro.')
    print()
    print('  Press Ctrl+C here to stop the listener.')
    print('=' * 50)
    print()

    keyboard.add_hotkey(HOTKEY, _run_fill, suppress=False)

    try:
        keyboard.wait()
    except KeyboardInterrupt:
        print('\nStopped.')


if __name__ == '__main__':
    main()
