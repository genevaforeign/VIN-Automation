@echo off
title VIN Hotkey Listener
echo Starting VIN Hotkey Listener...
echo Press Ctrl+Shift+V in Pinnacle to auto-fill MVR fields.
echo.
py -3.12 "%~dp0vin_hotkey.py"
pause
