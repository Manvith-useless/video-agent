@echo off
title PO Tracker
echo.
echo  Starting PO Tracker...
echo  Do not close this window while using the app.
echo.

cd /d "%~dp0"
python run.py

pause
