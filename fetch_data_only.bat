@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 python -m pip install -r requirements.txt
py scripts\fetch_live_jgsa.py
if errorlevel 1 python scripts\fetch_live_jgsa.py
pause
