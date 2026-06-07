@echo off
setlocal
cd /d "%~dp0"
echo Installing requirements...
py -m pip install -r requirements.txt
if errorlevel 1 python -m pip install -r requirements.txt
echo.
echo Fetching live JGSA data. Keep internet ON.
py scripts\fetch_live_jgsa.py
if errorlevel 1 python scripts\fetch_live_jgsa.py
echo.
echo Starting local server at http://localhost:8000
start http://localhost:8000
py -m http.server 8000
if errorlevel 1 python -m http.server 8000
