@echo off
setlocal
cd /d "%~dp0"
".venv\Scripts\python.exe" "engine\backup.py"
pause
