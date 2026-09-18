@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
echo Run Setup.cmd first. Do not launch the organizer yet.
pause
exit /b 1
)
".venv\Scripts\python.exe" "engine\migrate_data.py"
if errorlevel 1 (
echo Migration did not finish. Read the error above; keep the old data folder.
pause
exit /b 1
)
pause
