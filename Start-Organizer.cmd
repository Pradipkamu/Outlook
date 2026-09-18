@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
echo Run Setup.cmd first.
pause
exit /b 1
)
start "Follow-up watchdog" /min ".venv\Scripts\python.exe" "engine\watchdog.py"
".venv\Scripts\python.exe" "engine\service.py"
pause
