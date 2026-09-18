@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
echo Run Setup.cmd first, then run this installer again.
pause
exit /b 1
)
".venv\Scripts\python.exe" "engine\startup.py" install
if errorlevel 1 (
echo Installation or background launch failed. Copy the error above for review.
pause
exit /b 1
)
pause
