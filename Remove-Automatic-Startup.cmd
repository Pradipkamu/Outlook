@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
echo Open shell:startup from Windows Run and delete only Follow-up Organizer.lnk.
pause
exit /b 1
)
".venv\Scripts\python.exe" "engine\startup.py" remove
if errorlevel 1 (
echo Removal failed. Copy the error above for review.
pause
exit /b 1
)
pause
