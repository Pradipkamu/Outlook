@echo off
setlocal
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m unittest discover -s tests -v
if errorlevel 1 goto fail
echo Setup completed. Read START-HERE.md, then open Classic Outlook and run Start-Organizer.cmd.
pause
exit /b 0
:fail
echo Setup or tests failed. Do not enable automatic sending. Copy the error for review.
pause
exit /b 1
