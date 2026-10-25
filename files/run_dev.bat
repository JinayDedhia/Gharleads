@echo off
cd /d "%~dp0"
set PY=
for %%V in (3.12 3.13 3.11 3.10) do (
  if not defined PY (py -%%V --version >nul 2>nul && set "PY=py -%%V")
)
if not defined PY set "PY=python"
echo Using: %PY%
%PY% --version
if errorlevel 1 (
  echo.
  echo Python was not found. Install Python 3.12 from python.org and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

echo Installing packages (first time takes a few minutes)...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto fail
%PY% -m playwright install chromium
if errorlevel 1 goto fail
echo.
echo Starting GharLeads...
%PY% gharleads.py
pause
exit /b 0
:fail
echo.
echo Something failed - read the message above, then send me a screenshot.
pause