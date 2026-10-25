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

echo Building GharLeads.exe (takes several minutes)...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto fail
set PLAYWRIGHT_BROWSERS_PATH=0
%PY% -m playwright install chromium
if errorlevel 1 goto fail
%PY% -m PyInstaller --noconfirm --windowed --name GharLeads --icon icon.ico --add-data "icon.ico;." --collect-all playwright gharleads.py
if errorlevel 1 goto fail
echo.
echo Program built: dist\GharLeads\GharLeads.exe
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%ISCC%" goto setup
echo Inno Setup not found - install it from jrsoftware.org to also make the Setup.exe installer.
goto end
:setup
"%ISCC%" installer.iss
if errorlevel 1 goto fail
echo Installer built: installer_output\GharLeads_Setup_0.4.exe
:end
pause
exit /b 0
:fail
echo.
echo Something failed - read the message above, then send me a screenshot.
pause