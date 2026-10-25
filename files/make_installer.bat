@echo off
cd /d "%~dp0"
if not exist "dist\GharLeads\GharLeads.exe" (
  echo Build the program first with build.bat.
  pause
  exit /b 1
)
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo Could not find Inno Setup automatically.
  echo Find the file ISCC.exe on your PC, drag it into this window, then press Enter.
  set /p ISCC=
)
set ISCC=%ISCC:"=%
"%ISCC%" installer.iss
if errorlevel 1 (
  echo Something failed - send me a screenshot.
  pause
  exit /b 1
)
echo.
echo Installer ready: installer_output\GharLeads_Setup_0.4.exe
pause