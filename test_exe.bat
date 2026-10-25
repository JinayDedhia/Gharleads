@echo off
cd /d "%~dp0"
if not exist "dist\GharLeads\GharLeads.exe" (
  echo The program has not been built yet. Run build.bat first, wait until it finishes, then run this again.
  pause
  exit /b 1
)
echo Running the GharLeads self-test. Please wait up to 30 seconds - a message box will appear ...
start "" /wait "dist\GharLeads\GharLeads.exe" --selftest
echo.
echo Self-test finished. Check the message box that appeared.
pause