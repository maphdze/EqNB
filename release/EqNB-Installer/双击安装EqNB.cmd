@echo off
setlocal
cd /d "%~dp0"
echo Installing EqNB for Word...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Install-EqNB.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
  echo Install failed. Please close Word and try again.
) else (
  echo Install finished. Restart Word and open the EqNB tab.
)
echo.
pause
exit /b %ERR%
