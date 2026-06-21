@echo off
setlocal
cd /d "%~dp0"
echo Uninstalling EqNB from Word...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Uninstall-EqNB.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
  echo Uninstall failed. Please close Word and try again.
) else (
  echo Uninstall finished. Restart Word to unload EqNB.
)
echo.
pause
exit /b %ERR%
