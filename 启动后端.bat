@echo off
setlocal

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\start_local_backend.ps1"
set "exit_code=%ERRORLEVEL%"

if not "%exit_code%"=="0" (
  echo.
  echo Backend startup failed with exit code %exit_code%.
)

pause

endlocal
exit /b %exit_code%
