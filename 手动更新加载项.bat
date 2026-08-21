@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo  Updating and installing current Excel add-in
echo ========================================
echo Excel will be closed and old add-in registrations will be removed.
echo Save all workbooks before continuing.
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\install_debug_addin.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo Update failed. Exit code: %EXIT_CODE%
) else (
    echo Current add-in build, installation, and registration check completed.
)
pause
exit /b %EXIT_CODE%
