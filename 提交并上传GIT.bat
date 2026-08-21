@echo off
if /i not "%~1"=="__RUN" (
  start "Excel加载项 - 提交并上传GIT" "%ComSpec%" /k ""%~f0" __RUN"
  exit /b
)
shift
setlocal
title Excel加载项 - 提交并上传GIT
set "PULL_ATTEMPT=0"

cd /d "%~dp0"

set "GIT_EXE=git"
for %%I in ("%~dp0.") do set "REPO_DIR=%%~fI"
%GIT_EXE% --version >nul 2>nul
if errorlevel 1 if exist "%ProgramFiles%\Git\cmd\git.exe" set "GIT_EXE=%ProgramFiles%\Git\cmd\git.exe"
if errorlevel 1 if exist "%ProgramFiles(x86)%\Git\cmd\git.exe" set "GIT_EXE=%ProgramFiles(x86)%\Git\cmd\git.exe"
if errorlevel 1 if exist "%LocalAppData%\Programs\Git\cmd\git.exe" set "GIT_EXE=%LocalAppData%\Programs\Git\cmd\git.exe"

"%GIT_EXE%" --version >nul 2>nul
if errorlevel 1 (
  echo Git is not installed or not available in PATH.
  goto :failed
)

"%GIT_EXE%" -c safe.directory=* -C "%CD%" rev-parse --show-toplevel >nul 2>nul
if errorlevel 1 (
  echo This folder is not a Git repository:
  echo "%~dp0"
  echo Please clone the project here, or restore the .git folder, then run this script again.
  goto :failed
)

echo Pulling latest changes from Git...
:pull_retry
"%GIT_EXE%" -c safe.directory=* pull --ff-only
if not errorlevel 1 goto pull_complete

set /a PULL_ATTEMPT+=1
if %PULL_ATTEMPT% GEQ 3 goto pull_failed
echo.
echo Git pull failed. Retrying in 5 seconds (%PULL_ATTEMPT%/3)...
powershell -NoProfile -Command "Start-Sleep -Seconds 5"
goto pull_retry

:pull_failed
if errorlevel 1 (
  echo.
  echo Git pull failed. Please resolve the message above before uploading.
  echo If the message says it cannot connect to github.com:443, check your network,
  echo VPN/proxy settings, firewall, or GitHub access, then run this script again.
  goto :failed
)

:pull_complete

"%GIT_EXE%" -c safe.directory=* status --porcelain > "%TEMP%\excel_addin_git_status.txt"
for %%A in ("%TEMP%\excel_addin_git_status.txt") do set STATUS_SIZE=%%~zA
del "%TEMP%\excel_addin_git_status.txt" >nul 2>nul

if "%STATUS_SIZE%"=="0" (
  echo.
  echo No local changes to commit.
  goto :done
)

echo.
echo Staging local changes...
"%GIT_EXE%" -c safe.directory=* add .
if errorlevel 1 (
  echo.
  echo Git add failed.
  goto :failed
)

for /f %%A in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set SYNC_TIME=%%A
set COMMIT_MESSAGE=Auto sync %SYNC_TIME%

echo.
echo Committing: %COMMIT_MESSAGE%
"%GIT_EXE%" -c safe.directory=* commit -m "%COMMIT_MESSAGE%"
if errorlevel 1 (
  echo.
  echo Git commit failed.
  goto :failed
)

echo.
echo Pushing changes to GitHub...
"%GIT_EXE%" -c safe.directory=* push
if errorlevel 1 (
  echo.
  echo Git push failed. Your local commit was created, but it was not uploaded.
  goto :failed
)

echo.
echo Local changes have been committed and uploaded.
goto :done

:failed
echo.
echo Script finished with errors. Press any key to close.

:done
pause

endlocal
