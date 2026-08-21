@echo off
if /i not "%~1"=="__RUN" (
  start "Excel加载项 - 更新GIT" "%ComSpec%" /k ""%~f0" __RUN"
  exit /b
)
shift
setlocal
title Excel加载项 - 更新GIT

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
"%GIT_EXE%" -c safe.directory=* pull --ff-only
if errorlevel 1 (
  echo.
  echo Git pull failed. Please check the message above.
  goto :failed
)

echo.
echo Git repository is up to date.
goto :done

:failed
echo.
echo Script finished with errors. Press any key to close.

:done
pause

endlocal
