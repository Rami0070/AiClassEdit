@echo off
cd /d "%~dp0"

echo ========================================================
echo  GitHub (Rami0070/AiClassEdit) Push
echo ========================================================
echo.

set "GIT_CMD="

if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" (
    set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
    set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"
) else if exist "C:\Program Files\Git\cmd\git.exe" (
    set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    set "PATH=C:\Program Files\Git\cmd;%PATH%"
) else (
    set "GIT_CMD=git"
)

echo [1/3] Git executable: %GIT_CMD%
echo.

echo [2/3] Checking changed files...
"%GIT_CMD%" add .
"%GIT_CMD%" commit -m "chore: update project files" >nul 2>&1

echo.
echo [3/3] Pushing to GitHub (origin/main)...
echo Please approve authentication in your browser if prompted.
echo.

"%GIT_CMD%" push -u origin main

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Push failed. Please check the error message above.
) else (
    echo.
    echo ========================================================
    echo  Push completed successfully!
    echo ========================================================
)

echo.
pause
