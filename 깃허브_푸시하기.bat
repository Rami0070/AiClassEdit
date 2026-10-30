@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================================
echo  GitHub (Rami0070/AiClassEdit) 푸시 도구
echo ========================================================
echo.

set "GIT_CMD="

:: 1. PATH 내 git 확인
where git >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "GIT_CMD=git"
)

:: 2. LocalAppData MinGit 확인
if not defined GIT_CMD (
    if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" (
        set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
        set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"
    )
)

:: 3. Program Files Git 확인
if not defined GIT_CMD (
    if exist "C:\Program Files\Git\cmd\git.exe" (
        set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
        set "PATH=C:\Program Files\Git\cmd;%PATH%"
    )
)

:: 4. E:\Git 확인
if not defined GIT_CMD (
    if exist "E:\Git\cmd\git.exe" (
        set "GIT_CMD=E:\Git\cmd\git.exe"
        set "PATH=E:\Git\cmd;%PATH%"
    )
)

if not defined GIT_CMD (
    echo [오류] Git 실행 파일을 찾을 수 없습니다.
    echo Git이 설치되어 있는지 확인해주세요.
    echo.
    pause
    exit /b 1
)

echo [1/3] Git 실행 경로 확인: %GIT_CMD%
echo.

:: 변경 사항이 있는지 확인
"%GIT_CMD%" status --porcelain > "%TEMP%\git_status.tmp" 2>nul
for %%A in ("%TEMP%\git_status.tmp") do if %%~zA GTR 0 (
    echo [2/3] 변경된 파일들을 커밋에 추가합니다...
    "%GIT_CMD%" add .
    "%GIT_CMD%" commit -m "chore: update project files"
) else (
    echo [2/3] 작업 트리가 깨끗합니다. (커밋할 새 변경 없음)
)
if exist "%TEMP%\git_status.tmp" del "%TEMP%\git_status.tmp"

echo.
echo [3/3] GitHub 원격 저장소(origin/main)로 푸시를 진행합니다...
echo ※ 브라우저 인증 또는 로그인 창이 뜨면 승인해주세요.
echo.

"%GIT_CMD%" push -u origin main

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [오류] 푸시 중 문제가 발생했습니다.
    echo 위 에러 메시지를 확인해주세요.
) else (
    echo.
    echo ========================================================
    echo  성공적으로 GitHub에 푸시되었습니다!
    echo ========================================================
)

echo.
pause
