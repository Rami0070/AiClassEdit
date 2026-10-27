@echo off
chcp 65001 >nul
cd /d "C:\AI개발\학급편성"
echo ========================================================
echo  GitHub (Rami0070/AiClassEdit) 첫 푸시 진행
echo ========================================================
echo.
echo 브라우저 로그인 창이 뜨면 GitHub 계정으로 로그인(승인)해주세요.
echo.
"C:\Program Files\Git\cmd\git.exe" push -u origin main
echo.
echo ========================================================
echo 작업이 완료되었습니다.
echo ========================================================
pause
