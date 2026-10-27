@echo off
chcp 65001 >nul
title 스마트 학급편성 시스템 - 파이썬 EXE 빌더
echo ========================================================
echo  스마트 학급편성 시스템 (파이썬 독립 실행 EXE 빌더)
echo ========================================================
echo.
echo 1. 필요한 패키지(pyinstaller, openpyxl) 설치 확인 중...
pip install pyinstaller openpyxl
echo.
echo 2. 단독 실행 EXE 파일 생성 중...
pyinstaller --noconsole --onefile --name="스마트학급편성" --icon=app.ico class_assigner.py
echo.
echo ========================================================
echo  빌드 완료!
echo  생성된 파일: dist\스마트학급편성.exe
echo ========================================================
pause
