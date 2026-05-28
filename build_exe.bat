@echo off
REM ─────────────────────────────────────────────────────────────
REM  LMS 입과명단 분반기 · 단일 실행파일(.exe) 빌드 스크립트
REM  실행 후 dist\LMS입과명단분반기.exe 가 생성됩니다.
REM ─────────────────────────────────────────────────────────────

where pyinstaller >nul 2>nul
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다. 설치합니다...
    pip install pyinstaller || goto :err
)

pip install -r requirements.txt || goto :err

pyinstaller --noconfirm --clean ^
    --onefile --windowed ^
    --name "LMS입과명단분반기" ^
    app.py || goto :err

echo.
echo ─────────────────────────────────────────────
echo  빌드 완료 ▸  dist\LMS입과명단분반기.exe
echo ─────────────────────────────────────────────
exit /b 0

:err
echo.
echo [!] 빌드 중 오류가 발생했습니다.
exit /b 1
