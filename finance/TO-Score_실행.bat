@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo  TO-Score 회비 대시보드 실행 중...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [오류] Python이 설치되어 있지 않습니다.
    pause
    exit /b 1
)

pip show pandas >nul 2>&1
if errorlevel 1 (
    echo  필요한 라이브러리 설치 중...
    pip install pandas openpyxl -q
)

if not exist "data\TO-Score.xlsx" (
    echo  엑셀 파일 최초 생성 중...
    python src\create_excel.py
    echo.
)

start "" cmd /c "timeout /t 2 >nul && start http://localhost:8000"
python src\dashboard.py

pause
