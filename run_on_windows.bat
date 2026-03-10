@echo off
echo ==============================================
echo 자동 발주서 프로그램 실행 (Windows 환경)
echo ==============================================
echo.

:: 1. 의존성 설치 점검
echo 의존성 패키지를 설치/업데이트합니다...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [경고] requirements.txt 패키지 설치 중 오류가 발생했습니다.
)
pip install pywin32 pandas pyxlsb
echo.

:: 2. 프로그램 실행
echo 프로그램을 시작합니다...
python main.py
if %errorlevel% neq 0 (
    echo [오류] 프로그램 실행 중 문제가 발생했습니다. (Python이 설치되어 있는지 확인해주세요)
)
pause
