@echo off
chcp 65001 >nul
echo ========================================
echo 영화 데이터 수집을 시작합니다...
echo Starting movie data collection...
echo ========================================
python "system\tmdb_fetch.py"
if errorlevel 1 (
    echo.
    echo [오류] 파이썬이 실행되지 않았습니다. 
    echo 파이썬이 설치되어 있고 '환경변수(PATH)'에 추가되어 있는지 확인해주세요.
)
echo.
pause
