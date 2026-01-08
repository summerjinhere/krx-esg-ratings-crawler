@echo off
chcp 65001 > nul
echo ========================================
echo   KRX ESG 크롤러 실행
echo ========================================
echo.

REM Python 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    echo https://www.python.org/downloads/ 에서 설치하세요.
    pause
    exit /b 1
)

REM 패키지 설치 확인
echo [1/3] 패키지 확인 중...
pip show playwright > nul 2>&1
if errorlevel 1 (
    echo      패키지 설치 중...
    pip install playwright pandas openpyxl tqdm
    playwright install chromium
)

echo [2/3] 준비 완료
echo.

REM 실행 모드 선택
echo 실행 모드를 선택하세요:
echo   1. 테스트 모드 (3개 기업)
echo   2. 소규모 테스트 (10개 기업)
echo   3. 전체 크롤링 (KCGS 파일 필요)
echo   4. 중단된 곳에서 재개
echo.
set /p mode="선택 (1-4): "

if "%mode%"=="1" (
    echo.
    echo [3/3] 테스트 모드 실행...
    python krx_esg_crawler.py
) else if "%mode%"=="2" (
    set /p inputfile="KCGS Excel 파일명: "
    echo.
    echo [3/3] 소규모 테스트 실행 (10개 기업)...
    python krx_esg_crawler.py --input "%inputfile%" --max 10
) else if "%mode%"=="3" (
    set /p inputfile="KCGS Excel 파일명: "
    echo.
    echo [3/3] 전체 크롤링 실행...
    echo      (약 6-12시간 소요, Ctrl+C로 중단 가능)
    python krx_esg_crawler.py --input "%inputfile%" --headless
) else if "%mode%"=="4" (
    set /p inputfile="KCGS Excel 파일명: "
    echo.
    echo [3/3] 체크포인트에서 재개...
    python krx_esg_crawler.py --input "%inputfile%" --headless --resume
) else (
    echo 잘못된 선택입니다.
)

echo.
echo ========================================
echo   완료!
echo ========================================
pause
