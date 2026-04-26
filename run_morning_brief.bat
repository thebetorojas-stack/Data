@echo off
REM EM Morning Brief — runs the free-data Python tool and opens the result.
cd /d "%~dp0"
python em_morning_brief.py %*
if errorlevel 1 (
    echo.
    echo *** Run failed. If first run, install dependencies:
    echo     pip install pandas requests matplotlib yfinance rich
    echo.
    pause
    exit /b 1
)
echo.
pause
