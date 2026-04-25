@echo off
REM One-click refresh — pulls Bloomberg + Haver deltas, regenerates Chile.xlsx and Argentina.xlsx.
REM Double-click this file from Explorer.

cd /d "%~dp0"
echo.
echo === EM Macro & Credit refresh ===
echo.

python -m scripts.refresh_all
if errorlevel 1 (
    echo.
    echo *** Refresh failed. Common causes:
    echo     - Bloomberg Terminal not running or not logged in
    echo     - Haver DLX not on PYTHONPATH
    echo     - Python venv not activated
    echo.
    pause
    exit /b 1
)

echo.
echo Workbooks ready in .\outputs\
echo.
pause
