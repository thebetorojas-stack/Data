@echo off
REM Launches the Streamlit dashboard at http://localhost:8501
cd /d "%~dp0"
streamlit run src\dashboard\app.py
