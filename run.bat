@echo off
REM Double-click this file to start the Order Status Cleaner on this PC.
cd /d "%~dp0"
if not exist venv (
    echo First run: setting things up, please wait...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)
streamlit run app.py
pause
