@echo off
title Document Formatter App
color 0A

echo ------------------------------------------
echo  📄 Starting Document Formatter App...
echo ------------------------------------------

:: Step 1: Move to the current folder
cd /d "%~dp0"

:: Step 2: Check if Python is installed
where python >nul 2>nul
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH.
    echo 🔗 Please install Python from https://www.python.org/downloads
    pause
    exit /b
)

:: Step 3: Install required packages (streamlit, python-docx)
echo 🔄 Installing required Python packages...
pip install --quiet streamlit python-docx

:: Step 4: Run the Streamlit app
echo ✅ Launching the app in your browser...
python -m streamlit run app.py

echo ------------------------------------------
echo 🟢 If the browser didn’t open, visit: http://localhost:8501
echo ------------------------------------------
pause
