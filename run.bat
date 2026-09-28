@echo off
REM Exhibition Label Studio - set up profiles and settings (Windows)
chcp 65001 >nul
cd /d "%~dp0"

set "PY=python"
where python >nul 2>nul || set "PY=py"
where %PY% >nul 2>nul
if errorlevel 1 (
    echo Python 3 is not installed. Get it from https://www.python.org/downloads/
    echo During install, tick "Add python.exe to PATH", then run this again.
    pause
    exit /b 1
)

if not exist "venv\Scripts\activate.bat" (
    echo [SETUP] First run - creating a private Python environment...
    %PY% -m venv venv
    if errorlevel 1 ( echo Could not create the environment. & pause & exit /b 1 )
)
call venv\Scripts\activate.bat

python -c "import openpyxl, reportlab" >nul 2>nul
if errorlevel 1 (
    echo [SETUP] Installing required packages ^(one time, needs internet^)...
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet openpyxl reportlab
    if errorlevel 1 ( echo Package install failed - check your internet connection. & pause & exit /b 1 )
)

python Generate_Labels.py configure %*
