@echo off

chcp 65001 >nul

cd /d "%~dp0"

set "VENV=%~dp0.venv"

set "PYTHON=%VENV%\Scripts\python.exe"

 

if not exist "%PYTHON%" (

    echo [1/3] Creating Python virtual environment...

    py -3 -m venv "%VENV%"

    if errorlevel 1 goto :venv_error

)

 

"%PYTHON%" -c "import nicegui, openpyxl" >nul 2>&1

if errorlevel 1 (

    echo [2/3] Installing required packages...

    "%PYTHON%" -m pip install --disable-pip-version-check -r requirements.txt

    if errorlevel 1 goto :pip_error

)

 

echo [3/3] Starting application...

echo Local URL: http://127.0.0.1:8080

 

"%PYTHON%" app.py

pause

exit /b 0

 

:venv_error

echo Failed to create virtual environment.

pause

exit /b 1

 

:pip_error

echo Failed to install required packages.

pause

exit /b 1