@echo off
title Auto Note to Word App (Windows)
echo ==============================================
echo Checking Python installation...
echo ==============================================

:: Check if 'py' launcher is available
where py >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Python Launcher (py) not found.
    echo Please install Python from https://www.python.org/
    echo and make sure to check "Add Python to PATH" or "Install Python Launcher".
    echo.
    pause
    exit /b
)

echo ==============================================
echo Installing Python requirements...
echo ==============================================
py -m pip install -r requirements.txt

echo.
echo ==============================================
echo Finding your IP Address for Tablet Access...
echo ==============================================
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr "IPv4" ^| findstr /v "127.0.0.1"') do (
    set IP=%%a
)
set IP=%IP: =%
echo Your PC's Network IP: %IP%
echo.
echo To use this on your Galaxy Tab or Phone:
echo 1. Make sure the Tablet and this PC are on the same Wi-Fi.
echo 2. Open the Browser on your Tablet.
echo 3. Enter this address: http://%IP%:8501
echo ==============================================
echo.

echo Starting the Streamlit App...
echo ==============================================
echo A browser window should open automatically on this PC.
echo.
py -m streamlit run app.py --server.address 0.0.0.0
pause
