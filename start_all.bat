@echo off
title U1P + QuickBooks Integration
color 0A
echo.
echo  =====================================================
echo   U1P Order Conversion Tool + QB Integration Server
echo  =====================================================
echo.

:: Check Python
py --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found. Install from https://www.python.org
    pause
    exit /b 1
)

:: Install / update dependencies
echo  Checking dependencies...
py -m pip install -r requirements.txt -q --disable-pip-version-check

:: Build product catalog if missing
if not exist products.json (
    echo  Building product catalog...
    py build_catalog.py
)

echo.
echo  Starting QB SOAP server on port 8443...
start "QB SOAP Server" cmd /k "py qb_server.py"

:: Brief pause so the SOAP server initialises before Flask
timeout /t 2 /nobreak >nul

echo  Starting Flask app on port 5000...
echo.
echo  Flask UI  : http://localhost:5000
echo  QB SOAP   : http://localhost:8443/qbwc
echo.
echo  Remember: QB Web Connector must also be running on this machine.
echo  Press Ctrl+C in the Flask window to stop.
echo.
py app.py

pause
