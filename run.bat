@echo off
title U1P Order Conversion Tool
color 0A
echo.
echo  =========================================
echo   U1P Order Conversion Tool - Ultra1Plus
echo  =========================================
echo.

:: Check Python
py --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python is not installed or not in PATH.
    echo  Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install dependencies
echo  Checking dependencies...
py -m pip install -r requirements.txt -q --disable-pip-version-check

:: Build product catalog if it doesn't exist
if not exist products.json (
    echo.
    echo  Building product catalog from Excel...
    py build_catalog.py
    if errorlevel 1 (
        echo  WARNING: Could not build catalog. Make sure the Excel file is accessible.
    )
) else (
    echo  Product catalog: OK
)

echo.
echo  Starting server...
echo  Open your browser and go to: http://localhost:5000
echo.
echo  Press Ctrl+C to stop the server.
echo.
py app.py

pause
