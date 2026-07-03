@echo off
title CPSM Demand Forecast

echo.
echo ================================================================
echo    CPSM Demand Forecast - Automation System
echo ================================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install from https://www.python.org
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0"

echo [INFO] Starting demand_forecast.py ...
echo.

python demand_forecast.py

echo.
if errorlevel 1 (
    echo [FAILED] Check forecast_log.txt for details.
) else (
    echo [DONE] Output file: demand_forecast_YYYYMM.xlsx
)

echo.
pause
