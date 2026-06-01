@echo off
title Offline Installer (Internal PC)

echo.
echo ================================================================
echo   [STEP 2] Offline Installation - Internal PC
echo ================================================================
echo.

cd /d "%~dp0"

if not exist wheels (
    echo [ERROR] wheels\ folder not found.
    pause & exit /b 1
)

:: --- nested folder auto-fix ---
dir /b wheels\*.whl >nul 2>&1
if errorlevel 1 (
    if exist wheels\wheels\*.whl (
        echo [FIX] Moving .whl files up ...
        move wheels\wheels\*.whl wheels\ >nul
    ) else (
        echo [ERROR] No .whl files found.
        pause & exit /b 1
    )
)

echo [INFO] wheels\ contents:
dir /b wheels\*.whl
echo.

echo [INFO] Python version:
python --version
echo.

echo [INFO] Installing all packages and dependencies ...
echo.

:: --- install every .whl in the folder at once ---
python -m pip install --no-index --find-links ./wheels ^
    pandas numpy statsmodels scikit-learn openpyxl xlsxwriter ^
    python-dateutil six patsy scipy joblib threadpoolctl ^
    packaging tzdata et-xmlfile

echo.
echo [INFO] Verifying ...
python verify_packages.py

echo.
if errorlevel 1 (
    echo [FAILED] Check messages above.
) else (
    echo [DONE] All packages OK. Run ????_??.bat
)

echo.
pause
