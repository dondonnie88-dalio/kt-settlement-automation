@echo off
title Wheel Downloader (Internet PC) - cp314

echo.
echo ================================================================
echo   [STEP 1] Internet PC - Download wheels for cp314 / win_amd64
echo ================================================================
echo.

cd /d "%~dp0"
if not exist wheels mkdir wheels

set PY_VER=314
set PYPI=https://pypi.org/simple
set PKGS=pandas numpy statsmodels scikit-learn openpyxl xlsxwriter

echo [INFO] Target : cp%PY_VER%, win_amd64
echo [INFO] Dest   : .\wheels\
echo.

:: --- Package-by-package download (cp314 -> cp313 -> any fallback) ---
for %%P in (%PKGS%) do (
    echo [%%P] Trying cp%PY_VER% binary ...
    python -m pip download %%P -d ./wheels ^
        --platform win_amd64 ^
        --python-version %PY_VER% ^
        --implementation cp ^
        --only-binary=:all: ^
        --index-url %PYPI% >nul 2>&1
    if errorlevel 1 (
        echo [%%P] No cp%PY_VER% wheel. Trying cp313 ...
        python -m pip download %%P -d ./wheels ^
            --platform win_amd64 ^
            --python-version 313 ^
            --implementation cp ^
            --only-binary=:all: ^
            --index-url %PYPI% >nul 2>&1
        if errorlevel 1 (
            echo [%%P] No binary wheel. Downloading source + deps ...
            python -m pip download %%P -d ./wheels ^
                --index-url %PYPI%
            if errorlevel 1 (
                echo [%%P] FAILED - manual download required.
            ) else (
                echo [%%P] OK ^(source^)
            )
        ) else (
            echo [%%P] OK ^(cp313 binary^)
        )
    ) else (
        echo [%%P] OK ^(cp%PY_VER% binary^)
    )
    echo.
)

echo.
echo ================================================================
echo   Downloaded files in .\wheels\
echo ================================================================
dir /b wheels\
echo.
echo   Copy the entire [wheels] folder to the internal PC,
echo   then run 2_install_offline.bat
echo ================================================================
echo.
pause
