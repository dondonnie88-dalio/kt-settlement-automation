@echo off
chcp 65001 > nul

echo ============================================
echo   Tax Verification Tool  (run_tax_verify.py)
echo ============================================
echo.

set PYTHON=

if exist "%USERPROFILE%\anaconda3\python.exe"                  set "PYTHON=%USERPROFILE%\anaconda3\python.exe"
if exist "%USERPROFILE%\Anaconda3\python.exe"                  set "PYTHON=%USERPROFILE%\Anaconda3\python.exe"
if exist "%USERPROFILE%\miniconda3\python.exe"                 set "PYTHON=%USERPROFILE%\miniconda3\python.exe"
if exist "%USERPROFILE%\Miniconda3\python.exe"                 set "PYTHON=%USERPROFILE%\Miniconda3\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if exist "C:\Program Files\Python313\python.exe"               set "PYTHON=C:\Program Files\Python313\python.exe"
if exist "C:\Program Files\Python312\python.exe"               set "PYTHON=C:\Program Files\Python312\python.exe"
if exist "C:\Program Files\Python311\python.exe"               set "PYTHON=C:\Program Files\Python311\python.exe"

if not defined PYTHON (
    where python >nul 2>&1
    if not errorlevel 1 set "PYTHON=python"
)

if not defined PYTHON (
    echo [ERROR] Python not found. Install Python 3.11+ and retry.
    echo.
    pause
    exit /b 1
)

set "SCRIPT=%~dp0tax_verify_gui.py"

if not exist "%SCRIPT%" (
    echo [ERROR] tax_verify_gui.py not found: %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo Python : %PYTHON%
echo Script : %SCRIPT%
echo.

"%PYTHON%" "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo [ERROR] Script exited with an error.
    echo.
    pause
)
