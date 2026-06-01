@echo off
chcp 65001 > nul

echo ============================================
echo   KT Settlement Tool  v2.0
echo ============================================
echo.

:: -- Python path detection (Anaconda first: has pandas pre-installed) --
set PYTHON=

if exist "%USERPROFILE%\anaconda3\python.exe"                       set "PYTHON=%USERPROFILE%\anaconda3\python.exe"
if exist "%USERPROFILE%\Anaconda3\python.exe"                       set "PYTHON=%USERPROFILE%\Anaconda3\python.exe"
if exist "%USERPROFILE%\miniconda3\python.exe"                      set "PYTHON=%USERPROFILE%\miniconda3\python.exe"
if exist "%USERPROFILE%\Miniconda3\python.exe"                      set "PYTHON=%USERPROFILE%\Miniconda3\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"      set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"      set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"      set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if exist "C:\Program Files\Python313\python.exe"                    set "PYTHON=C:\Program Files\Python313\python.exe"
if exist "C:\Program Files\Python312\python.exe"                    set "PYTHON=C:\Program Files\Python312\python.exe"
if exist "C:\Program Files\Python311\python.exe"                    set "PYTHON=C:\Program Files\Python311\python.exe"

if not defined PYTHON (
    where python >nul 2>&1
    if not errorlevel 1 set "PYTHON=python"
)

if not defined PYTHON (
    echo [ERROR] Python not found.
    echo         Install Python 3.11+ and retry.
    echo.
    pause
    exit /b 1
)

echo Python : %PYTHON%
echo.

:: -- Script path (same folder as this bat file) --
set "SCRIPT=%~dp0run_gui.py"

if not exist "%SCRIPT%" (
    echo [ERROR] run_gui.py not found:
    echo         %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo Starting...
echo.

"%PYTHON%" "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo [ERROR] Exited with error. Check log above.
    echo.
    pause
)
