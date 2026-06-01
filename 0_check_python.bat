@echo off
echo.
echo ================================================================
echo   Internal PC - Python Environment Info
echo ================================================================
echo.
python --version
echo.
python -c "import sys,platform; print(\"Python tag  :\", \"cp\"+str(sys.version_info.major)+str(sys.version_info.minor)); print(\"Platform    :\", platform.machine()); print(\"Full path   :\", sys.executable)"
echo.
echo ================================================================
echo   Copy the lines above and send to the person running
echo   1_download_wheels.bat on the internet PC.
echo ================================================================
echo.
pause
