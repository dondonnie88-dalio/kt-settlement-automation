@echo off
title 통신/일반 분류 진단
chcp 65001 >nul

echo.
echo ================================================================
echo    통신/일반 분류 미분류 원인 진단
echo ================================================================
echo.

cd /d "%~dp0"

python diagnose_prod_type.py

echo.
pause
