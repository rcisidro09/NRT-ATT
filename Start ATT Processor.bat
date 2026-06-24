@echo off
title AT&T Billing Data Processor
echo.
echo  =========================================
echo   AT^&T Billing Data Processor
echo  =========================================
echo.
echo  Starting server...
cd /d "%~dp0"
start "" "http://localhost:3000"
node server.js
pause
