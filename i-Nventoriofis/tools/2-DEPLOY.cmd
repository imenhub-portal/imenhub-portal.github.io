@echo off
REM Double-click me to deploy Code.gs to Apps Script.
title i-Nventori Ofis - Deploy Code.gs
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1"
echo.
pause
