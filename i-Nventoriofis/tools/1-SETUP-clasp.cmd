@echo off
REM Double-click me. Wraps the .ps1 so Windows execution policy cannot
REM block it, and keeps the window open so the result stays readable.
title i-Nventori Ofis - Setup clasp (sekali sahaja)
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup-clasp.ps1"
echo.
pause
