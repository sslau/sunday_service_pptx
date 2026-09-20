@echo off
REM Windows quick launch (installs if needed, then runs the app).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" run
pause
