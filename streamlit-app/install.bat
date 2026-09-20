@echo off
REM Windows quick install (creates .venv and installs dependencies).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" install
if errorlevel 1 pause
