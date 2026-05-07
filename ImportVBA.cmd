@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0import_vba.ps1"
if %ERRORLEVEL% neq 0 pause
