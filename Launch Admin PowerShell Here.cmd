@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0open_admin_powershell_here.ps1"
endlocal
