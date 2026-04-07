@echo off
setlocal
cd /d "%~dp0"

echo ======================================
echo Build Master Budget Automation Tool EXE
echo ======================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PY_CMD=python"
    ) else (
        echo Python was not found on this computer.
        echo Please install Python 3.11 or later, then run this file again.
        echo.
        pause
        exit /b 1
    )
)

echo Installing build tools...
call %PY_CMD% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo Failed to install build tools.
    echo.
    pause
    exit /b 1
)

echo.
echo Building EXE...
call %PY_CMD% -m PyInstaller --noconfirm --clean "Master Budget Automation Tool v1.0.2.spec"
if errorlevel 1 (
    echo.
    echo EXE build failed.
    echo.
    pause
    exit /b 1
)

echo.
echo Build complete.
echo EXE location: dist\Master Budget Automation Tool v1.0.2.exe
pause
endlocal
