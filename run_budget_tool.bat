@echo off
setlocal
cd /d "%~dp0"

echo ======================================
echo Master Budget Automation Tool v1.0.2
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

echo Checking required packages...
call %PY_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Failed to install required packages.
    echo Please check your Python or internet connection, then try again.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting the app...
call %PY_CMD% app.py
set "APP_EXIT=%errorlevel%"

if not "%APP_EXIT%"=="0" (
    echo.
    echo The app closed with an error code %APP_EXIT%.
    echo Please take a screenshot of this window and send it to support.
    echo.
    pause
    exit /b %APP_EXIT%
)

endlocal
