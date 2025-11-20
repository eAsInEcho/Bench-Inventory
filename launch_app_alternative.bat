@echo off
echo Starting IT Bench Inventory Manager...

REM Try using Python from PATH first
where python >nul 2>nul
if %errorlevel% equ 0 (
    echo Using Python from PATH...
    python main.py
) else (
    REM Try common Python installation locations
    echo Python not found in PATH, trying common installation locations...
    
    if exist "C:\Program Files\Python312\python.exe" (
        echo Found Python 3.12...
        "C:\Program Files\Python312\python.exe" main.py
    ) else if exist "C:\Program Files\Python311\python.exe" (
        echo Found Python 3.11...
        "C:\Program Files\Python311\python.exe" main.py
    ) else if exist "C:\Program Files\Python310\python.exe" (
        echo Found Python 3.10...
        "C:\Program Files\Python310\python.exe" main.py
    ) else if exist "C:\Program Files\Python39\python.exe" (
        echo Found Python 3.9...
        "C:\Program Files\Python39\python.exe" main.py
    ) else if exist "C:\Python39\python.exe" (
        echo Found Python 3.9...
        "C:\Python39\python.exe" main.py
    ) else if exist "C:\Program Files (x86)\Python39\python.exe" (
        echo Found Python 3.9...
        "C:\Program Files (x86)\Python39\python.exe" main.py
    ) else (
        echo Python installation not found in common locations.
        echo Please edit this batch file to specify the correct path to python.exe
        pause
        exit /b 1
    )
)

REM Check if application exited with an error
if %errorlevel% neq 0 (
    echo.
    echo Application encountered an error (Error Code: %errorlevel%)
    echo Please check the console output above for details.
    pause
)
