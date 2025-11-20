@echo off
echo Starting IT Bench Inventory Manager...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo Application encountered an error (Error Code: %errorlevel%)
    echo Please check the console output above for details.
    pause
)
