@echo off
echo ===========================================
echo   FEEDBACK ANALYTICS FAST PORTAL
echo ===========================================
echo.
echo Checking for Flask...
python -c "import flask" 2>nul
if errorlevel 1 (
    echo Flask not found. Installing...
    pip install flask pandas openpyxl
)

echo.
echo Starting the local server...
start cmd /k "python app.py"

echo.
echo Opening the portal in your browser...
timeout /t 3 /nobreak >nul
start http://127.0.0.1:8000

echo.
echo Portal is running! Keep the server window open while using it.
echo You can close this window now.
pause
