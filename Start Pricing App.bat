@echo off
rem ---- Willbanks Pricing App launcher ----
rem Double-click this file to start the app and open it in your browser.
cd /d "%~dp0"

set "PY=C:\Users\sherie.brown\AppData\Local\Python\bin\python.exe"
if not exist "%PY%" set "PY=python"

rem free port 8000 first, in case a previous instance or stray server is still using it
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000 " ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1

echo Starting the Pricing App...
rem cmd /k keeps the server window open so any error stays visible instead of vanishing
start "Willbanks Pricing App" cmd /k ""%PY%" server.py 8000"

rem give the server a moment to come up, then open the browser
ping -n 5 127.0.0.1 >nul
start "" http://localhost:8000

echo.
echo The app is now running in the other window titled "Willbanks Pricing App".
echo Your browser should have opened to http://localhost:8000
echo.
echo To STOP the app, close that other window.
echo (You can close this window now.)
ping -n 6 127.0.0.1 >nul
