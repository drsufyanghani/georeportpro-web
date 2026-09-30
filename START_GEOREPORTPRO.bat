@echo off
cd /d "%~dp0"
title GeoReportPro Web Launcher

echo =============================================
echo   GeoReportPro Web - One Click Launcher
echo =============================================
echo.
echo Installing/checking required Python packages...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installation failed. Please check Python/internet access.
  pause
  exit /b 1
)

echo.
echo Starting GeoReportPro on http://127.0.0.1:8010
start "GeoReportPro Server" cmd /k "cd /d "%~dp0" && python -m uvicorn georeport_app:app --host 127.0.0.1 --port 8010"

timeout /t 3 /nobreak >nul
start "" http://127.0.0.1:8010

echo.
echo GeoReportPro should now be open in your browser.
echo Keep the 'GeoReportPro Server' window open while using the app.
echo.
pause
