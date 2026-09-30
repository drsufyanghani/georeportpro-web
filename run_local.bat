@echo off
cd /d "%~dp0"
echo Starting GeoReportPro Web on http://127.0.0.1:8010
python -m uvicorn georeport_app:app --reload --port 8010
pause
