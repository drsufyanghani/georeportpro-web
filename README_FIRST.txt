GEOREPORTPRO WEB - CLEAN BUILD
==============================

1. Open Command Prompt in this extracted folder.
2. Install dependencies once:
   python -m pip install -r requirements.txt

3. Start GeoReportPro locally:
   python -m uvicorn georeport_app:app --reload --port 8010

4. Open Chrome:
   http://127.0.0.1:8010

5. Confirm API title:
   http://127.0.0.1:8010/docs
   It should say GeoReportPro Web.

Render settings:
Build command:
   pip install -r requirements.txt

Start command:
   uvicorn georeport_app:app --host 0.0.0.0 --port $PORT

Health check:
   /health
