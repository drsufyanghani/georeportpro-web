# GeoReportPro Web

FastAPI web conversion of the GeoReportPro India v0.6 desktop proof-of-concept.

## Local run

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m uvicorn georeport_app:app --reload
```

Open: http://127.0.0.1:8000

## Optional demo password

Set environment variables before starting:

Windows CMD:

```bat
set APP_USER=georeport
set APP_PASSWORD=choose-a-password
python -m uvicorn georeport_app:app --reload
```

The browser will show an HTTP Basic login prompt.

## Render

Build command:

```text
pip install -r requirements.txt
```

Start command:

```text
uvicorn georeport_app:app --host 0.0.0.0 --port $PORT
```

Health check path:

```text
/health
```

## Important demo limitation

The application keeps each browser session in process memory. Uploaded workbooks and calculations are lost when the Render service restarts, redeploys, or a free service spins down. Use a database/object storage for production persistence.
