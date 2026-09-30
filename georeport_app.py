from __future__ import annotations

import base64
import io
import os
import secrets
import uuid
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from engine import (
    ProjectInfo,
    analyze_cbr_from_key_loads,
    analyze_compaction,
    analyze_spt,
    auto_populate_from_workbook,
    build_word_report,
    calculate_atterberg,
    calculate_terzaghi,
    classify_soil,
    df_preview,
    new_state,
    plot_png,
    run_qc_checks,
    state_payload,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")

app = FastAPI(
    title="GeoReportPro Web",
    version="1.0.0",
    description="IS-referenced geotechnical investigation reporting and calculation web application.",
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Demo/session store. Each browser gets a cookie and its own in-memory state.
# For a production multi-user system, replace this with PostgreSQL/object storage.
SESSIONS: Dict[str, Dict[str, Any]] = {}
MAX_SESSIONS = 100

APP_PASSWORD = os.getenv("APP_PASSWORD", "").strip()
AUTH_USER = os.getenv("APP_USER", "georeport").strip() or "georeport"


def _authorized(request: Request) -> bool:
    if not APP_PASSWORD:
        return True
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Basic "):
        return False
    try:
        decoded = base64.b64decode(auth[6:]).decode("utf-8")
        username, password = decoded.split(":", 1)
    except Exception:
        return False
    return secrets.compare_digest(username, AUTH_USER) and secrets.compare_digest(password, APP_PASSWORD)


@app.middleware("http")
async def session_and_optional_auth(request: Request, call_next):
    # Health route is left open for Render health checks.
    if request.url.path != "/health" and not _authorized(request):
        return Response(
            "Authentication required",
            status_code=401,
            headers={"WWW-Authenticate": 'Basic realm="GeoReportPro"'},
        )

    sid = request.cookies.get("geo_sid") or uuid.uuid4().hex
    request.state.sid = sid
    if sid not in SESSIONS:
        # Keep the demo memory bounded.
        if len(SESSIONS) >= MAX_SESSIONS:
            oldest = next(iter(SESSIONS))
            SESSIONS.pop(oldest, None)
        SESSIONS[sid] = new_state()

    response = await call_next(request)
    if request.cookies.get("geo_sid") != sid:
        response.set_cookie(
            "geo_sid",
            sid,
            httponly=True,
            samesite="lax",
            secure=request.url.scheme == "https",
            max_age=60 * 60 * 12,
        )
    return response


def get_state(request: Request) -> Dict[str, Any]:
    return SESSIONS[request.state.sid]


class ProjectPayload(BaseModel):
    project_name: str = ""
    client_name: str = ""
    location: str = ""
    report_number: str = ""
    prepared_by: str = ""
    checked_by: str = ""
    approved_by: str = ""
    report_date: str = ""
    remarks: str = ""


class ClassificationPayload(BaseModel):
    ll: float = Field(ge=0)
    pl: float = Field(ge=0)
    water_content: Optional[float] = Field(default=None, ge=0)
    gravel: float = Field(ge=0)
    sand: float = Field(ge=0)
    fines: float = Field(ge=0)


class CompactionPointPayload(BaseModel):
    water_content: float
    wet_density: float


class CompactionPayload(BaseModel):
    points: List[CompactionPointPayload]


class CBRPayload(BaseModel):
    load_2_5: float = Field(ge=0)
    load_5_0: float = Field(ge=0)
    adopt_5_if_higher_confirmed: bool = False


class SPTPayload(BaseModel):
    b1: int = Field(ge=0)
    b2: int = Field(ge=0)
    b3: int = Field(ge=0)
    correction_factor: float = Field(ge=0)
    saturated_fine_sand: bool = False


class BearingPayload(BaseModel):
    foundation_type: str
    width: float = Field(gt=0)
    depth: float = Field(ge=0)
    cohesion: float = Field(ge=0)
    unit_weight: float = Field(gt=0)
    phi: float = Field(ge=0, le=50)
    fos: float = Field(gt=0)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/")
def root():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.get("/api/state")
def api_state(request: Request):
    return state_payload(get_state(request))


@app.post("/api/reset")
def api_reset(request: Request):
    SESSIONS[request.state.sid] = new_state()
    return state_payload(SESSIONS[request.state.sid])


@app.post("/api/project")
def api_project(payload: ProjectPayload, request: Request):
    state = get_state(request)
    state["project_info"] = ProjectInfo(**payload.model_dump())
    return state_payload(state)


@app.post("/api/classification")
def api_classification(payload: ClassificationPayload, request: Request):
    state = get_state(request)
    try:
        if payload.ll < payload.pl:
            raise ValueError("Liquid limit cannot be less than plastic limit.")
        att = calculate_atterberg(payload.ll, payload.pl, payload.water_content)
        cls = classify_soil(payload.gravel, payload.sand, payload.fines, payload.ll, att.plasticity_index)
        state["atterberg_result"] = att
        state["classification_result"] = cls
        state["inputs"]["classification"] = payload.model_dump()
        return state_payload(state)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/compaction")
def api_compaction(payload: CompactionPayload, request: Request):
    state = get_state(request)
    try:
        points = payload.points
        result = analyze_compaction(
            [p.water_content for p in points],
            [p.wet_density for p in points],
        )
        state["compaction_result"] = result
        state["inputs"]["compaction"] = {
            "points": [p.model_dump() for p in points]
        }
        return state_payload(state)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/cbr")
def api_cbr(payload: CBRPayload, request: Request):
    state = get_state(request)
    try:
        result = analyze_cbr_from_key_loads(
            payload.load_2_5,
            payload.load_5_0,
            payload.adopt_5_if_higher_confirmed,
        )
        state["cbr_result"] = result
        state["inputs"]["cbr"] = payload.model_dump()
        return state_payload(state)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/spt")
def api_spt(payload: SPTPayload, request: Request):
    state = get_state(request)
    try:
        result = analyze_spt(
            payload.b1,
            payload.b2,
            payload.b3,
            payload.correction_factor,
            payload.saturated_fine_sand,
        )
        state["spt_result"] = result
        state["inputs"]["spt"] = payload.model_dump()
        return state_payload(state)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/bearing")
def api_bearing(payload: BearingPayload, request: Request):
    state = get_state(request)
    try:
        result = calculate_terzaghi(
            payload.foundation_type,
            payload.width,
            payload.depth,
            payload.cohesion,
            payload.unit_weight,
            payload.phi,
            payload.fos,
        )
        state["bearing_result"] = result
        state["inputs"]["bearing"] = payload.model_dump()
        return state_payload(state)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/upload")
async def api_upload(request: Request, file: UploadFile = File(...)):
    state = get_state(request)
    name = file.filename or "workbook.xlsx"
    ext = os.path.splitext(name)[1].lower()
    if ext not in {".xlsx", ".xls"}:
        raise HTTPException(status_code=400, detail="Please upload an .xlsx or .xls workbook.")

    content = await file.read()
    if len(content) > 20 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="Workbook is larger than the 20 MB demo limit.")

    try:
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, engine=engine)
        state["data"] = sheets
        state["excel_name"] = name
        auto_populate_from_workbook(state)
        preview = {sheet: df_preview(df, 30) for sheet, df in sheets.items()}
        return {
            "message": "Workbook imported successfully.",
            "preview": preview,
            "state": state_payload(state),
        }
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read workbook: {exc}")


@app.post("/api/qc")
def api_qc(request: Request):
    state = get_state(request)
    run_qc_checks(state)
    return state_payload(state)


@app.get("/api/plot/{kind}")
def api_plot(
    kind: str,
    request: Request,
    borehole: Optional[str] = None,
    sample: Optional[str] = None,
    spt_value: str = "Raw N",
):
    allowed = {"borehole", "psd", "plasticity", "compaction", "cbr", "spt"}
    if kind not in allowed:
        raise HTTPException(status_code=404, detail="Unknown plot type")
    try:
        png = plot_png(get_state(request), kind, borehole=borehole, sample=sample, spt_value=spt_value)
        return Response(content=png, media_type="image/png", headers={"Cache-Control": "no-store"})
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/report")
def api_report(request: Request):
    state = get_state(request)
    try:
        data = build_word_report(state)
        p = state["project_info"]
        safe_name = "".join(c for c in (p.project_name or "GeoReportPro") if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
        filename = f"{safe_name or 'GeoReportPro'}_GIR.docx"
        headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not generate report: {exc}")


@app.get("/api/preview/{sheet}")
def api_preview(sheet: str, request: Request):
    state = get_state(request)
    df = state["data"].get(sheet)
    if df is None:
        raise HTTPException(status_code=404, detail="Sheet not found")
    return df_preview(df, 100)
