from pathlib import Path
import shutil
import tempfile
from typing import Dict

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from fastapi.staticfiles import StaticFiles

from app.services.reconciliation import run_reconciliation

BASE_DIR = Path(__file__).resolve().parents[2]
RESULT_DIR = BASE_DIR / "results"
RESULT_INDEX: Dict[str, Dict[str, str]] = {}
STATIC_DIR = BASE_DIR / "app" / "static"

app = FastAPI(title="119 vs Mgmt C Reconciliation API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class HealthResponse(BaseModel):
    status: str

@app.get("/api/health", response_model=HealthResponse)
def health() -> HealthResponse:
    return HealthResponse(status="ok")

@app.post("/api/reconcile")
async def reconcile(raw119: UploadFile = File(...), mgmtc: UploadFile = File(...)):
    if not raw119.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="Raw 119 report must be an Excel file.")
    if not mgmtc.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="Raw Mgmt C report must be an Excel file.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        raw119_path = tmp_dir / raw119.filename
        mgmt_path = tmp_dir / mgmtc.filename
        with raw119_path.open("wb") as f:
            shutil.copyfileobj(raw119.file, f)
        with mgmt_path.open("wb") as f:
            shutil.copyfileobj(mgmtc.file, f)
        try:
            result = run_reconciliation(raw119_path, mgmt_path, raw119.filename, mgmtc.filename, RESULT_DIR)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Reconciliation failed: {exc}") from exc

    RESULT_INDEX[result.result_id] = {"path": result.output_path, "filename": result.filename}
    return {
        "resultId": result.result_id,
        "filename": result.filename,
        "latestDateLabel": result.latest_date_label,
        "reportingPeriod": result.reporting_period,
        "summary": result.summary,
        "metrics": result.metrics,
        "tabs": result.tabs,
    }

@app.get("/api/download/{result_id}")
def download(result_id: str):
    item = RESULT_INDEX.get(result_id)
    if not item:
        matches = list(RESULT_DIR.glob(f"{result_id}_*.xlsx"))
        if not matches:
            raise HTTPException(status_code=404, detail="Result not found. Run the reconciliation again.")
        path = matches[0]
        filename = path.name.split("_", 1)[1]
    else:
        path = Path(item["path"])
        filename = item["filename"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="Export file no longer exists. Run the reconciliation again.")
    return FileResponse(path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Serve the built React frontend from the same public web service.
# This prevents production "Failed to fetch" errors caused by frontend/backend URL mismatch.
if STATIC_DIR.exists():
    assets_dir = STATIC_DIR / "assets"
    if assets_dir.exists():
        app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")

    @app.get("/{full_path:path}", response_class=HTMLResponse)
    def serve_frontend(full_path: str):
        index_path = STATIC_DIR / "index.html"
        if index_path.exists():
            return HTMLResponse(index_path.read_text(encoding="utf-8"))
        raise HTTPException(status_code=404, detail="Frontend build not found.")
