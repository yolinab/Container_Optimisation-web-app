"""api.py — FastAPI web service for the Container Packing Optimizer.

Endpoints:
  GET  /         → serves the browser UI
  GET  /health   → health check
  POST /optimize → accepts Excel upload, runs pipeline, returns report.xlsx
"""

import sys
import shutil
import tempfile
from pathlib import Path

# Make app/ importable without installing as a package
sys.path.insert(0, str(Path(__file__).parent / "app"))


from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from pipeline import run_pipeline

app = FastAPI(title="Container Packing Optimizer")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    return (Path(__file__).parent / "static" / "index.html").read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/optimize")
async def optimize(
    file: UploadFile = File(...),
    count_col: str = Form(""),
):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are accepted.")

    col_override = count_col.strip() or None

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        excel_path = tmp_path / file.filename

        with excel_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)

        try:
            result = run_pipeline(
                excel_path=excel_path,
                out_dir=tmp_path / "outputs",
                count_col_override=col_override,
            )
        except RuntimeError as exc:
            raise HTTPException(status_code=422, detail=str(exc))
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Unexpected error: {exc}")

        report_bytes = result["report_path"].read_bytes()

    return Response(
        content=report_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=report.xlsx"},
    )
