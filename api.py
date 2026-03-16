"""api.py — FastAPI web service for the Container Packing Optimizer.

Endpoints:
  GET  /            → serves the browser UI
  GET  /health      → health check
  POST /optimize    → accepts Excel upload, runs pipeline, returns JSON result
  POST /report-bug  → sends a bug report email
"""

import base64
import json
import os
import shutil
import sys
import tempfile
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

# Use the non-interactive Agg backend so matplotlib renders to PNG buffers
# without needing a display.  Must be set before pyplot is first imported.
import matplotlib
matplotlib.use("Agg")

# Make app/ importable without installing as a package
sys.path.insert(0, str(Path(__file__).parent / "app"))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from pipeline import run_pipeline
from utils.visualize_row_blocks import render_container_to_png_b64
from config import (
    CONTAINER_WIDTH_CM, CONTAINER_HEIGHT_CM, CONTAINER_DOOR_HEIGHT_CM, ROW_GAP_CM,
)

app = FastAPI(title="Container Packing Optimizer")
app.mount("/static", StaticFiles(directory="static"), name="static")


# ── Helpers ────────────────────────────────────────────────────────────────

def _issues_to_frontend(issues: List[Dict[str, Any]]) -> Dict[str, List[str]]:
    """Split validation issues into user-friendly warning/error string lists."""
    warnings = [i["message"] for i in issues if i["level"] == "WARNING"]
    errors   = [i["message"] for i in issues if i["level"] == "ERROR"]
    return {"warnings": warnings, "errors": errors}


# ── Routes ─────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def index():
    return (Path(__file__).parent / "static" / "index.html").read_text(encoding="utf-8")


@app.get("/faq", response_class=HTMLResponse)
def faq():
    return (Path(__file__).parent / "static" / "faq.html").read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/optimize")
async def optimize(
    file: UploadFile = File(...),
):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are accepted.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path   = Path(tmp)
        excel_path = tmp_path / file.filename

        with excel_path.open("wb") as fh:
            shutil.copyfileobj(file.file, fh)

        try:
            result = run_pipeline(
                excel_path=excel_path,
                out_dir=tmp_path / "outputs",
            )
        except RuntimeError as exc:
            raise HTTPException(status_code=422, detail=str(exc))
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Unexpected error: {exc}")

        report_b64        = base64.b64encode(result["report_path"].read_bytes()).decode()
        issue_summary     = _issues_to_frontend(result.get("validation_issues", []))
        containers        = result["containers"]
        recs_by_idx       = {r["container_index"]: r for r in result.get("recommendations", [])}
        L_cm              = int(os.environ.get("CONTAINER_LENGTH_CM", 1203))
        overall_decisions = result.get("overall_decisions", {})

        # Minimal layout data for the browser 2-D overview strip
        layout_data = [
            {
                "idx": c["container_index"],
                "rows": [
                    {
                        "bt":  r["block_type"],
                        "y":   r["y_start_cm"],
                        "len": r["length_cm"],
                        "h":   r["height_cm"],
                        "n":   r["pallet_count"],
                    }
                    for r in c.get("rows", [])
                ],
                "zones": [
                    {
                        "y":   z["y_start_cm"],
                        "len": z["length_cm"],
                        "n":   sum(p["quantity"] for p in z.get("placed", [])),
                    }
                    for z in c.get("box_zones", [])
                ],
                "decisions": c.get("decisions", {}),
            }
            for c in containers
        ]

        # 3-D matplotlib renders — one PNG per container
        container_images = []
        for c in containers:
            rec = recs_by_idx.get(c["container_index"])
            try:
                img_b64 = render_container_to_png_b64(
                    container=c,
                    W=CONTAINER_WIDTH_CM,
                    L=L_cm,
                    H=CONTAINER_HEIGHT_CM,
                    gap_cm=ROW_GAP_CM,
                    rec=rec,
                )
                container_images.append(img_b64)
            except Exception as exc:
                print(f"[warn] Could not render 3-D image for container {c['container_index']}: {exc}")
                container_images.append(None)

    return JSONResponse({
        "report_b64":        report_b64,
        "filename":          "packing_report.xlsx",
        "container_count":   len(containers),
        "total_pallets":     sum(c.get("loaded_value", 0) for c in containers),
        "warnings":          issue_summary["warnings"],
        "errors":            issue_summary["errors"],
        "layout_data":       layout_data,
        "container_length_cm": L_cm,
        "container_images":  container_images,
        "overall_decisions": overall_decisions,
    })


# ── Bug report ─────────────────────────────────────────────────────────────

# GitHub repo that receives bug reports as Issues.
# Override via GITHUB_REPO env var if the repo moves.
_GITHUB_REPO = os.environ.get("GITHUB_REPO", "yolinab/Container_Optimisation-web-app")


class BugReportRequest(BaseModel):
    message:         str
    container_count: Optional[int]       = None
    total_pallets:   Optional[int]       = None
    warnings:        Optional[List[str]] = None
    errors:          Optional[List[str]] = None


@app.post("/report-bug")
async def report_bug(body: BugReportRequest):
    """
    Create a GitHub Issue in the app repo as a bug report.

    Required Render environment variable:
      GITHUB_TOKEN — a GitHub Personal Access Token (classic) with the
                     'repo' scope, or a fine-grained token with
                     'Issues: Read and write' on this repository.

    If GITHUB_TOKEN is not set, the report is logged to stdout (Render
    captures this in the service logs) and a 'logged' status is returned
    so the UI still shows a success message to the user.
    """
    token     = os.environ.get("GITHUB_TOKEN", "")
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    # ── Build Markdown issue body ──────────────────────────────────────
    sections: List[str] = []

    sections.append(
        f"**Submitted:** {timestamp}\n"
    )

    sections.append(
        "### User message\n"
        + (body.message.strip() or "_No message provided._")
    )

    if body.container_count is not None:
        sections.append(
            "### Last run summary\n"
            f"| | |\n|---|---|\n"
            f"| Containers packed | {body.container_count} |\n"
            f"| Total pallets | {body.total_pallets} |"
        )

    if body.errors:
        sections.append(
            "### Errors detected\n"
            + "\n".join(f"- ⚠️ {e}" for e in body.errors)
        )

    if body.warnings:
        sections.append(
            "### Warnings\n"
            + "\n".join(f"- {w}" for w in body.warnings)
        )

    issue_body = "\n\n".join(sections)
    issue_title = f"[Bug Report] {(body.message.strip()[:60] + '…') if len(body.message.strip()) > 60 else body.message.strip() or timestamp}"

    # ── Log regardless (Render captures stdout) ───────────────────────
    print(f"[BUG REPORT {timestamp}] {body.message[:200]}")

    if not token:
        # No token configured — report is already logged above
        return {"status": "logged", "note": "Logged to server output (GITHUB_TOKEN not configured)."}

    # ── Post to GitHub Issues API ──────────────────────────────────────
    payload = json.dumps({
        "title":  issue_title,
        "body":   issue_body,
        "labels": ["bug"],
    }).encode()

    req = urllib.request.Request(
        f"https://api.github.com/repos/{_GITHUB_REPO}/issues",
        data=payload,
        headers={
            "Authorization":        f"Bearer {token}",
            "Accept":               "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type":         "application/json",
            "User-Agent":           "container-packing-optimizer",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            resp_data = json.loads(resp.read().decode())
            issue_url = resp_data.get("html_url", "")
        print(f"[BUG REPORT] GitHub issue created: {issue_url}")
    except Exception as exc:
        print(f"[BUG REPORT] GitHub API error: {exc}")
        raise HTTPException(
            status_code=500,
            detail="Could not create bug report. The report has been logged — please contact support if this persists.",
        )

    return {"status": "sent"}
