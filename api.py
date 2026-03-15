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

# Make app/ importable without installing as a package
sys.path.insert(0, str(Path(__file__).parent / "app"))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from pipeline import run_pipeline

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
        tmp_path   = Path(tmp)
        excel_path = tmp_path / file.filename

        with excel_path.open("wb") as fh:
            shutil.copyfileobj(file.file, fh)

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

        report_b64    = base64.b64encode(result["report_path"].read_bytes()).decode()
        issue_summary = _issues_to_frontend(result.get("validation_issues", []))
        containers    = result["containers"]

    return JSONResponse({
        "report_b64":      report_b64,
        "filename":        "packing_report.xlsx",
        "container_count": len(containers),
        "total_pallets":   sum(c.get("loaded_value", 0) for c in containers),
        "warnings":        issue_summary["warnings"],
        "errors":          issue_summary["errors"],
    })


# ── Bug report ─────────────────────────────────────────────────────────────

class BugReportRequest(BaseModel):
    message:         str
    container_count: Optional[int]  = None
    total_pallets:   Optional[int]  = None
    warnings:        Optional[List[str]] = None
    errors:          Optional[List[str]] = None


@app.post("/report-bug")
async def report_bug(body: BugReportRequest):
    """
    Send a bug report email via Resend (https://resend.com).

    Required Render environment variable:
      RESEND_API_KEY — API key from resend.com dashboard
    Optional:
      BUG_EMAIL      — recipient address (defaults to yolina.yordanova@edelman.nl)
    """
    api_key   = os.environ.get("RESEND_API_KEY", "")
    bug_email = os.environ.get("BUG_EMAIL", "yolina.yordanova@edelman.nl")
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    if not api_key:
        # Gracefully degrade — log to stdout so Render captures it
        print(f"[BUG REPORT {timestamp}] RESEND_API_KEY not set. Message: {body.message}")
        return {"status": "logged", "note": "Email not configured — report logged to server output."}

    # Build plain-text body
    lines = [
        "Bug Report — Container Packing Optimizer",
        f"Submitted: {timestamp}",
        "",
        "━" * 52,
        "USER MESSAGE",
        "━" * 52,
        body.message.strip() or "(no message provided)",
        "",
    ]

    if body.container_count is not None:
        lines += [
            "━" * 52,
            "LAST RUN SUMMARY",
            "━" * 52,
            f"Containers packed : {body.container_count}",
            f"Total pallets     : {body.total_pallets}",
            "",
        ]

    if body.errors:
        lines += ["━" * 52, "ERRORS DETECTED", "━" * 52]
        lines += [f"  !! {e}" for e in body.errors]
        lines.append("")

    if body.warnings:
        lines += ["━" * 52, "WARNINGS", "━" * 52]
        lines += [f"  •  {w}" for w in body.warnings]
        lines.append("")

    payload = json.dumps({
        "from":    "Container Optimizer <onboarding@resend.dev>",
        "to":      [bug_email],
        "subject": f"[Bug Report] Container Optimizer — {timestamp}",
        "text":    "\n".join(lines),
    }).encode()

    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=payload,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type":  "application/json",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status not in (200, 201):
                raise RuntimeError(f"Resend returned HTTP {resp.status}")
    except Exception as exc:
        print(f"[BUG REPORT] Failed to send email: {exc}")
        raise HTTPException(
            status_code=500,
            detail="Could not send report — please try again or contact support directly.",
        )

    return {"status": "sent"}
