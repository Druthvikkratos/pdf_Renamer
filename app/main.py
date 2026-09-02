import sys
import traceback
import shutil
import threading
import uuid
from pathlib import Path
from typing import List
import os
from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from app.core.renamer import process_files
from app.core.job_manager import create_job, get_job, run_job



def resource_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).resolve().parent
    return base_path / relative_path


app = FastAPI(title="PDF Renamer")

STATIC_DIR = resource_path("static")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = STATIC_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.post("/api/process/start")
async def start_process(request: Request):
    try:
        form = await request.form(max_files=200_000, max_fields=200_000)
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": f"Could not parse upload: {e}"})

    excel = form.get("excel")
    pdfs = form.getlist("pdfs")

    if not excel or not pdfs:
        return JSONResponse(status_code=400, content={"error": "Missing excel file or pdf files."})

    job = create_job(total=len(pdfs))
    excel_bytes = await excel.read()

    pdf_paths = []
    used = set()
    for f in pdfs:
        safe_name = f.filename
        dest = os.path.join(job.upload_dir, safe_name)
        stem, ext = os.path.splitext(dest)
        counter = 1
        while dest in used:
            dest = f"{stem}__{counter}{ext}"
            counter += 1
        used.add(dest)
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        pdf_paths.append((f.filename, dest))

    thread = threading.Thread(target=run_job, args=(job.id, excel_bytes, pdf_paths), daemon=True)
    thread.start()

    return {"job_id": job.id}

@app.get("/api/process/status/{job_id}")
async def process_status(job_id: str):
    job = get_job(job_id)
    if not job:
        return JSONResponse(status_code=404, content={"error": "job not found"})
    status = job.status_dict()
    if job.done:
        status["result"] = job.result_dict()
    return status

@app.get("/api/process/download/{job_id}/{kind}")
async def download_zip(job_id: str, kind: str):
    job = get_job(job_id)
    if not job:
        return JSONResponse(status_code=404, content={"error": "job not found"})
    path = job.success_zip_path if kind == "success" else job.failed_zip_path
    if not os.path.exists(path):
        return JSONResponse(status_code=404, content={"error": "zip not available"})
    filename = "renamed_success.zip" if kind == "success" else "failed_pdfs.zip"
    return FileResponse(path, media_type="application/zip", filename=filename)