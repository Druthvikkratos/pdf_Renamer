import sys
import traceback
from pathlib import Path
from typing import List

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.core.renamer import process_files


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


@app.post("/api/process")
async def process(excel: UploadFile = File(...), pdfs: List[UploadFile] = File(...)):
    try:
        excel_bytes = await excel.read()

        pdf_files = []
        for f in pdfs:
            content = await f.read()
            pdf_files.append((f.filename, content))

        result = process_files(excel_bytes, pdf_files)
        return JSONResponse(result)

    except Exception as e:
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={
                "excel_error": f"Unexpected server error: {type(e).__name__}: {e}",
                "excel_warnings": [], "excel_total_rows": 0, "excel_skipped_rows": 0,
                "excel_matched_count": 0, "excel_unmatched": [], "success": [], "failed": [],
                "success_zip_b64": None, "failed_zip_b64": None,
                "total_pdfs": len(pdfs) if pdfs else 0,
            }
        )