"""
Desktop entry point - runs the same FastAPI app locally and opens it
in a native window via pywebview. This file is the .exe entry point.
"""
import base64
import subprocess
import sys
import threading
import time
import shutil
import tempfile

from pathlib import Path

import uvicorn
import webview

sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.main import app

HOST = "127.0.0.1"
PORT = 8756


def run_server():
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


class Api:
    """Exposed to JS as window.pywebview.api.* - handles saving files
    since pywebview's embedded window has no real download manager."""

    def _downloads_dir(self) -> Path:
        downloads = Path.home() / "Downloads"
        downloads.mkdir(parents=True, exist_ok=True)
        return downloads

    def save_job_zip(self, job_id: str, kind: str, filename: str) -> dict:
        try:
            job_dir = Path(tempfile.gettempdir()) / "pdfrenamer_jobs" / job_id
            src = job_dir / ("renamed_success.zip" if kind == "success" else "failed_pdfs.zip")
            if not src.exists():
                return {"ok": False, "message": "File not found."}

            downloads = self._downloads_dir()
            dest = downloads / filename
            stem, ext = dest.stem, dest.suffix
            counter = 1
            while dest.exists():
                dest = downloads / f"{stem} ({counter}){ext}"
                counter += 1

            shutil.copy(src, dest)
            self._reveal_in_explorer(dest)
            return {"ok": True, "path": str(dest)}
        except Exception as e:
            return {"ok": False, "message": str(e)}

    def _reveal_in_explorer(self, path: Path):
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(f'explorer /select,"{path}"')
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", str(path)])
            else:
                subprocess.run(["xdg-open", str(path.parent)])
        except Exception:
            pass

def main():
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    time.sleep(1.2)  # give uvicorn a moment to bind the port

    webview.create_window(
        "PDF Renamer",
        f"http://{HOST}:{PORT}",
        width=1100,
        height=800,
        min_size=(900, 650),
        js_api=Api(),
    )
    webview.start()


if __name__ == "__main__":
    main()