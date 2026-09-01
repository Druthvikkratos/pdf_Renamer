"""
Desktop entry point - runs the same FastAPI app locally and opens it
in a native window via pywebview. This file is the .exe entry point.
"""
import base64
import subprocess
import sys
import threading
import time
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

    def save_file(self, b64_data: str, filename: str) -> dict:
        try:
            downloads = self._downloads_dir()
            save_path = downloads / filename

            # avoid clobbering an existing file - add (1), (2)... suffix
            stem, ext = save_path.stem, save_path.suffix
            counter = 1
            while save_path.exists():
                save_path = downloads / f"{stem} ({counter}){ext}"
                counter += 1

            save_path.write_bytes(base64.b64decode(b64_data))
            self._reveal_in_explorer(save_path)
            return {"ok": True, "path": str(save_path)}
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
            pass  # file is already saved either way - not fatal


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