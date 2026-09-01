"""
Desktop entry point - runs the same FastAPI app locally and opens it
in a native window via pywebview. This file is the .exe entry point.
"""
import threading
import time
import sys
from pathlib import Path

import uvicorn
import webview

sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.main import app

HOST = "127.0.0.1"
PORT = 8756


def run_server():
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


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
    )
    webview.start()


if __name__ == "__main__":
    main()