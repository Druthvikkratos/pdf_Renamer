# PDF Renamer

## 1. Local setup (venv)

python -m venv venv

# Windows
venv\Scripts\activate
# Mac/Linux
source venv/bin/activate

pip install -r requirements.txt

## 2. Run locally (web mode)

uvicorn app.main:app --reload --port 8000

Open http://127.0.0.1:8000 in your browser.

## 3. Deploy to Render

- Push this repo to GitHub.
- On Render: New -> Web Service -> connect repo.
- Render auto-detects render.yaml (build: pip install -r requirements.txt,
  start: uvicorn app.main:app --host 0.0.0.0 --port $PORT).
- Deploy. Done.

## 4. Build the desktop .exe

pip install -r requirements-desktop.txt

# Windows
pyinstaller --onefile --noconsole --name PDFRenamer --add-data "app/static;static" desktop.py

# Mac/Linux
pyinstaller --onefile --noconsole --name PDFRenamer --add-data "app/static:static" desktop.py

The exe/binary will be in dist/PDFRenamer(.exe).
Double-click it - it opens the same UI in a native window, fully offline.