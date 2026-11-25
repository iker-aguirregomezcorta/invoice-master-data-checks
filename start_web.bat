@echo off
REM Double-click to open the invoice checks in your browser.
cd /d "%~dp0"
python -c "import fastapi, uvicorn, pdfplumber, openpyxl, multipart" 2>nul || (
  echo Installing dependencies, first time only...
  python -m pip install -r requirements.txt
)
python web.py
pause
