@echo off
rem Installs the Python side of the pdftoexcel skill on Windows. Run from this folder.
setlocal
where python >nul 2>&1
if errorlevel 1 (
  echo Python was not found on PATH. Install Python 3.10+ from python.org, then re-run.
  exit /b 1
)
python -m pip install --upgrade pip
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo Install failed. If you are behind a proxy, try: python -m pip install --proxy http://127.0.0.1:PORT -r requirements.txt
  exit /b 1
)
python - <<'PY'
import importlib
for m in ('pdfplumber', 'pypdfium2', 'numpy', 'cv2', 'openpyxl', 'PIL'):
    print('%-12s %s' % (m, 'ok' if importlib.util.find_spec(m) else 'MISSING'))
try:
    import rapidocr_onnxruntime
    print('rapidocr     ok (outline-ocr route available)')
except ImportError:
    print('rapidocr     absent - only needed for PDFs without a text layer')
PY
echo.
echo Done. The skill itself needs no install: keep this folder under %%USERPROFILE%%\.qwen\skills\pdftoexcel
endlocal
