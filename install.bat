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
python -c "import importlib.util as u; mods=('pdfplumber','pypdfium2','numpy','cv2','openpyxl','PIL'); [print(m.ljust(12), 'ok' if u.find_spec(m) else 'MISSING') for m in mods]; print('rapidocr     ' + ('ok (outline-ocr route available)' if u.find_spec('rapidocr_onnxruntime') else 'absent - install requirements-ocr.txt only for PDFs without a text layer'))"
echo.
echo Optional OCR support: python -m pip install -r "%~dp0requirements-ocr.txt"
echo Done. The skill itself needs no install: keep this folder under %%USERPROFILE%%\.qwen\skills\pdftoexcel
endlocal
