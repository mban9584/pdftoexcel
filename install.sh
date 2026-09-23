#!/bin/sh
# Installs the Python side of the pdftoexcel skill on macOS/Linux. Run from this folder.
set -e
command -v python3 >/dev/null || { echo "python3 not found"; exit 1; }
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
python3 -m pip install --user --upgrade pip
python3 -m pip install --user -r "$SCRIPT_DIR/requirements.txt"
python3 - <<'PY'
import importlib.util as u
for m in ('pdfplumber', 'pypdfium2', 'numpy', 'cv2', 'openpyxl', 'PIL'):
    print('%-12s %s' % (m, 'ok' if u.find_spec(m) else 'MISSING'))
print('rapidocr     %s' % ('ok (outline-ocr route available)' if u.find_spec('rapidocr_onnxruntime') else
                           'absent - install requirements-ocr.txt only for PDFs without a text layer'))
PY
echo
echo "Optional OCR support: python3 -m pip install --user -r \"$SCRIPT_DIR/requirements-ocr.txt\""
echo "Done. Keep this folder under ~/.qwen/skills/pdftoexcel so Qwen Code discovers it."
