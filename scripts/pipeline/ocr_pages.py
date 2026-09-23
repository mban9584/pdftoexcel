import json
import os
import sys

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')

import pypdfium2 as pdfium
import numpy as np
from rapidocr_onnxruntime import RapidOCR

PDF = PDF_PATH
OUT = os.path.join(WORK, 'ocr')
os.makedirs(OUT, exist_ok=True)
pages = [int(x) for x in sys.argv[1:]] or [2, 3, 4]

ocr = RapidOCR()
doc = pdfium.PdfDocument(PDF)
for i in pages:
    bmp = doc[i].render(scale=300 / 72.0)
    img = np.array(bmp.to_pil().convert('RGB'))
    res, _ = ocr(img)
    items = []
    for box, text, score in (res or []):
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        items.append({
            'text': text,
            'score': round(float(score), 3),
            'x0': round(min(xs) / img.shape[1] * 1000, 1),
            'x1': round(max(xs) / img.shape[1] * 1000, 1),
            'y0': round(min(ys) / img.shape[0] * 1000, 1),
            'y1': round(max(ys) / img.shape[0] * 1000, 1),
        })
    with open(os.path.join(OUT, f'p{i}.json'), 'w', encoding='utf-8') as fh:
        json.dump(items, fh, ensure_ascii=False, indent=1)
    print('page', i, 'boxes', len(items))
    bmp.close()
