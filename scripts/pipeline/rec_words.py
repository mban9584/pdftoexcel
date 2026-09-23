"""Recognition-only pass over the vector-derived word boxes (pages 2-4)."""
import json
import os
import sys

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')

import cv2
import numpy as np
import pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR

DIR = os.path.join(WORK, 'cells')
PDF = PDF_PATH
K = 400 / 72.0
PAD = 2.0

ocr = RapidOCR()
ocr.use_text_det = False
ocr.use_angle_cls = False
ocr.text_score = 0.0
doc = pdfium.PdfDocument(PDF)
for i in ([int(x) for x in sys.argv[1:]] or [2, 3, 4]):
    w = json.load(open(os.path.join(DIR, 'w%d.json' % i), encoding='utf-8'))
    arr = np.array(doc[i].render(scale=K).to_pil().convert('RGB'))
    out = []
    for n, box in enumerate(w['words']):
        x0, y0, x1, y1 = box
        a = (int(round((x0 - PAD) * K)), int(round((y0 - PAD) * K)))
        b = (int(round((x1 + PAD) * K)), int(round((y1 + PAD) * K)))
        a = (max(a[0], 0), max(a[1], 0))
        crop = arr[a[1]:b[1], a[0]:b[0]]
        if crop.size == 0:
            out.append({'box': box, 't': '', 's': 0.0})
            continue
        h, wd = crop.shape[:2]
        if h < 20:
            f = 24.0 / h
            crop = cv2.resize(crop, (int(wd * f), 24), interpolation=cv2.INTER_CUBIC)
        res, _ = ocr(crop)
        t, s = (res[0][1], float(res[0][2])) if res else ('', 0.0)
        out.append({'box': [round(v, 2) for v in box], 't': t, 's': round(s, 3)})
        if n % 100 == 0:
            print('  %d/%d' % (n, len(w['words'])), flush=True)
    json.dump(out, open(os.path.join(DIR, 'r%d.json' % i), 'w', encoding='utf-8'), ensure_ascii=False)
    ok = sum(1 for o in out if o['s'] >= 0.9)
    low = [o for o in out if o['s'] < 0.8]
    print('PAGE', i, 'words', len(out), 'score>=0.9', ok, 'score<0.8', len(low))
    for o in sorted(low, key=lambda z: z['s'])[:25]:
        print('   LOW %5.2f %r %s' % (o['s'], o['t'], [round(v) for v in o['box']]))
