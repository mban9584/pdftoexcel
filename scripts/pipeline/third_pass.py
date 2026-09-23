"""Third, independent read of every numeric cell: crop the ink from a 400dpi render and run
recognition only. A and B agreeing is not proof - both can misread the same glyph.
"""
import json
import os
import re
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
ocr = RapidOCR()
ocr.use_text_det = False
ocr.use_angle_cls = False
ocr.text_score = 0.0
doc = pdfium.PdfDocument(PDF)
NUM = re.compile(r'^\d+(\.\d+)?$')


def norm(s):
    s = re.sub(r'\s+', '', s)
    s = s.replace('．', '.').replace('。', '.').replace('，', '.').replace('、', '.')
    if NUM.match(s):
        f = float(s)
        return ('%d' % f if f == int(f) else '%g' % f)
    return s


for p in ([int(x) for x in sys.argv[1:]] or [2, 3, 4]):
    m = json.load(open(os.path.join(DIR, 'm%d.json' % p), encoding='utf-8'))
    rgb = np.asarray(doc[p].render(scale=K).to_pil().convert('RGB'))
    arr = np.asarray(doc[p].render(scale=K).to_pil().convert('L'))
    g = arr < 150
    xs, ys = m['xs'], m['ys']
    todo = [c for c in m['cells'] if NUM.match(c['text'].strip().replace('\n', '')) or c['text'].strip() == '/']
    print('p%d numeric cells to re-read: %d' % (p, len(todo)), flush=True)
    bad = 0
    for n, c in enumerate(todo):
        i0, j0, i1, j1 = c['cell']
        x0, y0, x1, y1 = int(xs[i0] * K), int(ys[j0] * K), int(xs[i1] * K), int(ys[j1] * K)
        sub = g[y0 + 3:y1 - 3, x0 + 3:x1 - 3]
        if sub.size == 0:
            continue
        rows = np.where(sub.any(axis=1) & (sub.mean(axis=1) < 0.85))[0]
        cols = np.where(sub.any(axis=0) & (sub.mean(axis=0) < 0.85))[0]
        if len(rows) < 4 or len(cols) < 2:
            continue
        crop = rgb[y0 + 3 + rows[0]:y0 + 3 + rows[-1] + 1, x0 + 3 + cols[0]:x0 + 3 + cols[-1] + 1]
        h, w = crop.shape[:2]
        if h < 40:
            f = 40.0 / h
            crop = cv2.resize(crop, (max(int(w * f), 8), 40), interpolation=cv2.INTER_CUBIC)
        pad = np.full((crop.shape[0] + 12, crop.shape[1] + 12, 3), 255, dtype=np.uint8)
        pad[6:-6, 6:-6] = crop
        res, _ = ocr(pad)
        t = res[0][1] if res else ''
        s = float(res[0][2]) if res else 0.0
        if norm(t) != norm(c['text'].strip()):
            bad += 1
            print('   DIFF cell(%s) written=%-12r third=%-12r score=%.2f' %
                  (c['cell'], c['text'].strip(), t, s), flush=True)
        if n % 100 == 0:
            print('  %d/%d' % (n, len(todo)), flush=True)
    print('p%d third-path mismatches: %d' % (p, bad))
