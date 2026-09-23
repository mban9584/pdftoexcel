"""Verify the workbook against the source model.

1. geometry: rebuild pt boundaries from the stored column widths / row heights and
   compare with the original grid (tells how far photos and merges can drift).
2. content: for every table cell, re-read the two independent recognition paths
   (A = full-page det+rec, B = vector word boxes) and compare with what was written.
"""
import json
import os
import re
import sys

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')
from bisect import bisect_right

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

DIR = os.path.join(WORK, 'cells')
OCRDIR = os.path.join(WORK, 'ocr')
OUT = OUT_PATH
wb = load_workbook(OUT)


def norm(s):
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'[（(]', '(', s).replace('）', ')').replace('，', ',').replace('：', ':').lower()
    return re.sub(r'^(\d+)\.0+$', r'\1', s)


def num(s):
    return float(s) if re.fullmatch(r'\d+(\.\d+)?', s) else None


def same(a, b):
    if a == b:
        return True
    x, y = num(a), num(b)
    return x is not None and y is not None and x == y




for page in ([int(x) for x in sys.argv[1:]] or [0, 1, 2, 3, 4]):
    m = json.load(open(os.path.join(DIR, 'm%d.json' % page), encoding='utf-8'))
    W, H = float(m['w']), float(m['h'])
    XG = [0.0] + [float(x) for x in m['xs']] + [W]
    YG = [0.0] + [float(y) for y in m['ys']] + [H]
    ncol, nrow = len(XG) - 1, len(YG) - 1
    ws = wb.worksheets[page]

    orig_x = [0.0] + [float(x) for x in m['xs']] + [W]
    realized, run = [], 0
    for k in range(ncol):
        dim = ws.column_dimensions.get(get_column_letter(k + 1))
        run += round(dim.width * 7 + 5) if dim else 64
        realized.append(run * 0.75)
    ys = [sum((ws.row_dimensions[r].height or 15) for r in range(1, k + 2)) for k in range(nrow)]
    dx = [abs(a - b) for a, b in zip(realized, orig_x[1:])]
    dy = [abs(a - b) for a, b in zip(ys, YG[1:])]
    print('p%d %-22s geometry: col drift max %.2fpt mean %.2fpt | row drift max %.2fpt mean %.2fpt' %
          (page, ws.title, max(dx), sum(dx) / len(dx), max(dy), sum(dy) / len(dy)))

    if page <= 1:      # text-layer page: only check that the sheet matches the model
        bad = 0
        for c in m['cells']:
            i0, j0 = c['cell'][0], c['cell'][1]
            val = ws.cell(row=j0 + 2, column=i0 + 2).value
            if not same(norm('' if val is None else str(val)), norm(c['text'])):
                bad += 1
                if bad <= 15:
                    print('   !! %s%d sheet=%r model=%r' % (get_column_letter(i0 + 2), j0 + 2, val, c['text']))
        print('   written-vs-model mismatches: %d' % bad)
        continue
    A = json.load(open(os.path.join(OCRDIR, 'p%d.json' % page), encoding='utf-8'))
    A = [(b['x0'] / 1000.0 * W, b['y0'] / 1000.0 * H, b['x1'] / 1000.0 * W, b['y1'] / 1000.0 * H, b['text']) for b in A]
    B = [(w['box'][0], w['box'][1], w['box'][2], w['box'][3], w['t'])
         for w in json.load(open(os.path.join(DIR, 'r%d.json' % page), encoding='utf-8'))]

    def line_of(toks, box):
        keep = [t for t in toks if t[4].strip() and box[0] - 1 <= (t[0] + t[2]) / 2 < box[2] - 1
                and box[1] - 1 <= (t[1] + t[3]) / 2 < box[3] - 1]
        return norm(''.join(t[4] for t in sorted(keep, key=lambda z: (round(z[1] / 6), z[0]))))

    diff = 0
    for c in m['cells']:
        i0, j0, i1, j1 = c['cell']
        box = (XG[i0 + 1], YG[j0 + 1], XG[i1 + 1], YG[j1 + 1])
        val = ws.cell(row=j0 + 2, column=i0 + 2).value
        got = norm('' if val is None else str(val))
        mine = norm(c['text'])
        if not same(got, mine) and not (got == '/' and mine == ''):
            print('   !! written differs from model at %s%d: %r vs %r' %
                  (get_column_letter(i0 + 2), j0 + 2, got, mine))
        a, b = line_of(A, box), line_of(B, box)
        if a and b and not same(a, b) and not same(a, got) and not same(b, got):
            diff += 1
            if diff <= 25:
                print('   ?? cell(%s) x%.0f y%.0f written=%r  A=%r  B=%r' %
                      (c['cell'], box[0], box[1], got, a, b))
    print('   cells where both paths disagree with the written text: %d' % diff)
