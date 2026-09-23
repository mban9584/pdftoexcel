import json
import os
import sys
from collections import defaultdict

DIR = os.path.join(WORK, 'cells')
OCR = os.path.join(WORK, 'ocr')
page = int(sys.argv[1])
g = json.load(open(os.path.join(DIR, 'g%d.json' % page), encoding='utf-8'))
r = json.load(open(os.path.join(DIR, 'r%d.json' % page), encoding='utf-8'))
xs, ys, cells = g['xs'], g['ys'], g['cells']


def find(cx, cy):
    for c in cells:
        if xs[c[0]] - 0.01 <= cx < xs[c[2]] - 0.01 and ys[c[1]] - 0.01 <= cy < ys[c[3]] - 0.01:
            return tuple(c)
    return None


B = defaultdict(list)
for it in r:
    x0, y0, x1, y1 = it['box']
    k = find((x0 + x1) / 2, (y0 + y1) / 2)
    if k:
        B[k].append((y0, x0, x1, it['t'], it['s']))

A = defaultdict(list)
full = json.load(open(os.path.join(OCR, 'p%d.json' % page), encoding='utf-8'))
W, H = g['w'] / 1000.0, g['h'] / 1000.0
for b in full:
    cx, cy = (b['x0'] + b['x1']) / 2 * W, (b['y0'] + b['y1']) / 2 * H
    k = find(cx, cy)
    if k:
        A[k].append((b['y0'] * H, b['x0'] * W, b['x1'] * W, b['text']))


def norm(s):
    return ''.join(s.split()).replace(' ', '').replace('.', '').replace('*', 'x').lower()


keys = sorted(set(A) | set(B))
same = only_a = only_b = diff = 0
report = []
for k in keys:
    a = ''.join(t for *_, t in sorted(A.get(k, []), key=lambda z: (round(z[0] / 5), z[1])))
    bl = sorted(B.get(k, []), key=lambda z: (round(z[0] / 5), z[1]))
    bb = ''
    prevline = None
    prevx1 = None
    for y0, x0, x1, t, s in bl:
        line = round(y0 / 5)
        if prevline is not None and line != prevline:
            bb += '\n'
        elif prevx1 is not None and x0 - prevx1 > 2.5:
            bb += ' '
        bb += t
        prevline, prevx1 = line, x1
    na, nb = norm(a), norm(bb)
    if na and nb:
        if na == nb:
            same += 1
        else:
            diff += 1
            report.append(('DIFF', k, repr(a), repr(bb)))
    elif na and not nb:
        only_a += 1
        report.append(('B-MISS', k, repr(a), ''))
    elif nb and not na:
        only_b += 1
        report.append(('A-MISS', k, '', repr(bb)))
print('page', page, 'cells', len(keys), 'agree', same, 'text-differs', diff, 'only-in-fullOCR', only_a, 'only-in-vectorOCR', only_b)
for tag, k, a, b in report[:60]:
    print('  %-7s %s A=%s B=%s' % (tag, k, a, b))
