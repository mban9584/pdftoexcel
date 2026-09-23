"""Recover per-page table geometry + content from sample.pdf."""
import json
import os
import sys

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')

import cv2
import numpy as np
import pdfplumber as p
import pypdfium2 as pdfium

PDF = PDF_PATH
OUT = os.path.join(WORK, 'cells')
SCALE = 4.0
PAD = 6
MINBOUND = 4.0
os.makedirs(OUT, exist_ok=True)


def segs(pg):
    h, v = [], []
    for o in pg.rects:
        w, hh = o['x1'] - o['x0'], o['bottom'] - o['top']
        if w > 5 and hh <= 3.0:
            h.append((o['top'], o['x0'], o['x1']))
        elif hh > 5 and w <= 3.0:
            v.append((o['x0'], o['top'], o['bottom']))
    for o in pg.curves:
        w, hh = o['x1'] - o['x0'], o['bottom'] - o['top']
        if w > 25 and hh <= 2.0:
            h.append((o['top'], o['x0'], o['x1']))
        elif hh > 25 and w <= 2.0:
            v.append((o['x0'], o['top'], o['bottom']))
    return h, v


def cluster(vals, tol=2.0):
    groups = []
    for x in sorted(vals):
        if groups and x - groups[-1][-1] <= tol:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [sum(g) / len(g) for g in groups]


def dedupe(vals, mindim=MINBOUND):
    out = []
    for x in sorted(vals):
        if not out or x - out[-1] >= mindim:
            out.append(x)
    return out


def nearest(vals, x):
    return min(range(len(vals)), key=lambda k: abs(vals[k] - x))


def detect_cells(pg):
    H = int(round(pg.height * SCALE)) + 2 * PAD
    W = int(round(pg.width * SCALE)) + 2 * PAD
    line = np.zeros((H, W), np.uint8)
    hh, vv = segs(pg)
    for y, a, b in hh:
        Y = int(round(y * SCALE)) + PAD
        cv2.line(line, (int(a * SCALE) + PAD, Y), (int(b * SCALE) + PAD, Y), 255, 3)
    for x, a, b in vv:
        X = int(round(x * SCALE)) + PAD
        cv2.line(line, (X, int(a * SCALE) + PAD), (X, int(b * SCALE) + PAD), 255, 3)
    n, lab, stats, _ = cv2.connectedComponentsWithStats(cv2.bitwise_not(line), 8)
    raw = []
    for i in range(1, n):
        x, y, w, h, _a = stats[i]
        if w < 10 or h < 10:
            continue
        if x <= 1 or y <= 1 or x + w >= W - 1 or y + h >= H - 1:
            continue
        raw.append(((x - PAD) / SCALE, (y - PAD) / SCALE, (x + w - PAD) / SCALE, (y + h - PAD) / SCALE))
    xs = dedupe(cluster([c[0] for c in raw] + [c[2] for c in raw]))
    ys = dedupe(cluster([c[1] for c in raw] + [c[3] for c in raw]))
    cells = set()
    for x0, y0, x1, y1 in raw:
        i0, i1 = nearest(xs, x0), nearest(xs, x1)
        j0, j1 = nearest(ys, y0), nearest(ys, y1)
        if i1 <= i0:
            i1 = i0 + 1
        if j1 <= j0:
            j1 = j0 + 1
        if i1 < len(xs) and j1 < len(ys):
            cells.add((i0, j0, i1, j1))
    cells = sorted(cells, key=lambda c: (c[1], c[0]))
    return cells, xs, ys


def overlaps(cells):
    """Report cell rectangles that intersect each other (would break merged ranges)."""
    bad = []
    for a in range(len(cells)):
        for b in range(a + 1, len(cells)):
            i0, j0, i1, j1 = cells[a]
            k0, l0, k1, l1 = cells[b]
            if i0 < k1 and k0 < i1 and j0 < l1 and l0 < j1:
                bad.append((cells[a], cells[b]))
    return bad


def _rgb(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):
        return (1 - c, 1 - c, 1 - c) if c <= 1 else None
    if len(c) == 1:
        return (1 - c[0],) * 3
    if len(c) == 3:
        return tuple(c)
    if len(c) == 4:  # CMYK
        cy, m, y0, k = c
        return ((1 - cy) * (1 - k), (1 - m) * (1 - k), (1 - y0) * (1 - k))
    return None


def colored_regions(pg):
    """Bounding boxes of non-black glyph outlines (red / blue text)."""
    boxes = []
    for o in pg.curves:
        w, h = o['x1'] - o['x0'], o['bottom'] - o['top']
        if w > 25 or h > 25:
            continue
        rgb = _rgb(o.get('non_stroking_color'))
        if rgb is None:
            continue
        r, g, b = rgb
        if r > 0.45 and g < 0.45 and b < 0.45:
            kind = 'red'
        elif b > 0.45 and r < 0.45:
            kind = 'blue'
        else:
            continue
        boxes.append([kind, round(o['x0'], 1), round(o['top'], 1), round(o['x1'], 1), round(o['bottom'], 1)])
    return boxes


def page_data(pg):
    cells, xs, ys = detect_cells(pg)
    imgs = [[round(o[k], 2) for k in ('x0', 'top', 'x1', 'bottom')] for o in pg.images
            if (o['x1'] - o['x0']) * (o['bottom'] - o['top']) < 0.8 * pg.width * pg.height]
    chars = [{'t': o['text'], 'x0': o['x0'], 'x1': o['x1'], 'top': o['top'], 'bot': o['bottom'],
              'size': round(o.get('size', 9), 2),
              'color': (list(o['non_stroking_color']) if isinstance(o.get('non_stroking_color'), (list, tuple))
                        else [o.get('non_stroking_color')])}
             for o in pg.chars]
    return {'w': round(pg.width, 2), 'h': round(pg.height, 2),
            'xs': [round(v, 2) for v in xs], 'ys': [round(v, 2) for v in ys],
            'cells': [list(c) for c in cells], 'images': imgs, 'chars': chars,
            'colored': colored_regions(pg)}


fp = p.open(PDF)
doc = pdfium.PdfDocument(PDF)
for i in ([int(x) for x in sys.argv[1:]] or [1, 2, 3, 4]):
    d = page_data(fp.pages[i])
    json.dump(d, open(os.path.join(OUT, 'g%d.json' % i), 'w', encoding='utf-8'), ensure_ascii=False)
    ov = overlaps([tuple(c) for c in d['cells']])
    print('PAGE', i, 'cells', len(d['cells']), 'xs', len(d['xs']), 'ys', len(d['ys']),
          'imgs', len(d['images']), 'chars', len(d['chars']), 'colored', len(d['colored']),
          'OVERLAPS', len(ov), ov[:4])
    k = 150 / 72.0
    arr = np.array(doc[i].render(scale=k).to_pil().convert('RGB'))
    for c in d['cells']:
        x0, y0 = d['xs'][c[0]], d['ys'][c[1]]
        x1, y1 = d['xs'][c[2]], d['ys'][c[3]]
        cv2.rectangle(arr, (int(x0 * k), int(y0 * k)), (int(x1 * k), int(y1 * k)), (0, 160, 255), 2)
    cv2.imwrite(os.path.join(OUT, 'v%d.png' % i), arr[:, :, ::-1])
fp.close()
