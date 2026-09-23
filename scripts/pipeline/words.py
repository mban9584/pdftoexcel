"""Recover word bounding boxes on pages without a text layer by clustering the
vector glyph outlines (curves) that make up each run of text."""
import json
import os
import sys

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')

import pdfplumber as p

DIR = os.path.join(WORK, 'cells')
MAXGLYPH = 26.0


def boxes(curves):
    out = []
    for c in curves:
        x0, y0, x1, y1 = c['x0'], c['top'], c['x1'], c['bottom']
        if 0.4 < (x1 - x0) <= MAXGLYPH and 0.4 < (y1 - y0) <= MAXGLYPH:
            out.append([x0, y0, x1, y1])
    return out


def merge_glyphs(bs, pad=0.6):
    """Union-find: curves that overlap in both x and y belong to one glyph."""
    n = len(bs)
    par = list(range(n))

    def find(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a

    def uni(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            par[rb] = ra

    bs.sort()
    for i in range(n):
        for j in range(i + 1, n):
            if bs[j][0] - bs[i][2] > pad:
                break
            xo = min(bs[i][2], bs[j][2]) - max(bs[i][0], bs[j][0])
            yo = min(bs[i][3], bs[j][3]) - max(bs[i][1], bs[j][1])
            wi = min(bs[i][2] - bs[i][0], bs[j][2] - bs[j][0])
            hi = min(bs[i][3] - bs[i][1], bs[j][3] - bs[j][1])
            if xo > 0.3 * wi and yo > 0.3 * hi:
                uni(i, j)
    g = {}
    for i in range(n):
        r = find(i)
        b = g.setdefault(r, [bs[i][0], bs[i][1], bs[i][2], bs[i][3]])
        b[0] = min(b[0], bs[i][0])
        b[1] = min(b[1], bs[i][1])
        b[2] = max(b[2], bs[i][2])
        b[3] = max(b[3], bs[i][3])
    return sorted(g.values(), key=lambda b: (b[1], b[0]))


def merge_words(gl, gapfrac=0.45, maxgap=9.0):
    """Join glyphs on the same line when they sit close together."""
    words = []
    used = [False] * len(gl)
    for i, a in enumerate(gl):
        if used[i]:
            continue
        w = list(a)
        used[i] = True
        changed = True
        while changed:
            changed = False
            for j, b in enumerate(gl):
                if used[j]:
                    continue
                yo = min(w[3], b[3]) - max(w[1], b[1])
                h = min(w[3] - w[1], b[3] - b[1])
                if yo < 0.5 * h:
                    continue
                gap = max(b[0] - w[2], w[0] - b[2])
                if gap <= max(gapfrac * h, maxgap):
                    w = [min(w[0], b[0]), min(w[1], b[1]), max(w[2], b[2]), max(w[3], b[3])]
                    used[j] = True
                    changed = True
        words.append(w)
    return sorted(words, key=lambda b: (round(b[1] / 4.0), b[0]))


fp = p.open(PDF_PATH)
for i in ([int(x) for x in sys.argv[1:]] or [2, 3, 4]):
    pg = fp.pages[i]
    gl = merge_glyphs(boxes(pg.curves))
    wd = merge_words(gl)
    json.dump({'glyphs': [[round(v, 2) for v in g] for g in gl],
               'words': [[round(v, 2) for v in w] for w in wd]},
              open(os.path.join(DIR, 'w%d.json' % i, ), 'w', encoding='utf-8'), ensure_ascii=False)
    print('PAGE', i, 'curves', len(pg.curves), 'glyphs', len(gl), 'words', len(wd))
    hs = sorted(w[3] - w[1] for w in wd)
    print('   word h: min %.1f med %.1f max %.1f' % (hs[0], hs[len(hs) // 2], hs[-1]))
    print('   sample', [[round(v) for v in w] for w in wd[:8]])
fp.close()
