"""Build the per-page content model: cell text, colours, sizes, free text, images.

Pages with a text layer (0,1) use it directly; pages 2-4 use per-cell OCR output
(o{i}.json, produced by content.py) plus the full-page OCR (ocr/p{i}.json) for text
that sits outside the table grid.
"""
import json
import os
import re
import sys
from collections import defaultdict

DIR = os.path.join(WORK, 'cells')
OCRDIR = os.path.join(WORK, 'ocr')


def clean(s):
    """Undo the spurious spaces the recognisers inject inside numbers: 110. 00 -> 110.00."""
    s = re.sub(r'(\d)\s*[.。]\s*(\d)', r'\1.\2', s)
    s = re.sub(r'\s*([（(])\s*', r'\1', s)
    s = re.sub(r'\s*([)）])\s*', r'\1', s)
    return s


def rgb_of(c):
    if c is None:
        return 'black'
    if isinstance(c, (int, float)):
        return 'black' if c < 0.6 else 'white'
    if len(c) == 1:
        return 'black' if c[0] < 0.6 else 'white'
    if len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.45 and g < 0.45 and b < 0.45:
            return 'red'
        if b > 0.45 and r < 0.45:
            return 'blue'
    return 'black'


def join_line(recs):
    """recs: (x0, x1, text, size, color) -> (string, size, color)."""
    recs = sorted(recs, key=lambda r: r[0])
    out, prev = '', None
    for x0, x1, t, sz, col in recs:
        if prev is not None and x0 - prev > 1.6:
            out += ' '
        out += t
        prev = x1
    w = defaultdict(float)
    for _, _, t, sz, _ in recs:
        w[round(sz, 1)] += len(t)
    c = defaultdict(int)
    for *_, col in recs:
        c[col] += 1
    color = 'black' if c.get('black', 0) * 2 >= len(recs) else max(c, key=c.get)
    return out.strip(), (max(w, key=w.get) if w else 9.0), color


def find_cell(cells, xs, ys, cx, cy):
    for c in cells:
        if xs[c[0]] - 0.01 <= cx < xs[c[2]] - 0.01 and ys[c[1]] - 0.01 <= cy < ys[c[3]] - 0.01:
            return tuple(c)
    return None


def cluster(vals, tol=2.0):
    groups = []
    for x in sorted(vals):
        if groups and x - groups[-1][-1] <= tol:
            groups[-1].append(x)
        else:
            groups.append([x])
    return [sum(g) / len(g) for g in groups]


def dedupe(vals, mindim=4.0):
    out = []
    for x in sorted(vals):
        if not out or x - out[-1] >= mindim:
            out.append(x)
    return out


def nearest(vals, x):
    return min(range(len(vals)), key=lambda k: abs(vals[k] - x))


def color_of(box, colored):
    """red/blue if a majority of coloured glyph boxes inside this region agree."""
    hits = defaultdict(int)
    for kind, x0, y0, x1, y1 in colored:
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if box[0] <= cx < box[2] and box[1] <= cy < box[3]:
            hits[kind] += 1
    return max(hits, key=hits.get) if hits else 'black'


def text_lines(chars):
    """Group chars (with an already-resolved 'col' name) into visual lines."""
    lines = defaultdict(list)
    for ch in chars:
        lines[round((ch['top'] + ch['bot']) / 2 / 3.0)].append(ch)
    out = []
    for _, recs in sorted(lines.items()):
        joined = [(r['x0'], r['x1'], r['t'], r['size'], r['col']) for r in recs]
        t, sz, col = join_line(joined)
        if not t:
            continue
        out.append({'text': t, 'size': sz, 'color': col,
                    'x0': min(r['x0'] for r in recs), 'x1': max(r['x1'] for r in recs),
                    'y0': min(r['top'] for r in recs), 'y1': max(r['bot'] for r in recs)})
    return out


def from_textlayer(g):
    xs, ys, cells = g['xs'], g['ys'], g['cells']
    model = {'page': g['page'], 'w': g['w'], 'h': g['h'], 'xs': xs, 'ys': ys,
             'cells': [{'cell': c, 'text': '', 'size': 9.0, 'color': 'black'} for c in cells],
             'free': [], 'images': g['images']}
    incell = defaultdict(list)
    outside = []
    for ch in g['chars']:
        cx = (ch['x0'] + ch['x1']) / 2
        cy = (ch['top'] + ch['bot']) / 2
        c = find_cell(cells, xs, ys, cx, cy)
        rec = (ch['x0'], ch['x1'], ch['t'], ch['size'], rgb_of(ch['color']), ch['top'], ch['bot'])
        (incell[tuple(c)] if c else outside).append((cy, rec))
    for k, c in enumerate(cells):
        got = incell.get(tuple(c), [])
        if not got:
            continue
        lines = defaultdict(list)
        for cy, rec in got:
            lines[round(cy / 3.0)].append(rec[:5])
        txt, szs, cols = [], [], []
        for _, recs in sorted(lines.items()):
            t, sz, col = join_line(recs)
            txt.append(t)
            szs.append(sz)
            cols.append(col)
        model['cells'][k]['text'] = '\n'.join(x for x in txt if x)
        model['cells'][k]['size'] = max(szs, key=szs.count)
        model['cells'][k]['color'] = max(set(cols), key=cols.count)
    model['free'] = text_lines([{'x0': r[0], 'x1': r[1], 't': r[2], 'size': r[3], 'col': r[4],
                                 'top': r[5], 'bot': r[6]} for _, r in outside]) if outside else []
    return model


def _cover(items, b):
    """Total width of the union of item x-intervals, clipped to b's interval."""
    segs = sorted((max(i[0], b[0]), min(i[2], b[2])) for i in items)
    total, cur = 0.0, None
    for lo, hi in segs:
        if hi <= lo:
            continue
        if cur is not None and lo <= cur[1]:
            cur = (cur[0], max(cur[1], hi))
        else:
            if cur is not None:
                total += cur[1] - cur[0]
            cur = (lo, hi)
    if cur is not None:
        total += cur[1] - cur[0]
    return total


def from_ocr(g, page):
    """Union of two recognition paths: (A) full-page det+rec, (B) vector-derived
    word boxes + recognition. B catches tiny numbers A drops; A catches labels B's
    glyph clustering misses."""
    xs, ys, cells = g['xs'], g['ys'], g['cells']
    W, H = g['w'], g['h']
    full = json.load(open(os.path.join(OCRDIR, 'p%d.json' % page), encoding='utf-8'))
    words = json.load(open(os.path.join(DIR, 'r%d.json' % page), encoding='utf-8'))
    model = {'page': page, 'w': W, 'h': H, 'xs': xs, 'ys': ys, 'cells': [], 'free': [],
             'images': g['images']}
    A = []
    for b in full:
        A.append((b['x0'] / 1000.0 * W, b['y0'] / 1000.0 * H,
                  b['x1'] / 1000.0 * W, b['y1'] / 1000.0 * H, b['text'], b['score'], 'A'))
    B = [(w['box'][0], w['box'][1], w['box'][2], w['box'][3], w['t'], w['s'], 'B') for w in words if w['t']]
    toks = A + [t for t in B if not any(
        t[4] and a[4] and min(t[2], a[2]) - max(t[0], a[0]) > 0.6 * (t[2] - t[0])
        and min(t[3], a[3]) - max(t[1], a[1]) > 0.5 * (t[3] - t[1]) for a in A)]
    # A det can split one word into fragments while the vector path emits it as a single box;
    # the union above then writes the word twice ('S200S200 C款）C款'). Where one B box is
    # >=80% covered by two or more co-linear A boxes, keep B and drop those fragments.
    drop = set()
    for b in (t for t in toks if t[6] == 'B'):
        same = [a for a in toks if a[6] == 'A' and a[4].strip()
                and min(a[3], b[3]) - max(a[1], b[1]) > 0.5 * min(a[3] - a[1], b[3] - b[1])
                and min(a[2], b[2]) - max(a[0], b[0]) > 0]
        if len(same) < 2:
            continue
        covered = _cover(same, b)
        if covered >= 0.8 * (b[2] - b[0]):
            drop.update(id(s) for s in same)
    if drop:
        toks = [t for t in toks if id(t) not in drop]
    heights = [t[3] - t[1] for t in toks if t[4].strip()]
    med = sorted(heights)[len(heights) // 2] if heights else 11.0
    seen = defaultdict(list)
    for t in toks:
        cx, cy = (t[0] + t[2]) / 2, (t[1] + t[3]) / 2
        k = find_cell(cells, xs, ys, cx, cy)
        seen[k].append(t)
    for c in cells:
        got = seen.get(tuple(c), [])
        lines = defaultdict(list)
        for t in got:
            lines[round(((t[1] + t[3]) / 2) / 6.0)].append(t)
        txt = []
        for _, its in sorted(lines.items()):
            its.sort(key=lambda z: z[0])
            s, prev = '', None
            for t in its:
                if prev is not None and t[0] - prev > 2.2:
                    s += ' '
                s += t[4]
                prev = t[2]
            if s.strip():
                txt.append(s.strip())
        box = (xs[c[0]], ys[c[1]], xs[c[2]], ys[c[3]])
        model['cells'].append({'cell': c, 'text': clean('\n'.join(txt)), 'size': round(med / 0.72, 1),
                               'color': color_of(box, g['colored']),
                               'src': ''.join(sorted({t[6] for t in got}))})
    model['free'] = [{'text': clean(t[4]), 'size': round((t[3] - t[1]) / 0.72, 1),
                      'color': color_of(t[:4], g['colored']), 'x0': t[0], 'x1': t[2], 'y0': t[1], 'y1': t[3]}
                     for t in sorted(seen.get(None, []), key=lambda z: (round(z[1] / 6), z[0])) if t[4].strip()]
    return model, med


def cover(g, page):
    """No table grid: build a grid from image and text edges."""
    lines = text_lines([dict(c, col=rgb_of(c['color'])) for c in g['chars']])
    edges = []
    for im in g['images']:
        edges += [im[0], im[2]]
    for l in lines:
        edges += [l['x0'], l['x1']]
    xs = dedupe(cluster(edges, 2.0), 4.0)
    yedges = []
    for im in g['images']:
        yedges += [im[1], im[3]]
    for l in lines:
        yedges += [l['y0'], l['y1']]
    ys = dedupe(cluster(yedges, 2.0), 4.0)
    for l in lines:
        l['i0'] = nearest(xs, l['x0'])
        l['i1'] = max(l['i0'] + 1, nearest(xs, l['x1']) + 1)
        l['j0'] = nearest(ys, l['y0'])
        l['j1'] = max(l['j0'] + 1, nearest(ys, l['y1']) + 1)
    return {'page': page, 'w': g['w'], 'h': g['h'], 'xs': xs, 'ys': ys, 'cells': [],
            'free': lines, 'images': g['images']}


for i in ([int(x) for x in sys.argv[1:]] or [0, 1, 2, 3, 4]):
    g = json.load(open(os.path.join(DIR, 'g%d.json' % i), encoding='utf-8'))
    g['page'] = i
    med = None
    if i == 0:
        m = cover(g, i)
    elif g['chars']:
        m = from_textlayer(g)
    else:
        m, med = from_ocr(g, i)
        for f in m['free']:
            f['size'] = round((f['y1'] - f['y0']) / 0.72, 1)
    for f in m['free']:
        if f.get('size') is None:
            f['size'] = 12.0
    json.dump(m, open(os.path.join(DIR, 'm%d.json' % i), 'w', encoding='utf-8'), ensure_ascii=False)
    filled = sum(1 for c in m['cells'] if c['text'].strip())
    print('PAGE', i, 'cells', len(m['cells']), 'with-text', filled, '| grid', len(m['xs']), 'x', len(m['ys']),
          '| free', len(m['free']), '| imgs', len(m['images']), '| ocr-median-pt', med)
    for f in sorted(m['free'], key=lambda z: z['y0']):
        print('   free %5.1fpt %s %s' % (f['size'], f['color'], repr(f['text'])))
