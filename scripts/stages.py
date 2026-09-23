"""The two stages that turn a PDF into a layout-faithful workbook.

    model(g*.json [, ocr/*.json, words/*.json]) -> cells/m*.json
    build(cells/m*.json + the PDF)              -> the .xlsx

No path is hard-coded anywhere in this file: everything comes from _paths, so a
copy of this skill folder works on any machine. See pdftoexcel.py for the CLI.
"""
import json
import os
import re
from bisect import bisect_right

import numpy as np
import pdfplumber
import pypdfium2 as pdfium
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

EMU_PT = 12700.0
INK_SCALE = 4.0
CJK_INK_RATIO = 0.917        # verify on a new machine/font: see SKILL.md TODO(verify)
CAP_INK_RATIO = 0.681
SLASH_W, SLASH_H = (2.5, 9.0), (7.0, 14.0)
MIN_COL_PX = 6               # below this a column cannot be expressed in width units
COLOR = {'red': 'FFFF0000', 'blue': 'FF0000FF', 'black': 'FF000000', 'white': 'FFFFFFFF'}


# ---------------------------------------------------------------- geometry utils
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


def find_cell(cells, xs, ys, cx, cy):
    for c in cells:
        if xs[c[0]] - .01 <= cx < xs[c[2]] - .01 and ys[c[1]] - .01 <= cy < ys[c[3]] - .01:
            return tuple(c)
    return None


def snap_grid(vals, total):
    """Boundaries in pt snapped onto the pixel grid, with a MIN_COL_PX floor."""
    b = [0] + [int(round(v * 4 / 3)) for v in list(vals) + [total]]
    for k in range(1, len(b)):
        b[k] = max(b[k], b[k - 1] + MIN_COL_PX)
    return [x * 0.75 for x in b], [b[k + 1] - b[k] for k in range(len(b) - 1)]


def span(a, b, G, n):
    k0 = max(0, min(bisect_right(G, a + 0.4) - 1, n - 1))
    return k0, max(k0, min(bisect_right(G, b - 0.4) - 1, n - 1))


# ------------------------------------------------------------------- text utils
def clean(s):
    s = re.sub(r'(\d)\s*[.。]\s*(\d)', r'\1.\2', s)
    s = re.sub(r'\s*([（(])\s*', r'\1', s)
    return re.sub(r'\s*([)）])\s*', r'\1', s)


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
    """(x0, x1, text, size, color) per glyph -> one line. Drops stroke+fill reprints."""
    recs = sorted(recs, key=lambda r: r[0])
    uniq = []
    for r in recs:
        if uniq and r[2] == uniq[-1][2] and abs(r[0] - uniq[-1][0]) < 0.6:
            continue
        uniq.append(r)
    recs = uniq
    out, prev = '', None
    for x0, x1, t, sz, col in recs:
        if prev is not None and x0 - prev > 1.6:
            out += ' '
        out += t
        prev = x1
    w = {}
    for _, _, t, sz, _ in recs:
        w[round(sz, 1)] = w.get(round(sz, 1), 0) + len(t)
    c = {}
    for *_, col in recs:
        c[col] = c.get(col, 0) + 1
    color = 'black' if c.get('black', 0) * 2 >= len(recs) else max(c, key=c.get)
    return out.strip(), (max(w, key=w.get) if w else 9.0), color


def text_lines(chars):
    lines = {}
    for ch in chars:
        lines.setdefault(round((ch['top'] + ch['bot']) / 2 / 3.0), []).append(ch)
    out = []
    for _, recs in sorted(lines.items()):
        t, sz, col = join_line([(r['x0'], r['x1'], r['t'], r['size'], r['col']) for r in recs])
        if not t:
            continue
        out.append({'text': t, 'size': sz, 'color': col,
                    'x0': min(r['x0'] for r in recs), 'x1': max(r['x1'] for r in recs),
                    'y0': min(r['top'] for r in recs), 'y1': max(r['bot'] for r in recs)})
    return out


def color_of(box, colored):
    hits = {}
    for kind, x0, y0, x1, y1 in colored:
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if box[0] <= cx < box[2] and box[1] <= cy < box[3]:
            hits[kind] = hits.get(kind, 0) + 1
    return max(hits, key=hits.get) if hits else 'black'


# ----------------------------------------------------------------- stage: model
def from_textlayer(g):
    xs, ys, cells = g['xs'], g['ys'], g['cells']
    m = {'page': g['page'], 'w': g['w'], 'h': g['h'], 'xs': xs, 'ys': ys,
         'cells': [{'cell': c, 'text': '', 'size': 9.0, 'color': 'black'} for c in cells],
         'free': [], 'images': g['images'], 'text_layer': True}
    incell, outside = {}, []
    for ch in g['chars']:
        cx, cy = (ch['x0'] + ch['x1']) / 2, (ch['top'] + ch['bot']) / 2
        rec = (ch['x0'], ch['x1'], ch['t'], ch['size'], rgb_of(ch['color']), ch['top'], ch['bot'])
        c = find_cell(cells, xs, ys, cx, cy)
        (incell.setdefault(tuple(c), []) if c else outside).append((cy, rec))
    for k, c in enumerate(cells):
        got = incell.get(tuple(c), [])
        if not got:
            continue
        lines = {}
        for cy, rec in got:
            lines.setdefault(round(cy / 3.0), []).append(rec[:5])
        parsed = [join_line(ln) for ln in sorted(lines.values(), key=lambda r: r[0][0])]
        cols = [p[2] for p in parsed]
        m['cells'][k]['text'] = clean('\n'.join(p[0] for p in parsed if p[0]))
        m['cells'][k]['size'] = max(p[1] for p in parsed) if parsed else 9.0
        m['cells'][k]['color'] = max(set(cols), key=cols.count) if cols else 'black'
    if outside:
        m['free'] = text_lines([dict(t=r[2], x0=r[0], x1=r[1], size=r[3], col=r[4], top=r[5], bot=r[6])
                                for _, r in outside])
    return m


def vector_words(curves, maxglyph=26.0):
    """Glyph boxes from curve outlines, merged into words (pages with no text layer)."""
    bs = [[c['x0'], c['top'], c['x1'], c['bottom']] for c in curves
          if 0.4 < c['x1'] - c['x0'] <= maxglyph and 0.4 < c['bottom'] - c['top'] <= maxglyph]
    n = len(bs)
    par = list(range(n))

    def find(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a

    bs.sort()
    for i in range(n):
        for j in range(i + 1, n):
            if bs[j][0] - bs[i][2] > 0.6:
                break
            xo = min(bs[i][2], bs[j][2]) - max(bs[i][0], bs[j][0])
            yo = min(bs[i][3], bs[j][3]) - max(bs[i][1], bs[j][1])
            if xo > 0.3 * min(bs[i][2] - bs[i][0], bs[j][2] - bs[j][0]) and \
               yo > 0.3 * min(bs[i][3] - bs[i][1], bs[j][3] - bs[j][1]):
                ra, rb = find(i), find(j)
                if ra != rb:
                    par[rb] = ra
    gl = {}
    for i in range(n):
        b = gl.setdefault(find(i), list(bs[i]))
        b[0], b[1] = min(b[0], bs[i][0]), min(b[1], bs[i][1])
        b[2], b[3] = max(b[2], bs[i][2]), max(b[3], bs[i][3])
    gl = sorted(gl.values(), key=lambda b: (b[1], b[0]))
    words, used = [], [False] * len(gl)
    for i, a in enumerate(gl):
        if used[i]:
            continue
        w, used[i], changed = list(a), True, True
        while changed:
            changed = False
            for j, b in enumerate(gl):
                if used[j]:
                    continue
                h = min(w[3] - w[1], b[3] - b[1])
                if min(w[3], b[3]) - max(w[1], b[1]) < 0.5 * h:
                    continue
                if max(b[0] - w[2], w[0] - b[2]) <= max(0.45 * h, 9.0):
                    w = [min(w[0], b[0]), min(w[1], b[1]), max(w[2], b[2]), max(w[3], b[3])]
                    used[j] = changed = True
        words.append(w)
    return sorted(words, key=lambda b: (round(b[1] / 4.0), b[0]))


def _cover(items, b):
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
    return total + (cur[1] - cur[0] if cur else 0.0)


def from_ocr(g, A, B):
    """Union of full-page detection (A) and vector word boxes (B); B wins on overlap."""
    xs, ys, cells, W, H = g['xs'], g['ys'], g['cells'], g['w'], g['h']
    a = [(b['x0'] / 1000.0 * W, b['y0'] / 1000.0 * H, b['x1'] / 1000.0 * W, b['y1'] / 1000.0 * H,
          b['text'], 'A') for b in A]
    b = [(w['box'][0], w['box'][1], w['box'][2], w['box'][3], w['t'], 'B') for w in B if w['t'].strip()]
    toks = a + [t for t in b if not any(
        min(t[2], u[2]) - max(t[0], u[0]) > 0.6 * (t[2] - t[0])
        and min(t[3], u[3]) - max(t[1], u[1]) > 0.5 * (t[3] - t[1]) for u in a)]
    drop = set()
    for t in (x for x in toks if x[5] == 'B'):
        same = [u for u in toks if u[5] == 'A' and u[4].strip()
                and min(u[3], t[3]) - max(u[1], t[1]) > 0.5 * min(u[3] - u[1], t[3] - t[1])
                and min(u[2], t[2]) - max(u[0], t[0]) > 0]
        if len(same) >= 2 and _cover(same, t) >= 0.8 * (t[2] - t[0]):
            drop.update(id(s) for s in same)
    toks = [t for t in toks if id(t) not in drop]
    # font size from measured glyph height, not from a det box padded by the detector
    hs = [t[3] - t[1] for t in toks if t[4].strip() and t[5] == 'B'] or \
         [max(t[3] - t[1] - 1.2, 1.0) for t in toks if t[4].strip()]
    hs.sort()
    body = ''.join(t[4] for t in toks if t[4].strip())
    cjk = sum(1 for ch in body if '\u4e00' <= ch <= '\u9fff') >= len(body) / 2
    size = round(hs[len(hs) // 2] / (CJK_INK_RATIO if cjk else CAP_INK_RATIO), 1) if hs else 12.0
    m = {'page': g['page'], 'w': W, 'h': H, 'xs': xs, 'ys': ys, 'cells': [], 'free': [],
         'images': g['images'], 'text_layer': False, 'size': size}
    seen = {}
    for t in toks:
        c = find_cell(cells, xs, ys, (t[0] + t[2]) / 2, (t[1] + t[3]) / 2)
        seen.setdefault(c if c is None else tuple(c), []).append(t)
    for c in cells:
        got = seen.get(tuple(c), [])
        lines = {}
        for t in got:
            lines.setdefault(round(((t[1] + t[3]) / 2) / 6.0), []).append(t)
        txt = []
        for _, its in sorted(lines.items()):
            s, prev = '', None
            for t in sorted(its, key=lambda z: z[0]):
                if prev is not None and t[0] - prev > 2.2:
                    s += ' '
                s += t[4]
                prev = t[2]
            if s.strip():
                txt.append(s.strip())
        box = (xs[c[0]], ys[c[1]], xs[c[2]], ys[c[3]])
        m['cells'].append({'cell': c, 'text': clean('\n'.join(txt)), 'size': size,
                           'color': color_of(box, g.get('colored') or []), 'src': ''})
    m['free'] = [{'text': clean(t[4]), 'size': round((t[3] - t[1]) / CJK_INK_RATIO, 1),
                  'color': color_of(t[:4], g.get('colored') or []),
                  'x0': t[0], 'x1': t[2], 'y0': t[1], 'y1': t[3]}
                 for t in sorted(seen.get(None, []), key=lambda z: (round(z[1] / 6), z[0])) if t[4].strip()]
    return m


def cover(g, toks, chars=None):
    """No table rules (cover / photo wall): tile from image and text edges."""
    lines = text_lines([dict(t=c['t'], x0=c['x0'], x1=c['x1'], size=c['size'],
                             col=rgb_of(c['color']), top=c['top'], bot=c['bot']) for c in chars]) \
        if chars else \
        [{'text': clean(t[4]), 'size': round((t[3] - t[1]) / CJK_INK_RATIO, 1), 'color': 'black',
          'x0': t[0], 'x1': t[2], 'y0': t[1], 'y1': t[3]} for t in toks if t[4].strip()]
    xs = [0.0] + dedupe(cluster([v for im in g['images'] for v in (im[0], im[2])] +
                                [l['x0'] for l in lines] + [l['x1'] for l in lines], 2.0), 4.0) + [g['w']]
    ys = [0.0] + dedupe(cluster([v for im in g['images'] for v in (im[1], im[3])] +
                                [l['y0'] for l in lines] + [l['y1'] for l in lines], 2.0), 4.0) + [g['h']]
    fx = [x for x in xs[1:] if 0.6 < x < g['w'] - 0.6]
    fy = [y for y in ys[1:] if 0.6 < y < g['h'] - 0.6]
    for l in lines:                       # indices into the returned grid (XG = [0]+xs+[W])
        l['i0'] = min(max(nearest(xs, l['x0']) - 1, 0), max(len(fx) - 1, 0))
        l['i1'] = min(max(l['i0'] + 1, nearest(xs, l['x1'])), max(len(fx), 1))
        l['j0'] = min(max(nearest(ys, l['y0']) - 1, 0), max(len(fy) - 1, 0))
        l['j1'] = min(max(l['j0'] + 1, nearest(ys, l['y1'])), max(len(fy), 1))
    return {'page': g['page'], 'w': g['w'], 'h': g['h'], 'xs': fx, 'ys': fy,
            'cells': [], 'free': lines, 'images': g['images'], 'text_layer': bool(chars)}


def build_model(pdf, dirs, pages=None, verbose=True):
    """cells/g*.json (+ ocr cache if any) -> cells/m*.json."""
    doc = pdfium.PdfDocument(pdf)
    with pdfplumber.open(pdf) as fp:
        todo = pages or sorted(int(re.fullmatch(r'g(\d+)\.json', f).group(1))
                               for f in os.listdir(dirs['cells']) if re.fullmatch(r'g\d+\.json', f))
        out = []
        for i in todo:
            gp = os.path.join(dirs['cells'], 'g%d.json' % i)
            g = json.load(open(gp, encoding='utf-8'))
            g['page'] = i
            apath = os.path.join(dirs['ocr'], 'p%d.json' % i)
            bpath = os.path.join(dirs['ocr'], 'w%d.json' % i)
            if g['chars']:
                # a handful of stray rectangles is not a table - tiling from them leaves
                # merged voids that collide with the captions
                m = from_textlayer(g) if len(g['cells']) >= 20 else cover(g, _tokens_from_chars(g), g['chars'])
                note = 'text-layer'
            elif os.path.exists(apath):
                A = json.load(open(apath, encoding='utf-8'))
                B = json.load(open(bpath, encoding='utf-8')) if os.path.exists(bpath) else []
                toks = [(b['x0'] / 1000 * g['w'], b['y0'] / 1000 * g['h'],
                         b['x1'] / 1000 * g['w'], b['y1'] / 1000 * g['h'], b['text']) for b in A]
                if len(g['cells']) < 20:
                    m, note = cover(g, toks), 'no-grid(ocr A)'
                else:
                    m, note = from_ocr(g, A, B), 'outline(ocr A+B)'
            else:
                m, note = None, 'NO DATA - run: pdftoexcel.py ocr --pdf ... --pages %d' % i
            if m is None:
                print('   skip p%d: %s' % (i, note))
                continue
            json.dump(m, open(os.path.join(dirs['cells'], 'm%d.json' % i), 'w', encoding='utf-8'),
                      ensure_ascii=False)
            out.append(i)
            if verbose:
                print('p%-3d %-18s cells=%-4d text=%-4d free=%-3d imgs=%d' %
                      (i, note, len(m['cells']), sum(1 for c in m['cells'] if c['text'].strip()),
                       len(m['free']), len(m['images'])))
    return out


def _tokens_from_chars(g):
    return [(c['x0'], c['top'], c['x1'], c['bot'], c['t']) for c in g['chars']]


# ------------------------------------------------------------------ stage: build
def ocr_cache(pdf, dirs, pages, dpi=300):
    """Full-page det+rec (A) plus vector word boxes recognised on their own (B)."""
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        print('  rapidocr-onnxruntime is not installed - skipping the OCR stage (pip install -r requirements.txt)')
        return []
    ocr = RapidOCR()
    ocr.text_score = 0.0
    doc = pdfium.PdfDocument(pdf)
    done = []
    for i in pages:
        with pdfplumber.open(pdf) as fp:
            pg = fp.pages[i]
            curves = pg.curves
            g = json.load(open(os.path.join(dirs['cells'], 'g%d.json' % i), encoding='utf-8'))
        img = doc[i].render(scale=dpi / 72.0).to_pil()
        arr = np.asarray(img)
        ocr.use_text_det = True
        ocr.use_angle_cls = False
        res, _ = ocr(arr)
        W, H = img.size
        A = []
        for b, t, s in (res or []):
            bx = [p[0] for p in b]
            by = [p[1] for p in b]
            A.append({'x0': round(min(bx) / W * 1000, 2), 'y0': round(min(by) / H * 1000, 2),
                      'x1': round(max(bx) / W * 1000, 2), 'y1': round(max(by) / H * 1000, 2),
                      'text': t, 'score': float(s)})
        json.dump(A, open(os.path.join(dirs['ocr'], 'p%d.json' % i), 'w', encoding='utf-8'),
                  ensure_ascii=False)
        words = vector_words(curves)
        ocr.use_text_det = False
        # rec-only crops want 400dpi: at 300dpi a thin "/" glyph resamples into "1"
        import cv2
        k = 400 / 72.0
        rgb = np.asarray(doc[i].render(scale=k).to_pil().convert('RGB'))
        B = []
        for box in words:
            x0, y0, x1, y1 = [int(round(v * k)) for v in (box[0] - 2, box[1] - 2, box[2] + 2, box[3] + 2)]
            crop = rgb[max(y0, 0):y1, max(x0, 0):x1]
            if crop.size == 0:
                B.append({'box': [round(v, 2) for v in box], 't': '', 's': 0.0})
                continue
            if crop.shape[0] < 24:
                crop = cv2.resize(crop, (max(int(crop.shape[1] * 24.0 / crop.shape[0]), 8), 24),
                                  interpolation=cv2.INTER_CUBIC)
            res, _ = ocr(crop)
            t, s = (res[0][1], float(res[0][2])) if res else ('', 0.0)
            B.append({'box': [round(v, 2) for v in box], 't': t, 's': round(s, 3)})
        json.dump(B, open(os.path.join(dirs['ocr'], 'w%d.json' % i), 'w', encoding='utf-8'),
                  ensure_ascii=False)
        print('p%-3d A=%d boxes, B=%d words' % (i, len(A), len(B)))
        done.append(i)
    return done


def write_xlsx(pdf, dirs, out, pages=None, dpi=300, font='宋体', slash=None, verbose=True):
    """cells/m*.json -> a workbook with one sheet per page, laid out like the PDF."""
    pages = pages or sorted(int(re.fullmatch(r'm(\d+)\.json', f).group(1))
                            for f in os.listdir(dirs['cells']) if re.fullmatch(r'm\d+\.json', f))
    doc = pdfium.PdfDocument(pdf)
    scale = dpi / 72.0
    if slash is None:
        slash = not all(json.load(open(os.path.join(dirs['cells'], 'm%d.json' % p), encoding='utf-8')
                                 ).get('text_layer', True) for p in pages)
    grey = {}

    def ink(page, box):
        if page not in grey:
            for k in list(grey):
                if k != page:
                    del grey[k]
            grey[page] = np.asarray(doc[page].render(scale=INK_SCALE).to_pil().convert('L'))
        a = grey[page]
        x0, y0, x1, y1 = box
        sub = a[int((y0 + 1.2) * INK_SCALE):int((y1 - 1.2) * INK_SCALE),
                int((x0 + 1.2) * INK_SCALE):int((x1 - 1.2) * INK_SCALE)]
        if sub.size == 0:
            return None
        dark = sub < 150
        dark = dark & (dark.mean(axis=1) < 0.85)[:, None] & (dark.mean(axis=0) < 0.85)[None, :]
        rows, cols = np.where(dark.any(axis=1))[0], np.where(dark.any(axis=0))[0]
        if len(rows) == 0:
            return None
        return ((cols[-1] - cols[0] + 1) / INK_SCALE, (rows[-1] - rows[0] + 1) / INK_SCALE)

    thin = Side(style='thin', color='FF000000')
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb, used, report = Workbook(), set(), []
    wb.remove(wb.active)
    for page in pages:
        m = json.load(open(os.path.join(dirs['cells'], 'm%d.json' % page), encoding='utf-8'))
        W, H = float(m['w']), float(m['h'])
        XG, XPX = snap_grid([float(x) for x in m['xs']], W)
        YG = [0.0] + [float(y) for y in m['ys']] + [H]
        ncol, nrow = len(XG) - 1, len(YG) - 1
        ws = wb.create_sheet(sheet_name(m, page, used, font))
        ws.sheet_view.showGridLines = False
        for k in range(ncol):
            ws.column_dimensions[get_column_letter(k + 1)].width = (XPX[k] - 5) / 7
        for k in range(nrow):
            ws.row_dimensions[k + 1].height = max(YG[k + 1] - YG[k], 1.0)
        taken = [[False] * nrow for _ in range(ncol)]
        ranges = []
        for c in m['cells']:
            i0, j0, i1, j1 = c['cell']
            ranges.append((c, i0 + 2, i1 + 1, j0 + 2, j1 + 1,
                           (XG[i0 + 1], YG[j0 + 1], XG[i1 + 1], YG[j1 + 1])))
            for x in range(i0 + 2, min(i1 + 1, ncol) + 1):
                for y in range(j0 + 2, min(j1 + 1, nrow) + 1):
                    taken[x - 1][y - 1] = True

        def put(a0, a1, b0, b1, value, fnt, align, border):
            if a1 > a0 or b1 > b0:
                ws.merge_cells(start_row=b0, start_column=a0, end_row=b1, end_column=a1)
            for x in range(a0, min(a1, ncol) + 1):
                for y in range(b0, min(b1, nrow) + 1):
                    cell = ws.cell(row=y, column=x)
                    cell.font = fnt
                    if border:
                        cell.border = box_border
            head = ws.cell(row=b0, column=a0)
            head.alignment = align
            if value is not None:
                head.value = value
            return head

        n_text = n_slash = 0
        for c, a0, a1, b0, b1, gbox in ranges:
            text = c['text'].strip()
            if slash and (not text or (len(text) == 1 and text in '1lI|、,，.。/／')):
                shape = ink(page, gbox)
                # geometry beats recognition for a lone glyph: a thin slash reads as 1 / 、 / |
                if shape and SLASH_W[0] <= shape[0] <= SLASH_W[1] and SLASH_H[0] <= shape[1] <= SLASH_H[1]:
                    text, n_slash = '/', n_slash + 1
            has_img = False
            for im in m['images']:
                dx, dy = min(gbox[2], im[2]) - max(gbox[0], im[0]), min(gbox[3], im[3]) - max(gbox[1], im[1])
                if dx > 0 and dy > 0 and dx * dy > 0.2 * max(1.0, (im[2] - im[0]) * (im[3] - im[1])):
                    has_img = True
                    break
            fnt = Font(name=font, size=round(c['size'] or 12.0, 1), color=COLOR.get(c['color'], 'FF000000'))
            align = Alignment(horizontal='center', vertical='bottom' if (has_img and b1 > b0) else 'center',
                              wrap_text=True)
            if text and re.fullmatch(r'\d+(\.\d+)?', text):
                dp = len(text.split('.')[1]) if '.' in text else 0
                head = put(a0, a1, b0, b1, int(text) if dp == 0 else float(text), fnt, align, True)
                head.number_format = '0' if dp == 0 else '0.' + '0' * dp
            else:
                put(a0, a1, b0, b1, text or None, fnt, align, True)
            n_text += 1 if text else 0

        for f in m['free']:
            if not f['text'].strip():
                continue
            if page == 0 or 'i0' in f:
                a0, a1, b0, b1 = f.get('i0', 0) + 2, f.get('i1', 0) + 1, f.get('j0', 0) + 2, f.get('j1', 0) + 1
            else:
                (a0, a1), (b0, b1) = span(f['x0'], f['x1'], XG, ncol), span(f['y0'], f['y1'], YG, nrow)
                a0, a1, b0, b1 = a0 + 1, a1 + 1, b0 + 1, b1 + 1
            if any(taken[x - 1][y - 1] for x in range(a0, min(a1, ncol) + 1) for y in range(b0, min(b1, nrow) + 1)):
                a1, b1 = a0, b0
            fnt = Font(name=font, size=round(f['size'] or 12.0, 1), color=COLOR.get(f['color'], 'FF000000'))
            put(min(a0, ncol), min(a1, ncol), min(b0, nrow), min(b1, nrow), f['text'].strip(), fnt,
                Alignment(horizontal='center', vertical='center', wrap_text=True), False)

        pageimg = doc[page].render(scale=scale).to_pil()
        n_img = 0
        for n, im in enumerate(m['images']):
            x0, y0, x1, y1 = im
            img = pageimg.crop(tuple(int(v * scale) for v in (x0, y0, x1, y1)))
            if img.size[0] < 2 or img.size[1] < 2:
                continue
            path = os.path.join(dirs['imgs'], 'p%d_%d.png' % (page, n))
            img.save(path)
            a0, b0 = span(x0, x0 + 1, XG, ncol)[0], span(y0, y0 + 1, YG, nrow)[0]
            pic = XImage(path)
            pic.anchor = OneCellAnchor(
                _from=AnchorMarker(col=a0, colOff=int((x0 - XG[a0]) * EMU_PT),
                                   row=b0, rowOff=int((y0 - YG[b0]) * EMU_PT)),
                ext=XDRPositiveSize2D(int((x1 - x0) * EMU_PT), int((y1 - y0) * EMU_PT)))
            ws.add_image(pic)
            n_img += 1
        del pageimg

        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = 'portrait'
        ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.2, bottom=0.2, header=0.05, footer=0.05)
        ws.print_area = 'A1:%s%d' % (get_column_letter(ncol), nrow)
        if verbose:
            report.append((ws.title, page, ncol, nrow, len(ranges), n_text, n_slash, n_img))
    wb.save(out)
    for r in report:
        print('sheet %-30s p%-3d %dx%-4d cells=%-4d text=%-4d slash=%d imgs=%d' % r)
    print('saved %s (%d bytes)' % (out, os.path.getsize(out)))
    return wb


def tidy(s):
    s = s.replace('\n', ' ').replace('■', ' ')
    s = re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def sheet_name(m, page, used, font=None):
    """Page title: it usually lives inside a merged header cell, not in margin text."""
    if page == 0:
        base = '封面'
    else:
        cand = [(f['size'], tidy(f['text'])) for f in m['free']]
        cand += [(c['size'], tidy(c['text'])) for c in m['cells'] if c['text'].strip()]
        head = [(s, t) for s, t in cand
                if s >= 14 and len(t) >= (8 if m['cells'] else 4) and re.search(r'[\u4e00-\u9fff]', t)
                and not re.fullmatch(r'[0-9.元条/件%（）()\s自]*', t)]
        base = max(head, key=lambda st: (st[0], len(st[1])))[1] if head else ''
        base = base[:26]
    base = re.sub(r'[\\*?:\[\]]', '', base).replace('/', '·')[:28] or '第%d页' % (page + 1)
    name, k = base, 0
    while name in used:
        k += 1
        name = '%s 第%d页' % (base[:24], page + 1) if k == 1 else '%s%d' % (base[:22], k)
    used.add(name)
    return name[:31]
