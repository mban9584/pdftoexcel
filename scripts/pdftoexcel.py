#!/usr/bin/env python
"""Layout-preserving PDF -> XLSX diagnostics.

  probe    : decide the route for each page (text layer? how are the rules drawn?
             is find_tables trustworthy here? recover the real cell grid)
  fontsize : measure the ink height of a region and convert it to a point size
  check    : audit a produced workbook (merge overlaps, geometry drift vs the PDF grid,
             paper/fit settings, images)

Nothing here writes the xlsx - see pipeline/ for the stage scripts and SKILL.md for order.
Paths may contain spaces: quote them.
"""
import argparse
import json
import os
import re
import sys

import cv2
import numpy as np
import pdfplumber
import pypdfium2 as pdfium

RULE_H_MIN_LEN, RULE_TOL = 5.0, 3.0     # a rule segment: long in one axis, thin in the other


def rules(page):
    """Horizontal / vertical rule segments, taken from page.edges.

    Table rules come in three drawings and a file often mixes them: thin filled rects
    (one per cell edge, CorelDRAW style), real stroked lines, and multi-point stroked
    paths that trace a whole table. `page.edges` is the superset - every line plus every
    side of every rect/curve - so it catches all three; scanning rects or lines alone
    silently returns zero cells on the other styles.

    Edges lying on the page frame are dropped: keeping them turns the outer blank margin
    into one giant "cell", which shows up as (cells N, overlaps N-1).
    """
    h, v = [], []
    W, H = page.width, page.height
    for e in page.edges:
        if e.get('object_type') == 'filter':
            continue
        w, hh = e['x1'] - e['x0'], e['bottom'] - e['top']
        if e['orientation'] == 'h' and w > RULE_H_MIN_LEN and hh <= RULE_TOL:
            y = (e['top'] + e['bottom']) / 2.0
            if y > 1.5 and y < H - 1.5:
                h.append((y, e['x0'], e['x1']))
        elif e['orientation'] == 'v' and hh > RULE_H_MIN_LEN and w <= RULE_TOL:
            x = (e['x0'] + e['x1']) / 2.0
            if x > 1.5 and x < W - 1.5:
                v.append((x, e['top'], e['bottom']))
    return h, v


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


def detect_cells(page, scale=4.0, pad=6, min_side=10):
    """Cell rectangles = enclosed white regions of the drawn rules.

    Flood-filling is robust to rules that are split per cell edge, to merged cells and to
    photo frames; `page.find_tables()` is not (it invents or merges tables on such files).
    """
    hh, vv = rules(page)
    if not hh and not vv:
        return [], [], []
    H = int(round(page.height * scale)) + 2 * pad
    W = int(round(page.width * scale)) + 2 * pad
    img = np.zeros((H, W), np.uint8)
    for y, a, b in hh:
        Y = int(round(y * scale)) + pad
        cv2.line(img, (int(a * scale) + pad, Y), (int(b * scale) + pad, Y), 255, 3)
    for x, a, b in vv:
        X = int(round(x * scale)) + pad
        cv2.line(img, (X, int(a * scale) + pad), (X, int(b * scale) + pad), 255, 3)
    n, _, stats, _ = cv2.connectedComponentsWithStats(cv2.bitwise_not(img), 8)
    raw = []
    for i in range(1, n):
        x, y, w, h, _ = stats[i]
        if w < min_side or h < min_side:
            continue
        if x <= 1 or y <= 1 or x + w >= W - 1 or y + h >= H - 1:
            continue
        raw.append(((x - pad) / scale, (y - pad) / scale, (x + w - pad) / scale, (y + h - pad) / scale))
    if not raw:
        return [], [], []
    xs = dedupe(cluster([c[0] for c in raw] + [c[2] for c in raw]))
    ys = dedupe(cluster([c[1] for c in raw] + [c[3] for c in raw]))
    cells = set()
    for x0, y0, x1, y1 in raw:
        i0, i1, j0, j1 = nearest(xs, x0), nearest(xs, x1), nearest(ys, y0), nearest(ys, y1)
        if i1 <= i0:
            i1 = i0 + 1
        if j1 <= j0:
            j1 = j0 + 1
        if i1 < len(xs) and j1 < len(ys):
            cells.add((i0, j0, i1, j1))
    return sorted(cells, key=lambda c: (c[1], c[0])), xs, ys


def overlaps(cells):
    bad = []
    for a in range(len(cells)):
        for b in range(a + 1, len(cells)):
            i0, j0, i1, j1 = cells[a]
            k0, l0, k1, l1 = cells[b]
            if i0 < k1 and k0 < i1 and j0 < l1 and l0 < j1:
                bad.append((cells[a], cells[b]))
    return bad


def _rgb(c):
    if c is None or isinstance(c, (int, float)):
        return None
    if len(c) == 1:
        return (1 - c[0],) * 3
    if len(c) == 3:
        return tuple(c)
    if len(c) == 4:
        cy, m, y0, k = c
        return ((1 - cy) * (1 - k), (1 - m) * (1 - k), (1 - y0) * (1 - k))
    return None


def colored_glyphs(page):
    """Where non-black text is drawn, so red/blue cells survive without a text layer."""
    boxes = []
    for o in page.curves:
        w, h = o['x1'] - o['x0'], o['bottom'] - o['top']
        if w > 25 or h > 25:
            continue
        rgb = _rgb(o.get('non_stroking_color'))
        if not rgb:
            continue
        r, g, b = rgb
        kind = 'red' if (r > 0.45 and g < 0.45 and b < 0.45) else ('blue' if (b > 0.45 and r < 0.45) else None)
        if kind:
            boxes.append([kind, round(o['x0'], 1), round(o['top'], 1), round(o['x1'], 1), round(o['bottom'], 1)])
    return boxes


def cmd_probe(a):
    import _paths
    root = a.work or _paths.workdir(a.pdf, a.out)
    out = root
    os.makedirs(out, exist_ok=True)
    os.makedirs(os.path.join(out, 'cells'), exist_ok=True)
    pdf = pdfium.PdfDocument(a.pdf)
    with pdfplumber.open(a.pdf) as fp:
        todo = range(len(fp.pages)) if not a.pages else [int(x) for x in a.pages.split(',')]
        for i in todo:
            pg = fp.pages[i]
            hh, vv = rules(pg)
            cells, xs, ys = detect_cells(pg)
            try:
                ntab = len(pg.find_tables())
            except Exception as exc:
                ntab = 'err:%s' % type(exc).__name__
            text = (pg.extract_text() or '').strip()
            g = {'page': i, 'w': round(pg.width, 2), 'h': round(pg.height, 2),
                 'xs': [round(v, 2) for v in xs], 'ys': [round(v, 2) for v in ys],
                 'cells': [list(c) for c in cells], 'colored': colored_glyphs(pg),
                 'images': [[round(o[k], 2) for k in ('x0', 'top', 'x1', 'bottom')] for o in pg.images
                            if (o['x1'] - o['x0']) * (o['bottom'] - o['top']) < 0.8 * pg.width * pg.height],
                 'chars': [{'t': c['text'], 'x0': c['x0'], 'x1': c['x1'], 'top': c['top'],
                            'bot': c['bottom'], 'size': round(c.get('size', 9), 2),
                            'color': (list(c['non_stroking_color'])
                                      if isinstance(c.get('non_stroking_color'), (list, tuple))
                                      else [c.get('non_stroking_color')])} for c in pg.chars]}
            path = os.path.join(out, 'cells', 'g%d.json' % i)
            json.dump(g, open(path, 'w', encoding='utf-8'), ensure_ascii=False)
            k = 150 / 72.0
            arr = np.array(pdf[i].render(scale=k).to_pil().convert('RGB'))
            for c in cells:
                cv2.rectangle(arr, (int(xs[c[0]] * k), int(ys[c[1]] * k)),
                              (int(xs[c[2]] * k), int(ys[c[3]] * k)), (0, 160, 255), 2)
            cv2.imwrite(os.path.join(out, 'cells', 'v%d.png' % i), arr[:, :, ::-1])
            sizes = sorted({round(c['size'], 1) for c in g['chars']})
            print('p%-2d %6.1fx%6.1fpt  rules h=%-5d v=%-5d  cells=%-4d (grid %dx%d, overlap %d)  imgs=%d'
                  % (i, pg.width, pg.height, len(hh), len(vv), len(cells), len(xs), len(ys),
                     len(overlaps(cells)), len(g['images'])))
            print('    text layer: %d chars, sizes %s, extract_text %d chars | curves=%d rects=%d lines=%d'
                  % (len(g['chars']), sizes[:6], len(text), len(pg.curves), len(pg.rects), len(pg.lines)))
            n_ov = len(overlaps(cells))
            print('    find_tables=%s (cross-check only - never take the grid from it)' % ntab)
            route = ('text-layer: take text/size/colour from chars, skip OCR and skip '
                     'the font calibration' if g['chars'] else
                     'outline-ocr: no text layer, render + both recognition paths'
                     if len(pg.curves) > 500 else
                     'image-only: no text and no outlines, place images and re-render text')
            if not cells:
                route = ('no-grid: %s' % ('cover/photo page - tile from image and text edges'
                                          if g['images'] or len(pg.curves) < 500 else
                                          'RULES NOT DETECTED - check how the lines are drawn'))
            warn = ''
            if cells and n_ov >= len(cells) - 1:
                warn = '\n    !! overlaps ~= cells-1: the page frame was taken for a rule (see rules())'
            elif n_ov:
                warn = '\n    !! %d overlapping cells - Excel merges cannot overlap; fix the grid' % n_ov
            print('    route -> %s%s' % (route, warn))
    print('\nnext: eyeball %s/cells/v*.png (blue boxes must sit on the real rules), then SKILL.md stage 2'
          % out.replace('\\', '\\\\'))


def cmd_fontsize(a):
    x0, y0, x1, y1 = [float(v) for v in a.rect.split(',')]
    doc = pdfium.PdfDocument(a.pdf)
    S = 6.0
    g = np.asarray(doc[a.page].render(scale=S).to_pil().convert('L'))
    sub = g[int((y0 + 1.2) * S):int((y1 - 1.2) * S), int((x0 + 1.2) * S):int((x1 - 1.2) * S)]
    dark = sub < 150
    dark = dark & (dark.mean(axis=1) < 0.85)[:, None] & (dark.mean(axis=0) < 0.85)[None, :]
    rows = np.where(dark.any(axis=1))[0]
    if not len(rows):
        print('no ink in that rect')
        return
    ink = (rows[-1] - rows[0] + 1) / S
    print('ink height %.2fpt. Pick the ratio matching the script in that rect:' % ink)
    print('  CJK glyphs  -> %.1fpt  (ink/em 0.917)' % (ink / 0.917))
    print('  latin/digit -> %.1fpt  (ink/em 0.681, cap height only)' % (ink / 0.681))
    print('An OCR detection box is ~1pt taller than the ink; dividing a box height by 0.72 overstates')
    print('the size by 25-35%% - calibrate against a page that does have a text layer.')


def cmd_check(a):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    wb = load_workbook(a.xlsx)
    grids = {}
    if a.grid:
        for f in sorted(os.listdir(a.grid)):
            mt = re.fullmatch(r'([mg])(\d+)\.json', f)
            if not mt:
                continue
            n = int(mt.group(2))
            if n in grids and grids[n]['stem'] == 'm':
                continue                     # the model wins: it is what the build consumed
            d = json.load(open(os.path.join(a.grid, f), encoding='utf-8'))
            grids[n] = {'stem': mt.group(1), 'w': d['w'], 'h': d['h'], 'xs': d['xs'], 'ys': d['ys'],
                        'cells': d.get('cells'), 'images': d.get('images') or []}

    def shown(cell):
        """What Excel will actually display - never compare str(value), `72` + '0.00' shows 72.00."""
        v = cell.value
        if v is None:
            return ''
        if isinstance(v, (int, float)):
            m = re.fullmatch(r'0(?:\.(0+))?', cell.number_format or 'General')
            if m and m.group(1):
                return ('%.' + str(len(m.group(1))) + 'f') % v
            if m:
                return '%d' % v
            return ('%d' % v) if float(v) == int(v) else ('%g' % v)
        return str(v)
    ok = True
    for n, ws in enumerate(wb.worksheets):
        mr = list(ws.merged_cells.ranges)
        clash = [(str(p), str(q)) for i, p in enumerate(mr) for q in mr[i + 1:]
                 if not (p.max_col < q.min_col or q.max_col < p.min_col
                         or p.max_row < q.min_row or q.max_row < p.min_row)]
        maxc = max((r.max_col for r in mr), default=ws.max_column)
        line = '%-24s merges=%-4d cols=%-3d rows=%-3d imgs=%-3d overlap=%d' % (
            ws.title, len(mr), maxc, max((r.max_row for r in mr), default=ws.max_row),
            len(ws._images), len(clash))
        g = grids.get(n)
        if g:
            orig = [0.0] + [float(v) for v in g['xs']] + [float(g['w'])]
            real, run = [], 0
            for k in range(1, len(orig)):
                dim = ws.column_dimensions.get(get_column_letter(k))
                run += round((dim.width if dim else 8.43) * 7 + 5)
                real.append(run * 0.75)
            drift = [abs(r - v) for r, v in zip(real, orig[1:])]
            line += '  col edges off PDF by max %.2fpt mean %.2fpt' % (max(drift), sum(drift) / len(drift))
            if max(drift) > 1.5:
                ok = False
                line += '  <-- drifts (boundaries must be snapped to whole px, see SKILL.md stage 4)'
        if g and g.get('cells') and g['stem'] == 'm':
            xs, ys = [0.0] + [float(v) for v in g['xs']] + [float(g['w'])], \
                     [0.0] + [float(v) for v in g['ys']] + [float(g['h'])]
            miss = inferred = 0
            for c in g['cells']:
                i0, j0, i1, j1 = c['cell']
                want = re.sub(r'\s+', ' ', c['text']).strip()
                got = re.sub(r'\s+', ' ', shown(ws.cell(row=j0 + 2, column=i0 + 2))).strip()
                if got == want:
                    continue
                if not want and got == '/':
                    inferred += 1             # 无货 slash, added from ink at build time
                    continue
                bx = (xs[i0 + 1], ys[j0 + 1], xs[i1 + 1], ys[j1 + 1])
                if not want and any(min(bx[2], im[2]) - max(bx[0], im[0]) > 2
                                    and min(bx[3], im[3]) - max(bx[1], im[1]) > 2 for im in g['images']):
                    continue                  # photo carries it
                miss += 1
                if miss <= 5:
                    print('    MISMATCH %s%d model=%r sheet=%r' %
                          (get_column_letter(i0 + 2), j0 + 2, want[:40], got[:40]))
            if miss:
                ok = False
                line += '  content differs in %d cells' % miss
            else:
                line += '  content ok (%d cells%s)' % (
                    len(g['cells']), ', %d 无货 inferred' % inferred if inferred else '')
        print(line)
        for c in clash[:5]:
            ok = False
            print('    OVERLAPPING MERGES %s %s  (Excel will offer to repair)' % c)
        print('    paper=%s orient=%s fitW=%s fitH=%s fitToPage=%s margins=%s gridlines=%s' % (
            ws.page_setup.paperSize, ws.page_setup.orientation, ws.page_setup.fitToWidth,
            ws.page_setup.fitToHeight, getattr(ws.sheet_properties.pageSetUpPr, 'fitToPage', None),
            [round(getattr(ws.page_margins, s), 2) for s in ('left', 'right', 'top', 'bottom')],
            ws.sheet_view.showGridLines))
    print('OK' if ok else 'PROBLEMS FOUND')


def cmd_model(a):
    import _paths
    import stages
    d = _paths.ensure(a.pdf, a.work)
    pages = [int(x) for x in a.pages.split(',')] if a.pages else None
    print('work dir: %s' % d['root'])
    stages.build_model(a.pdf, d, pages)


def cmd_ocr(a):
    import _paths
    import stages
    d = _paths.ensure(a.pdf, a.work)
    pages = [int(x) for x in a.pages.split(',')] if a.pages else []
    print('work dir: %s' % d['root'])
    stages.ocr_cache(a.pdf, d, pages, a.dpi)


def cmd_xlsx(a):
    import _paths
    import stages
    d = _paths.ensure(a.pdf, a.work)
    pages = [int(x) for x in a.pages.split(',')] if a.pages else None
    out = _paths.desktop(a.pdf, a.out)
    print('work dir: %s\noutput  : %s' % (d['root'], out))
    stages.write_xlsx(a.pdf, d, out, pages, dpi=a.dpi, font=a.font,
                      slash=a.slash)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest='cmd', required=True)

    def common(name, func, xlsx=False):
        p = sub.add_parser(name)
        p.add_argument('--pdf', required=not xlsx, help='source PDF (paths work on Windows/macOS/Linux)')
        p.add_argument('--work', help='work directory; default ~/.qwen/tmp/<pdf name>')
        p.add_argument('--pages', help='comma list of 0-based pages, default all')
        p.set_defaults(func=func)
        return p

    p1 = common('probe', cmd_probe)
    p1.add_argument('--out', help='legacy alias for --work')
    common('model', cmd_model)
    p4 = common('ocr', cmd_ocr)
    p4.add_argument('--dpi', type=int, default=300)
    p2 = common('xlsx', cmd_xlsx)
    p2.add_argument('--out', help='output .xlsx; default ~/Desktop/<pdf name>.xlsx')
    p2.add_argument('--dpi', type=int, default=300, help='photo crop resolution')
    p2.add_argument('--font', default='宋体')
    slash = p2.add_mutually_exclusive_group()
    slash.add_argument('--auto-slash', dest='slash', action='store_const', const=None, default=None,
                       help='infer "/" from ink only for pages without a text layer (default)')
    slash.add_argument('--slash', dest='slash', action='store_true',
                       help='force 无货 "/" inference from ink')
    slash.add_argument('--no-slash', dest='slash', action='store_false',
                       help='disable 无货 "/" inference')
    p3 = common('check', cmd_check, xlsx=True)
    p3.add_argument('--xlsx')
    p3.add_argument('--grid', help='directory with g*/m*.json; default the --pdf work dir')
    p5 = sub.add_parser('fontsize')
    p5.add_argument('--pdf', required=True)
    p5.add_argument('--page', type=int, required=True)
    p5.add_argument('--rect', required=True, help='x0,y0,x1,y1 in pt')
    p5.set_defaults(func=cmd_fontsize)
    a = ap.parse_args()
    if getattr(a, 'grid', None) is None and getattr(a, 'pdf', None):
        import _paths
        a.grid = _paths.ensure(a.pdf, a.work)['cells']
    if a.cmd == 'check' and not getattr(a, 'xlsx', None):
        import _paths
        a.xlsx = _paths.desktop(a.pdf)
    a.func(a)


if __name__ == '__main__':
    sys.exit(main())

