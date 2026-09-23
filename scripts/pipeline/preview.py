import json
import os
import sys

DIR = os.path.join(WORK, 'cells')
i = int(sys.argv[1])
src = sys.argv[2] if len(sys.argv) > 2 else 'm'
d = json.load(open(os.path.join(DIR, '%s%d.json' % (src, i)), encoding='utf-8'))
if src == 'o':
    rows = defaultdict = {}
    for r in d:
        i0, j0, i1, j1 = r['cell']
        t = '\n'.join(it['t'] for it in sorted(r['items'], key=lambda x: (round(x['y'] / 5), x['x'])))
        rows.setdefault(j0, []).append(((i0, i1, j0, j1), t))
    for j0 in sorted(rows):
        parts = []
        for (i0, i1, _, _), t in sorted(rows[j0], key=lambda p: p[0][0]):
            if t.strip():
                parts.append('%s%s%s' % (('[' + t.replace('\n', '|') + ']') if i1 > i0 + 1 else t.replace('\n', '|'),
                                         '' if i1 == i0 + 1 else '', ''))
        if parts:
            print('r%-3d %s' % (j0, ' '.join(parts)))
    sys.exit()

print('FREE:', [(f['text'], round(f['y0'])) for f in d['free'] if f['text'].strip()])
byrow = {}
for c in d['cells']:
    i0, j0, i1, j1 = c['cell']
    if c['text'].strip():
        byrow.setdefault(j0, []).append((i0, i1, c['text']))
for j0 in sorted(byrow):
    line = []
    for i0, i1, t in sorted(byrow[j0]):
        span = 'x%d-%d' % (i0, i1) if i1 > i0 + 1 else 'x%d' % i0
        line.append('%s{%s}' % (t.replace('\n', '|'), span))
    print('r%-3d %s' % (j0, '  '.join(line)))
