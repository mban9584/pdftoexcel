"""Cross-machine path resolution - the only place a path is decided.

Priority for the working directory: explicit --work argument, then $PDFTOEXCEL_WORK,
then ~/.qwen/tmp/<pdf file name>. Nothing in this skill needs editing to move machines.
"""
import os
import re

SUBDIRS = ('cells', 'ocr', 'pages', 'imgs')


def expand(p):
    return os.path.abspath(os.path.expanduser(os.path.expandvars(p))) if p else None


def home_tmp():
    return os.path.join(os.path.expanduser('~'), '.qwen', 'tmp')


def slug(pdf):
    # split on both separators: os.path.basename only knows the host's, so a Windows path
    # handed to a Linux/macOS run (or the CI) would keep its folders in the name
    stem = os.path.splitext(re.split(r'[\\/]', pdf)[-1])[0]
    stem = re.sub(r'[^\w.\u4e00-\u9fff-]+', '_', stem).strip('_')
    return stem[:48] or 'task'


def workdir(pdf, work=None):
    w = expand(work) if work else expand(os.environ.get('PDFTOEXCEL_WORK'))
    return w or os.path.join(home_tmp(), slug(pdf))


def sub(root, name):
    path = os.path.join(root, name)
    os.makedirs(path, exist_ok=True)
    return path


def desktop(pdf=None, out=None, ext='xlsx'):
    """Output path: explicit --out, else <Desktop>/<pdf name>.<ext> (Desktop falls back to home)."""
    if out:
        p = expand(out)
        return p if p.lower().endswith('.' + ext) else os.path.join(p, os.path.splitext(os.path.basename(pdf))[0] + '.' + ext)
    for cand in (os.path.join(os.path.expanduser('~'), 'Desktop'),
                 os.path.join(os.path.expanduser('~'), '桌面'),
                 os.path.expanduser('~')):
        if os.path.isdir(cand):
            return os.path.join(cand, os.path.splitext(os.path.basename(pdf))[0] + '.' + ext)
    return os.path.join(os.getcwd(), os.path.splitext(os.path.basename(pdf))[0] + '.' + ext)


def ensure(pdf, work=None):
    root = workdir(pdf, work)
    os.makedirs(root, exist_ok=True)
    return {n: sub(root, n) for n in SUBDIRS} | {'root': root}
