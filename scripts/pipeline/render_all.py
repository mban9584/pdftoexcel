import pypdfium2 as pdfium
import os

PDF_PATH = os.environ.get('PDFTOEXCEL_PDF', os.path.expanduser('~/Desktop/sample.pdf'))
WORK = os.environ.get('PDFTOEXCEL_WORK', os.path.expanduser('~/pdftoexcel-work'))
OUT_PATH = os.path.join(os.path.expanduser('~/Desktop'), 'out.xlsx')

PDF = PDF_PATH
OUT = os.path.join(WORK, 'pages')
os.makedirs(OUT, exist_ok=True)
doc = pdfium.PdfDocument(PDF)
for i in range(len(doc)):
    page = doc[i]
    bmp = page.render(scale=200 / 72.0)
    img = bmp.to_pil()
    img.save(os.path.join(OUT, f'p{i}.png'))
    print('p%d.png' % i, img.size)
print('--- OCR modules ---')
for m in ('pytesseract', 'paddleocr', 'cnocr', 'easyocr', 'rapidocr_onnxruntime', 'rapidocr'):
    try:
        __import__(m)
        print(m, 'AVAILABLE')
    except Exception:
        print(m, 'no')
