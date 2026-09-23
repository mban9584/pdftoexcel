# pipeline/ - frozen reference, not the tool

These 11 scripts are the **as-run originals** from two real jobs on one machine
(2026-09-22/23: a 5-page price list and a 25-page product catalogue). They are kept so the
techniques can be read, not executed elsewhere:

- every file has that machine's absolute paths hard-coded at the top
  `~/Desktop/...` and `~/pdftoexcel-work/...`;
- page ranges and one output name are baked in (`build_xlsx.py` writes out.xlsx);
- the numbers they printed (389/360/313/281 cells, 6 无货 cells on p3, 12pt body) belong to
  **those two files** - never reuse them as assumptions for a new one.

Everything reusable now lives one level up:

| need | use |
|---|---|
| geometry + route decision | `../pdftoexcel.py probe --pdf 输入.pdf` |
| text from a text layer, or from OCR | `../pdftoexcel.py model` / `ocr` |
| the workbook | `../pdftoexcel.py xlsx --pdf 输入.pdf` |
| verification | `../pdftoexcel.py check --pdf 输入.pdf` |
| font-size calibration | `../pdftoexcel.py fontsize` |

Worth reading here anyway, because they hold the reference implementations of two tricky
steps: `words.py` (curve outlines -> glyph -> word boxes, mirrored in `stages.vector_words`)
and `assign.py` (the A/B recognition union, mirrored in `stages.from_ocr`).
