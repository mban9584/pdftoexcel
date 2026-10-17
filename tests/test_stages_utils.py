import importlib.util
import json
import sys
import tempfile
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import stages


def _stub_pdfium():
    """Replace pypdfium2 with a page that renders to a real white PIL image.

    write_xlsx only touches the PDF to crop photos out of the rendered page, so a stub
    page (size + render) is enough to exercise the whole sheet-geometry path offline.
    """
    from PIL import Image

    class _Img:
        def __init__(self, w, h):
            self._im = Image.new("RGB", (max(1, int(w)), max(1, int(h))), "white")

        def to_pil(self):
            return self._im

    class _Page:
        def __init__(self, w, h):
            self._w, self._h = w, h

        def render(self, scale=1.0):
            return _Img(self._w * scale, self._h * scale)

    class _Doc:
        def __init__(self, path):
            pass

        def __getitem__(self, i):
            return _Page(600, 800)

    fake = sys.modules.get("pypdfium2") or types.ModuleType("pypdfium2")
    fake.PdfDocument = _Doc
    sys.modules["pypdfium2"] = fake
    return fake


class StageUtilityTests(unittest.TestCase):
    def test_snap_grid_snaps_boundaries_not_widths(self):
        points, pixels = stages.snap_grid([10.2, 20.6], 30.0)
        self.assertEqual(points[0], 0.0)
        self.assertEqual(points[-1], 30.0)
        self.assertEqual(len(pixels), len(points) - 1)
        self.assertTrue(all(px >= stages.MIN_COL_PX for px in pixels))

    def test_sheet_name_uses_large_chinese_heading_and_dedupes(self):
        used = set()
        model = {
            "cells": [{"text": "2026 产品报价单", "size": 16}],
            "free": [],
        }
        self.assertEqual(stages.sheet_name(model, 1, used), "2026 产品报价单")
        self.assertEqual(stages.sheet_name(model, 1, used), "2026 产品报价单 第2页")

    def test_tidy_removes_cjk_inner_spaces(self):
        self.assertEqual(stages.tidy("产 品  报 价"), "产品报价")

    def test_pick_orientation_wide_grid_is_landscape(self):
        wide = {"xs": [0, 300, 600], "ys": [0, 200, 400], "w": 600, "h": 400}
        tall = {"xs": [0, 200, 400], "ys": [0, 300, 600], "w": 400, "h": 600}
        self.assertEqual(stages.pick_orientation(wide), "landscape")
        self.assertEqual(stages.pick_orientation(tall), "portrait")

    def test_shift_rows_moves_values_and_row_heights(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1).value = "top band"
        ws.cell(row=2, column=1).value = "header"
        ws.row_dimensions[2].height = 30.0
        stages.shift_rows(ws, 1)
        self.assertEqual(ws.cell(row=1, column=1).value, "header")
        self.assertEqual(ws.cell(row=2, column=1).value, None)
        self.assertEqual(ws.row_dimensions[1].height, 30.0)


class WriteXlsxTests(unittest.TestCase):
    def setUp(self):
        self._real = sys.modules.get("pypdfium2")
        self._had = "pypdfium2" in sys.modules
        _stub_pdfium()
        importlib.reload(stages)

    def tearDown(self):
        if self._had:
            sys.modules["pypdfium2"] = self._real
        else:
            sys.modules.pop("pypdfium2", None)
        importlib.reload(stages)

    def _model(self):
        return {"page": 0, "w": 600.0, "h": 800.0, "xs": [100.0, 300.0, 500.0],
                "ys": [200.0, 400.0, 600.0],
                "cells": [{"cell": [0, 0, 1, 1], "text": "品名称", "size": 12, "color": "black"},
                          {"cell": [1, 0, 2, 1], "text": "12.50", "size": 12, "color": "black"},
                          {"cell": [0, 1, 1, 2], "text": "规格", "size": 12, "color": "black"}],
                "colored": [], "images": [], "chars": [], "free": [], "text_layer": True}

    def _run(self, d, **kw):
        out = str(Path(d) / "out.xlsx")
        stages.write_xlsx("x.pdf", {"cells": d, "imgs": d}, out, pages=[0], verbose=False, **kw)
        from openpyxl import load_workbook
        return load_workbook(out).worksheets[0]

    def test_top_band_default_and_no_top_band(self):
        with tempfile.TemporaryDirectory() as d:
            json.dump(self._model(), open(Path(d) / "m0.json", "w", encoding="utf-8"),
                      ensure_ascii=False)
            default = self._run(d)
            self.assertEqual(default.cell(row=1, column=1).value, None)   # the 0->ys[0] band
            self.assertEqual(default.cell(row=2, column=2).value, "品名称")  # col +2: 左边界带占 A 列
            nb = self._run(d, no_top_band=True)
            self.assertEqual(nb.cell(row=1, column=2).value, "品名称")     # header promoted

    def test_landscape_auto_and_forced(self):
        with tempfile.TemporaryDirectory() as d:
            json.dump(self._model(), open(Path(d) / "m0.json", "w", encoding="utf-8"),
                      ensure_ascii=False)
            self.assertEqual(self._run(d).page_setup.orientation, "portrait")   # 400<600 tall
            self.assertEqual(self._run(d, landscape="on").page_setup.orientation, "landscape")
            self.assertEqual(self._run(d, landscape="off").page_setup.orientation, "portrait")

    def test_sheet_names_and_pattern_override(self):
        with tempfile.TemporaryDirectory() as d:
            json.dump(self._model(), open(Path(d) / "m0.json", "w", encoding="utf-8"),
                      ensure_ascii=False)
            self.assertEqual(self._run(d, sheet_names=["第1页"]).title, "第1页")
            self.assertEqual(self._run(d, sheet_pattern="第{page}页").title, "第1页")


if __name__ == "__main__":
    unittest.main()
