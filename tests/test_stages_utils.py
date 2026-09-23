import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import stages


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


if __name__ == "__main__":
    unittest.main()
