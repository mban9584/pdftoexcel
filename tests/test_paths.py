import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import _paths


class PathTests(unittest.TestCase):
    def test_slug_keeps_cjk_and_safe_ascii(self):
        self.assertEqual(_paths.slug(r"C:\in\报价单 A B?.pdf"), "报价单_A_B")

    def test_workdir_prefers_explicit_path(self):
        with tempfile.TemporaryDirectory() as d:
            got = _paths.workdir("x.pdf", d)
            self.assertEqual(got, os.path.abspath(d))

    def test_desktop_accepts_output_directory(self):
        with tempfile.TemporaryDirectory() as d:
            got = _paths.desktop("price list.pdf", d)
            self.assertEqual(got, os.path.join(os.path.abspath(d), "price list.xlsx"))


if __name__ == "__main__":
    unittest.main()

