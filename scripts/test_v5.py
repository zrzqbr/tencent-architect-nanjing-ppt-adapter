#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
南京模板适配器基础测试套件。

运行：
  python scripts/test_v5.py
"""

import sys
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import apply_template as T
import brand_palette as P


class TestNanjingPalette(unittest.TestCase):
    def test_allowed_palette(self):
        allowed = {
            "#3272DC", "#08194B", "#FFFFFF", "#00C8D8",
            "#01A4FF", "#44474F", "#8B8C8C",
        }
        actual = {f"#{r:02X}{g:02X}{b:02X}" for r, g, b, _ in T.BRAND_PALETTE}
        self.assertEqual(actual, allowed)

    def test_navy_is_allowed(self):
        self.assertFalse(T._is_forbidden_color(0x08, 0x19, 0x4B))
        self.assertTrue(P.validate_color("#08194B")["safe"])

    def test_old_red_is_forbidden(self):
        self.assertTrue(T._is_forbidden_color(0xD8, 0x0C, 0x01))
        self.assertFalse(P.validate_color("#D80C01")["safe"])

    def test_old_warm_colors_are_not_allowed(self):
        for color in ["#FF0000", "#FAF6EE", "#D4A574", "#FAD16A", "#F19D19"]:
            self.assertFalse(P.validate_color(color)["safe"], color)


class TestAssets(unittest.TestCase):
    def test_nanjing_assets_exist(self):
        assets = T.get_assets("nanjing")
        for key in ["bg_cover", "bg_section", "bg_content", "bg_end"]:
            self.assertTrue(Path(assets[key]).exists(), key)
        self.assertFalse(assets["logo_enabled"])
        self.assertIsNone(assets["logo_main"])

    def test_no_old_logo_assets(self):
        self.assertFalse((T.ASSETS / "logos" / "logo-main.png").exists())
        self.assertFalse((T.ASSETS / "logos" / "logo-corner.png").exists())


class TestColorMapping(unittest.TestCase):
    def test_black_maps_to_allowed_palette_when_not_skipped(self):
        nearest = T.find_nearest_brand_color(0, 0, 0, skip_white_black=False)
        self.assertIsNotNone(nearest)
        hex_color = f"#{nearest[0]:02X}{nearest[1]:02X}{nearest[2]:02X}"
        self.assertIn(hex_color, P.ALLOWED_LIST)

    def test_red_maps_to_allowed_palette(self):
        nearest = T.find_nearest_brand_color(0xD8, 0x0C, 0x01, skip_white_black=False)
        self.assertIsNotNone(nearest)
        hex_color = f"#{nearest[0]:02X}{nearest[1]:02X}{nearest[2]:02X}"
        self.assertIn(hex_color, P.ALLOWED_LIST)


class TestCLIVersion(unittest.TestCase):
    def test_description_mentions_nanjing(self):
        import subprocess
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "apply_template.py"), "--help"],
            capture_output=True, text=True, timeout=10
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn("南京", result.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
