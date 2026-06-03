from __future__ import annotations

import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "05_SCRIPTS"))

from dimension_parser import parse_architectural_dimension


class DimensionParserTests(unittest.TestCase):
    def test_requested_dimension_formats(self) -> None:
        expectations = {
            "2'-8\"": 32.0,
            "2' - 8\"": 32.0,
            "6'-0\"": 72.0,
            "15'-6\"": 186.0,
            "4'-0 11/16\"": 48.6875,
            "3'-8 1/16\"": 44.0625,
            "2 - 8": 32.0,
            "9 - 0": 108.0,
        }
        for raw_value, expected_inches in expectations.items():
            with self.subTest(raw_value=raw_value):
                self.assertEqual(parse_architectural_dimension(raw_value), expected_inches)

    def test_rejects_unusable_values(self) -> None:
        for raw_value in ("", "TBD", "2'-14\"", None):
            with self.subTest(raw_value=raw_value):
                self.assertIsNone(parse_architectural_dimension(raw_value))
