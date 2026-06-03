from __future__ import annotations

import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "05_SCRIPTS"))

from validators import confidence_score, validate_door_row, validate_glazing_row


class GlazingValidatorTests(unittest.TestCase):
    def test_valid_swing_door_glazing_row(self) -> None:
        row = {
            "tag": "G01",
            "description": "SWING DOOR",
            "count": 1,
            "width_inches": 72.0,
            "height_inches": 108.0,
            "remarks": "ENTRY, LATCH & DEADBOLT HARDWARE",
            "noa": "FL 26942",
        }
        warnings = validate_glazing_row(row)
        self.assertEqual(warnings, [])
        self.assertEqual(confidence_score(row, warnings, "tag"), 1.0)

    def test_fixed_window_entry_hardware_is_flagged(self) -> None:
        row = {
            "tag": "G03",
            "description": "FIXED WINDOW",
            "count": 1,
            "width_inches": 84.0,
            "height_inches": 62.0,
            "remarks": "ENTRY HARDWARE",
            "noa": "FL 40676.1",
        }
        self.assertIn(
            "Fixed window should not include entry-hardware remarks.",
            validate_glazing_row(row),
        )


class DoorValidatorTests(unittest.TestCase):
    def test_valid_numeric_door_row(self) -> None:
        row = {
            "door_number": "101",
            "quantity": 1,
            "width_inches": 32.0,
            "height_inches": 108.0,
            "type": "SWING",
        }
        self.assertEqual(validate_door_row(row), [])

    def test_level_header_is_flagged(self) -> None:
        row = {
            "door_number": "LEVEL 1",
            "quantity": 1,
            "width_inches": None,
            "height_inches": None,
            "type": "",
        }
        self.assertIn(
            "Level header must not become a door item row.",
            validate_door_row(row),
        )
