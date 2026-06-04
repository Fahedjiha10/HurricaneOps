from __future__ import annotations

import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "05_SCRIPTS"))

from job_intake import _sheet_compartments


class JobIntakeClassificationTests(unittest.TestCase):
    def test_storefront_window_schedule_is_not_lost_when_title_reads_elevations(self) -> None:
        compartments = _sheet_compartments(
            Path("6300 Block_Combined Set.pdf"),
            "W-2",
            "ELEVATIONS",
            "",
            "\n".join(
                [
                    "ALUM. STOREFRONT / WINDOW SYSTEM",
                    "WINDOW SCHEDULE",
                    "A-8.02 WINDOW SCHEDULE & DETAILS",
                    "MARK OPERATION MATERIAL WINDOW SIZE REMARKS",
                ]
            ),
        )
        self.assertIn("window_schedules", compartments)
        self.assertIn("storefront_schedules", compartments)


if __name__ == "__main__":
    unittest.main()
