from __future__ import annotations

import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "05_SCRIPTS"))

from municipality_lookup import lookup_municipality


class MunicipalityLookupTests(unittest.TestCase):
    def test_incorporated_place_lookup(self) -> None:
        def fake_fetch(url: str):
            self.assertIn("address=160+NW+44+St", url)
            return {
                "result": {
                    "addressMatches": [
                        {
                            "matchedAddress": "160 NW 44TH ST, MIAMI, FL, 33127",
                            "coordinates": {"x": -80.2, "y": 25.8},
                            "geographies": {
                                "Incorporated Places": [
                                    {"BASENAME": "Miami", "LSAD": "city", "NAME": "Miami city"}
                                ],
                                "Counties": [{"NAME": "Miami-Dade County"}],
                            },
                        }
                    ]
                }
            }

        result = lookup_municipality("160 NW 44 St", fetch_json=fake_fetch)

        self.assertEqual(result.status, "FOUND")
        self.assertEqual(result.municipality, "Miami city")
        self.assertEqual(result.county, "Miami-Dade County")
        self.assertEqual(result.state, "FL")
        self.assertEqual(result.display_name, "Miami city, Miami-Dade County, FL")
        self.assertEqual(result.latitude, 25.8)
        self.assertEqual(result.longitude, -80.2)

    def test_unincorporated_lookup_flags_review(self) -> None:
        result = lookup_municipality(
            "123 Example Rd, Miami FL",
            fetch_json=lambda _url: {
                "result": {
                    "addressMatches": [
                        {
                            "matchedAddress": "123 EXAMPLE RD, MIAMI, FL, 33156",
                            "coordinates": {},
                            "geographies": {
                                "County Subdivisions": [{"NAME": "Kendall CCD"}],
                                "Counties": [{"NAME": "Miami-Dade County"}],
                            },
                        }
                    ]
                }
            },
        )

        self.assertEqual(result.status, "UNINCORPORATED_OR_UNKNOWN")
        self.assertEqual(result.municipality, "Kendall CCD")
        self.assertIn("permit jurisdiction as unconfirmed", result.warning)

    def test_no_address_skips_lookup(self) -> None:
        result = lookup_municipality("  ")

        self.assertEqual(result.status, "NOT_PROVIDED")
        self.assertEqual(result.display_name, "")


if __name__ == "__main__":
    unittest.main()
