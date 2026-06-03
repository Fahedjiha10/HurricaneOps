from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from typing import Any, Callable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import urlopen


CENSUS_GEOCODER_URL = "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress"
SOURCE_NAME = "US Census Geocoder"


@dataclass
class MunicipalityLookupResult:
    query_address: str
    status: str
    display_name: str
    municipality: str
    municipality_type: str
    county: str
    state: str
    matched_address: str
    latitude: float | None
    longitude: float | None
    source: str
    warning: str

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def _empty_result(address: str, status: str, warning: str) -> MunicipalityLookupResult:
    return MunicipalityLookupResult(
        query_address=address,
        status=status,
        display_name="",
        municipality="",
        municipality_type="",
        county="",
        state="",
        matched_address="",
        latitude=None,
        longitude=None,
        source=SOURCE_NAME,
        warning=warning,
    )


def _fetch_json(url: str, timeout: int) -> dict[str, Any]:
    with urlopen(url, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _first_geography(
    geographies: dict[str, Any], *names: str
) -> dict[str, Any] | None:
    for name in names:
        candidates = geographies.get(name)
        if isinstance(candidates, list) and candidates:
            first = candidates[0]
            if isinstance(first, dict):
                return first
    return None


def _geography_label(geography: dict[str, Any] | None) -> str:
    if not geography:
        return ""
    basename = str(geography.get("BASENAME") or "").strip()
    lsad = str(geography.get("LSAD") or "").strip()
    name = str(geography.get("NAME") or "").strip()
    if basename and lsad and lsad.lower() not in basename.lower():
        return f"{basename} {lsad}".strip()
    return name or basename


def _state_from_match(match: dict[str, Any], state_geography: dict[str, Any] | None) -> str:
    state_name = _geography_label(state_geography)
    if state_name:
        return state_name
    matched_address = str(match.get("matchedAddress") or "")
    state_match = re.search(r",\s*([A-Z]{2})\s*,?\s*\d{5}(?:-\d{4})?\s*$", matched_address)
    return state_match.group(1) if state_match else ""


def _display_name(municipality: str, county: str, state: str) -> str:
    parts = [part for part in (municipality, county, state) if part]
    return ", ".join(parts)


def lookup_municipality(
    address: str,
    timeout: int = 8,
    fetch_json: Callable[[str], dict[str, Any]] | None = None,
) -> MunicipalityLookupResult:
    query_address = address.strip()
    if not query_address:
        return _empty_result(
            query_address,
            "NOT_PROVIDED",
            "No project address was provided, so municipality lookup was skipped.",
        )

    url = (
        f"{CENSUS_GEOCODER_URL}?"
        + urlencode(
            {
                "address": query_address,
                "benchmark": "Public_AR_Current",
                "vintage": "Current_Current",
                "layers": "all",
                "format": "json",
            }
        )
    )
    try:
        payload = fetch_json(url) if fetch_json else _fetch_json(url, timeout)
    except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
        return _empty_result(
            query_address,
            "LOOKUP_FAILED",
            f"Municipality lookup failed: {exc}",
        )

    matches = payload.get("result", {}).get("addressMatches", [])
    if not matches:
        return _empty_result(
            query_address,
            "NO_MATCH",
            "No geocoder match was returned for this address. Review the address spelling.",
        )

    match = matches[0]
    geographies = match.get("geographies", {})
    coordinates = match.get("coordinates", {})
    county = _geography_label(_first_geography(geographies, "Counties"))
    state = _state_from_match(match, _first_geography(geographies, "States"))
    incorporated_place = _first_geography(
        geographies,
        "Incorporated Places",
        "Places",
    )
    cdp = _first_geography(geographies, "Census Designated Places")
    county_subdivision = _first_geography(geographies, "County Subdivisions")

    warning = ""
    status = "FOUND"
    municipality_type = "incorporated municipality"
    municipality = _geography_label(incorporated_place)
    if not municipality:
        status = "UNINCORPORATED_OR_UNKNOWN"
        municipality_type = "unincorporated / census geography"
        municipality = _geography_label(cdp) or _geography_label(county_subdivision)
        warning = (
            "No incorporated municipality was returned for this address. "
            "Treat the permit jurisdiction as unconfirmed until reviewed."
        )

    display_name = _display_name(municipality, county, state)
    if not display_name:
        display_name = str(match.get("matchedAddress") or query_address)
        warning = warning or "A match was found, but no municipality/county label was returned."

    return MunicipalityLookupResult(
        query_address=query_address,
        status=status,
        display_name=display_name,
        municipality=municipality,
        municipality_type=municipality_type,
        county=county,
        state=state,
        matched_address=str(match.get("matchedAddress") or ""),
        latitude=coordinates.get("y") if isinstance(coordinates, dict) else None,
        longitude=coordinates.get("x") if isinstance(coordinates, dict) else None,
        source=SOURCE_NAME,
        warning=warning,
    )
