from __future__ import annotations

import re
from fractions import Fraction


_FEET_INCHES_RE = re.compile(
    r"""
    ^\s*
    (?P<feet>\d+)
    \s*(?:'\s*-?|ft\s*-?|feet\s*-?|-)\s*
    (?P<inches>\d+)
    (?:\s+(?P<fraction>\d+\s*/\s*\d+))?
    \s*(?:"|in|inches)?\s*
    $
    """,
    flags=re.IGNORECASE | re.VERBOSE,
)


def _normalize_quotes(value: str) -> str:
    return (
        value.replace("\u2032", "'")
        .replace("\u2019", "'")
        .replace("\u2033", '"')
        .replace("\u201d", '"')
        .replace("\u2212", "-")
        .strip()
    )


def parse_architectural_dimension(value: str | None) -> float | None:
    """Return an architectural feet-and-inches dimension as total inches."""
    if value is None:
        return None
    normalized = _normalize_quotes(str(value))
    match = _FEET_INCHES_RE.fullmatch(normalized)
    if not match:
        return None
    feet = int(match.group("feet"))
    inches = int(match.group("inches"))
    if inches >= 12:
        return None
    fraction_text = match.group("fraction")
    fraction = float(Fraction(fraction_text.replace(" ", ""))) if fraction_text else 0.0
    return float(feet * 12 + inches + fraction)


def area_square_feet(width_inches: float | None, height_inches: float | None) -> float | None:
    if width_inches is None or height_inches is None:
        return None
    return round(width_inches * height_inches / 144.0, 3)
