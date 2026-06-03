from __future__ import annotations

import re
from typing import Any


HUMAN_REVIEW_THRESHOLD = 0.85
EXPECTED_DOOR_TYPES = ("SWING", "POCKET", "BARN", "FIXED", "SLIDING", "OVERHEAD", "ROLL")


def validate_glazing_row(row: dict[str, Any]) -> list[str]:
    warnings: list[str] = []
    tag = str(row.get("tag") or "").strip()
    description = str(row.get("description") or "").upper()
    remarks = str(row.get("remarks") or "").upper()
    noa = str(row.get("noa") or "").strip()
    if not tag.upper().startswith("G"):
        warnings.append("Glazing tag should usually start with G.")
    if int(row.get("count") or 0) <= 0:
        warnings.append("Glazing count must be positive.")
    if row.get("width_inches") is None:
        warnings.append("Glazing width could not be parsed.")
    if row.get("height_inches") is None:
        warnings.append("Glazing height could not be parsed.")
    if noa and not noa.upper().startswith("FL"):
        warnings.append("Glazing NOA should usually start with FL.")
    if "DOOR" in description and row.get("height_inches") is not None:
        if float(row["height_inches"]) <= 72:
            warnings.append("Glazing door height should usually exceed 72 inches.")
    if "FIXED WINDOW" in description and ("ENTRY" in remarks or "HARDWARE" in remarks):
        warnings.append("Fixed window should not include entry-hardware remarks.")
    return warnings


def validate_door_row(row: dict[str, Any]) -> list[str]:
    warnings: list[str] = []
    door_number = str(row.get("door_number") or "").strip()
    door_type = str(row.get("type") or "").upper()
    quantity = int(row.get("quantity") or 0)
    if re.fullmatch(r"LEVEL\s*\d+", door_number, flags=re.IGNORECASE):
        warnings.append("Level header must not become a door item row.")
    elif not door_number.isdigit():
        warnings.append("Door number should usually be numeric.")
    if quantity <= 0:
        warnings.append("Door quantity must be positive.")
    elif quantity != 1:
        warnings.append("Door quantity should usually be 1.")
    if row.get("width_inches") is None:
        warnings.append("Door width could not be parsed.")
    if row.get("height_inches") is None:
        warnings.append("Door height could not be parsed.")
    if door_type and not any(expected in door_type for expected in EXPECTED_DOOR_TYPES):
        warnings.append("Door type is outside the expected type vocabulary.")
    if not door_type:
        warnings.append("Door type is missing.")
    return warnings


def confidence_score(row: dict[str, Any], warnings: list[str], identifier_key: str) -> float:
    score = 1.0
    if not str(row.get(identifier_key) or "").strip():
        score -= 0.25
    if int(row.get("count") or row.get("quantity") or 0) <= 0:
        score -= 0.16
    if row.get("width_inches") is None:
        score -= 0.18
    if row.get("height_inches") is None:
        score -= 0.18
    score -= min(len(warnings) * 0.04, 0.24)
    return round(max(score, 0.0), 2)


def add_human_review_warning(confidence: float, warnings: list[str]) -> list[str]:
    output = list(warnings)
    if confidence < HUMAN_REVIEW_THRESHOLD:
        output.append("HUMAN REVIEW REQUIRED: confidence below 0.85.")
    return output
