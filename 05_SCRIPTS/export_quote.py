from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Any

from schemas import ExtractionResult
from validators import HUMAN_REVIEW_THRESHOLD


UNCERTAIN_HEADERS = (
    "schedule_type",
    "source_page",
    "identifier",
    "confidence",
    "warnings",
    "width_raw",
    "height_raw",
)


def write_extraction_outputs(result: ExtractionResult, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    schedules_path = output_dir / "extracted_schedules.json"
    audit_path = output_dir / "extraction_audit.json"
    uncertain_path = output_dir / "uncertain_schedule_rows.csv"
    schedules_path.write_text(json.dumps(result.to_dict(), indent=2) + "\n", encoding="utf-8")
    audit_path.write_text(json.dumps(result.audit, indent=2) + "\n", encoding="utf-8")
    uncertain_rows: list[dict[str, Any]] = []
    for schedule in result.schedules:
        for item in schedule.items:
            if item.confidence >= HUMAN_REVIEW_THRESHOLD:
                continue
            item_dict = item.to_dict()
            uncertain_rows.append(
                {
                    "schedule_type": schedule.schedule_type,
                    "source_page": schedule.source_page,
                    "identifier": item_dict.get("tag") or item_dict.get("door_number"),
                    "confidence": item.confidence,
                    "warnings": " | ".join(item.warnings),
                    "width_raw": item_dict.get("width_raw", ""),
                    "height_raw": item_dict.get("height_raw", ""),
                }
            )
    with uncertain_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=UNCERTAIN_HEADERS)
        writer.writeheader()
        writer.writerows(uncertain_rows)
    return {
        "extracted_schedules.json": schedules_path,
        "extraction_audit.json": audit_path,
        "uncertain_schedule_rows.csv": uncertain_path,
    }


def quote_rows_from_structured(result: ExtractionResult) -> dict[str, list[dict[str, str]]]:
    """Map only validated structured rows into the quote workbook adapter shape."""
    windows: list[dict[str, str]] = []
    storefronts: list[dict[str, str]] = []
    doors: list[dict[str, str]] = []
    for schedule in result.schedules:
        for item in schedule.items:
            if item.confidence < HUMAN_REVIEW_THRESHOLD:
                continue
            values = item.to_dict()
            if schedule.schedule_type == "glazing":
                description = str(values["description"])
                panel_match = re.search(r"\b(\d+)\s*[- ]?\s*PANEL\b", description.upper())
                brand_product = str(values.get("brand_product") or "")
                common = {
                    "mark": str(values["tag"]),
                    "quantity": str(values["count"]),
                    "width": str(values["width_raw"]),
                    "height": str(values["height_raw"]),
                    "type": description,
                    "material": str(values.get("material") or ""),
                    "glass_type": str(values.get("glass_type") or brand_product),
                    "finish": str(values.get("finish") or ""),
                    "noa": str(values.get("noa") or ""),
                    "brand_product": brand_product,
                    "level": str(values.get("level") or ""),
                    "remarks": str(values.get("remarks") or ""),
                    "panels": panel_match.group(1) if panel_match else "",
                    "source_schedule": "GLAZING SCHEDULE",
                    "extraction_status": "STRUCTURED VALIDATED - glazing schedule",
                    "source": "extracted_schedules.json",
                }
                # Keep mixed glazing schedules together. Splitting their window and
                # door rows across quote tabs makes drawing-order review unreliable.
                storefronts.append(common)
                if "WINDOW" in description.upper():
                    windows.append(common.copy())
            elif schedule.schedule_type == "door":
                material_parts = [
                    part.strip()
                    for part in str(values.get("material") or "").split(",")
                    if part.strip()
                ]
                door_material = (
                    material_parts[0] if material_parts else str(values.get("material") or "")
                )
                frame_finish = (
                    material_parts[1]
                    if len(material_parts) > 1
                    else str(values.get("frame_finish") or "")
                )
                noa = str(values.get("noa") or "")
                remarks = str(values.get("remarks") or "")
                if noa and remarks.upper() == noa.upper():
                    remarks = ""
                doors.append(
                    {
                        "door_no": str(values["door_number"]),
                        "location": str(values["location"]),
                        "type": str(values.get("type") or ""),
                        "width": str(values["width_raw"]),
                        "height": str(values["height_raw"]),
                        "thickness": str(values.get("thickness") or ""),
                        "door_material": door_material,
                        "door_finish": "",
                        "frame_material": str(values.get("jamb") or ""),
                        "frame_finish": frame_finish,
                        "fire_rating": "",
                        "noa": noa,
                        "panic_hardware": str(values.get("hardware") or ""),
                        "remarks": remarks,
                        "level": str(values.get("level") or ""),
                        "quantity": str(values["quantity"]),
                        "panels": str(values.get("panels") or ""),
                        "fixed_panels": str(values.get("fixed_panels") or ""),
                        "scope": "garage door schedule"
                        if str(values["door_number"]).upper().startswith("G")
                        else "door schedule",
                        "extraction_status": "STRUCTURED VALIDATED - door schedule",
                        "raw_text": "",
                        "source": "extracted_schedules.json",
                    }
                )
    # Preserve drawing order so estimators can reconcile workbook columns
    # directly against the architectural schedule.
    return {"windows": windows, "storefronts": storefronts, "doors": doors}
