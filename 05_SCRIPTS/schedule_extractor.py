from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Iterable

import fitz
import pdfplumber

from dimension_parser import area_square_feet, parse_architectural_dimension
from schemas import DoorItem, DoorSchedule, ExtractionResult, GlazingItem, GlazingSchedule
from validators import (
    add_human_review_warning,
    confidence_score,
    validate_door_row,
    validate_glazing_row,
)


SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg"}
WEAK_TEXT_CHARACTER_THRESHOLD = 80
TITLE_PATTERNS = {
    "garage_door": re.compile(r"\bGARAGE\s+DOOR\s+SCHEDULE\b", flags=re.IGNORECASE),
    "storefront": re.compile(r"\bSTOREFRONT\s+SCHEDULE\b", flags=re.IGNORECASE),
    "glazing": re.compile(r"\bGLAZING\s+SCHEDULE\b", flags=re.IGNORECASE),
    "window": re.compile(r"\bWINDOW\s+SCHEDULE\b", flags=re.IGNORECASE),
    "door": re.compile(r"\bDOOR\s+SCHEDULE\b", flags=re.IGNORECASE),
}
GLAZING_TAG_RE = re.compile(r"^[A-Z]+\d+[A-Z]?$", flags=re.IGNORECASE)
DOOR_NUMBER_RE = re.compile(r"^\d{2,5}$")
LEVEL_RE = re.compile(r"\bLEVEL\s+\d+\b", flags=re.IGNORECASE)
DIMENSION_TEXT_RE = r"""\d+\s*'\s*-\s*\d+(?:\s+\d+\s*/\s*\d+)?\s*" """


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _int_or_default(value: Any, default: int = 0) -> int:
    match = re.search(r"\d+", _clean(value))
    return int(match.group()) if match else default


def _optional_int(value: Any) -> int | None:
    cleaned = _clean(value)
    return int(cleaned) if cleaned.isdigit() else None


def _repair_split_fraction(value: str) -> str:
    """Repair PDF table cells that split 11/16 into `11"` and `16` lines."""
    cleaned = _clean(value)
    return re.sub(r'(\d+)"\s+(\d+)$', r'\1/\2"', cleaned)


class ScheduleExtractor:
    """Extract schedule tables into validated JSON-ready schema objects."""

    def __init__(
        self,
        project_address: str | None = None,
        render_scale: float = 3.0,
    ) -> None:
        self.project_address = project_address or None
        self.render_scale = render_scale
        self.audit: list[dict[str, Any]] = []

    def extract_files(self, paths: Iterable[Path | str]) -> ExtractionResult:
        result = ExtractionResult()
        for path in paths:
            result.extend(self.extract_file(Path(path)))
        result.schedules = self._dedupe_schedules(result.schedules)
        result.audit = list(self.audit)
        return result

    def extract_file(self, path: Path | str) -> ExtractionResult:
        input_path = Path(path)
        if input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported input extension {input_path.suffix!r}. "
                f"Expected one of: {', '.join(sorted(SUPPORTED_EXTENSIONS))}."
            )
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        self._audit("input_started", input_path, method="dispatch")
        if input_path.suffix.lower() == ".pdf":
            result = self._extract_pdf(input_path)
        else:
            result = self._extract_image(input_path, source_page=None, rendered_from_pdf=False)
        result.schedules = self._dedupe_schedules(result.schedules)
        result.audit = list(self.audit)
        return result

    def _extract_pdf(self, path: Path) -> ExtractionResult:
        result = ExtractionResult()
        with pdfplumber.open(path) as pdf, fitz.open(path) as fitz_document:
            for page_index, page in enumerate(pdf.pages):
                source_page = page_index + 1
                text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                tables = page.extract_tables() or []
                titles = self._detect_schedule_titles(text + "\n" + path.stem)
                self._audit(
                    "pdf_text_table_attempt",
                    path,
                    source_page=source_page,
                    method="pdfplumber",
                    text_characters=len(text.strip()),
                    detected_titles=titles,
                    table_count=len(tables),
                )
                page_result = self._extract_table_rows(
                    tables, text, path, source_page, method="pdfplumber"
                )
                result.extend(page_result)
                parsed_items = sum(len(schedule.items) for schedule in page_result.schedules)
                weak_text = len(text.strip()) < WEAK_TEXT_CHARACTER_THRESHOLD
                if parsed_items == 0 and (weak_text or titles):
                    reason = "weak_or_empty_pdf_text" if weak_text else "recognized_schedule_without_rows"
                    self._audit(
                        "image_fallback_requested",
                        path,
                        source_page=source_page,
                        method="pymupdf_page_get_pixmap",
                        reason=reason,
                    )
                    with tempfile.TemporaryDirectory(prefix="hurricaneops-render-") as temp_dir:
                        image_path = Path(temp_dir) / f"page-{source_page:03d}.png"
                        matrix = fitz.Matrix(self.render_scale, self.render_scale)
                        pixmap = fitz_document[page_index].get_pixmap(matrix=matrix, alpha=False)
                        pixmap.save(image_path)
                        result.extend(
                            self._extract_image(
                                image_path,
                                source_page=source_page,
                                rendered_from_pdf=True,
                                source_file=path,
                            )
                        )
        return result

    def _extract_table_rows(
        self,
        tables: list[list[list[Any]]],
        page_text: str,
        source_file: Path,
        source_page: int | None,
        method: str,
    ) -> ExtractionResult:
        glazing_items: list[GlazingItem] = []
        door_items: list[DoorItem] = []
        for table_index, table in enumerate(tables):
            level: str | None = None
            for row_index, raw_row in enumerate(table):
                values = [_clean(value) for value in raw_row]
                row_text = " ".join(value for value in values if value)
                level_match = LEVEL_RE.search(row_text)
                if level_match and len([value for value in values if value]) <= 2:
                    level = level_match.group().upper()
                    continue
                glazing_item = self._glazing_item_from_cells(values, level)
                if glazing_item:
                    glazing_items.append(glazing_item)
                    self._audit(
                        "structured_row",
                        source_file,
                        source_page=source_page,
                        method=method,
                        table_index=table_index,
                        row_index=row_index,
                        schedule_type="glazing",
                        raw_cells=values,
                        confidence=glazing_item.confidence,
                    )
                    continue
                door_item = self._door_item_from_cells(values, level)
                if door_item:
                    door_items.append(door_item)
                    self._audit(
                        "structured_row",
                        source_file,
                        source_page=source_page,
                        method=method,
                        table_index=table_index,
                        row_index=row_index,
                        schedule_type="door",
                        raw_cells=values,
                        confidence=door_item.confidence,
                    )
        for item in self._garage_door_items_from_text(page_text):
            door_items.append(item)
            self._audit(
                "structured_row",
                source_file,
                source_page=source_page,
                method=f"{method}_garage_text_region",
                schedule_type="door",
                raw_cells=[item.door_number, item.location, item.width_raw, item.height_raw],
                confidence=item.confidence,
            )
        existing_tags = {item.tag.upper() for item in glazing_items}
        for item in self._glazing_items_from_text(page_text):
            if item.tag.upper() in existing_tags:
                continue
            glazing_items.append(item)
            existing_tags.add(item.tag.upper())
            self._audit(
                "structured_row",
                source_file,
                source_page=source_page,
                method=f"{method}_glazing_text_region_supplement",
                schedule_type="glazing",
                raw_cells=[
                    item.tag,
                    item.description,
                    item.width_raw,
                    item.height_raw,
                ],
                confidence=item.confidence,
            )
        schedules = []
        glazing_items = self._dedupe_items(glazing_items, "tag")
        door_items = self._dedupe_items(door_items, "door_number")
        if glazing_items:
            schedules.append(
                GlazingSchedule(
                    project_address=self.project_address,
                    source_page=source_page,
                    items=glazing_items,
                )
            )
        if door_items:
            schedules.append(
                DoorSchedule(
                    project_address=self.project_address,
                    source_page=source_page,
                    items=door_items,
                )
            )
        return ExtractionResult(schedules=schedules)

    def _glazing_item_from_cells(
        self, values: list[str], level: str | None
    ) -> GlazingItem | None:
        if len(values) < 5 or not GLAZING_TAG_RE.fullmatch(values[0]):
            return None
        if not any(
            opening_word in values[1].upper()
            for opening_word in ("DOOR", "WINDOW", "STOREFRONT", "GLAZING")
        ):
            return None
        count = _int_or_default(values[2])
        width_raw = _repair_split_fraction(values[3])
        height_raw = _repair_split_fraction(values[4])
        width_inches = parse_architectural_dimension(width_raw)
        height_inches = parse_architectural_dimension(height_raw)
        row = {
            "level": level,
            "tag": values[0],
            "description": values[1],
            "count": count,
            "width_raw": width_raw,
            "height_raw": height_raw,
            "width_inches": width_inches,
            "height_inches": height_inches,
            "area_sf": area_square_feet(width_inches, height_inches),
            "remarks": values[5] if len(values) > 5 and values[5] else None,
            "noa": values[6] if len(values) > 6 and values[6] else None,
            "brand_product": values[7] if len(values) > 7 and values[7] else None,
        }
        warnings = validate_glazing_row(row)
        confidence = confidence_score(row, warnings, "tag")
        row["warnings"] = add_human_review_warning(confidence, warnings)
        row["confidence"] = confidence
        return GlazingItem(**row)

    def _glazing_items_from_text(self, text: str) -> list[GlazingItem]:
        dimension = r"""\d+'\s*-\s*\d+(?:\s+\d+/\d+)?\""""
        pattern = re.compile(
            rf"\b(G\d+[A-Z]?)\s+(.+?)\s+(\d+)\s+({dimension})\s+"
            rf"({dimension})(?:\s+(.*))?$",
            flags=re.IGNORECASE,
        )
        items: list[GlazingItem] = []
        level: str | None = None
        for line in text.splitlines():
            cleaned = _clean(line)
            if LEVEL_RE.fullmatch(cleaned):
                level = cleaned.upper()
                continue
            match = pattern.search(cleaned)
            if not match:
                continue
            tag, description, quantity, width, height, remainder = match.groups()
            noa_match = re.search(r"\bFL\s+\d+(?:\.\d+)?\b", remainder or "")
            remarks = ""
            noa = ""
            brand_product = ""
            if noa_match:
                remarks = (remainder or "")[: noa_match.start()].strip()
                noa = noa_match.group()
                brand_product = (remainder or "")[noa_match.end() :].strip()
            else:
                remarks = (remainder or "").strip()
            item = self._glazing_item_from_cells(
                [tag, description, quantity, width, height, remarks, noa, brand_product],
                level,
            )
            if item:
                items.append(item)
        return items

    def _door_item_from_cells(
        self, values: list[str], level: str | None
    ) -> DoorItem | None:
        if len(values) < 9 or not DOOR_NUMBER_RE.fullmatch(values[0]):
            return None
        row = {
            "level": level,
            "door_number": values[0],
            "location": values[1],
            "quantity": _int_or_default(values[2]),
            "width_raw": values[3],
            "height_raw": values[4],
            "width_inches": parse_architectural_dimension(values[3]),
            "height_inches": parse_architectural_dimension(values[4]),
            "panels": _optional_int(values[5]) if len(values) > 5 else None,
            "fixed_panels": _optional_int(values[6]) if len(values) > 6 else None,
            "jamb": values[7] if len(values) > 7 and values[7] else None,
            "type": values[8] if len(values) > 8 and values[8] else None,
            "material": values[9] if len(values) > 9 and values[9] else None,
            "hardware": values[11] if len(values) > 11 and values[11] else None,
            "remarks": values[12] if len(values) > 12 and values[12] else None,
        }
        warnings = validate_door_row(row)
        confidence = confidence_score(row, warnings, "door_number")
        row["warnings"] = add_human_review_warning(confidence, warnings)
        row["confidence"] = confidence
        return DoorItem(**row)

    def _garage_door_items_from_text(self, text: str) -> list[DoorItem]:
        if not TITLE_PATTERNS["garage_door"].search(text):
            return []
        garage_section = TITLE_PATTERNS["garage_door"].split(text, maxsplit=1)[-1]
        pattern = re.compile(
            rf"(?m)^(G\d+)\s+(.+?)\s+(\d+)\s+({DIMENSION_TEXT_RE})\s+"
            rf"({DIMENSION_TEXT_RE})\s+(\d+)\s+([A-Z]+)\s+"
            r"(OVERHEAD|SWING|SLIDING|ROLL(?:ING)?)\s+(.+)$",
            flags=re.IGNORECASE | re.VERBOSE,
        )
        items: list[DoorItem] = []
        for match in pattern.finditer(garage_section):
            door_number, location, quantity, width, height, panels, jamb, door_type, rest = (
                _clean(value) for value in match.groups()
            )
            material = None
            hardware = None
            remarks = rest or None
            detail_match = re.fullmatch(
                r"(?P<material>[A-Z][A-Z /-]*?)\s+"
                r"(?P<hardware>BY\s+MANUF\.?)\s+"
                r"(?P<remarks>.+)",
                rest,
                flags=re.IGNORECASE,
            )
            if detail_match:
                material = _clean(detail_match.group("material")) or None
                hardware = _clean(detail_match.group("hardware")) or None
                remarks = _clean(detail_match.group("remarks")) or None
            row = {
                "level": None,
                "door_number": door_number,
                "location": location,
                "quantity": int(quantity),
                "width_raw": width,
                "height_raw": height,
                "width_inches": parse_architectural_dimension(width),
                "height_inches": parse_architectural_dimension(height),
                "panels": _optional_int(panels),
                "fixed_panels": None,
                "jamb": jamb or None,
                "type": door_type or None,
                "material": material,
                "hardware": hardware,
                "remarks": remarks,
            }
            warnings = validate_door_row(row)
            confidence = confidence_score(row, warnings, "door_number")
            row["warnings"] = add_human_review_warning(confidence, warnings)
            row["confidence"] = confidence
            items.append(DoorItem(**row))
        return items

    def _extract_image(
        self,
        image_path: Path,
        source_page: int | None,
        rendered_from_pdf: bool,
        source_file: Path | None = None,
    ) -> ExtractionResult:
        audit_file = source_file or image_path
        dependencies = self._image_dependencies(audit_file, source_page)
        if dependencies is None:
            return ExtractionResult()
        cv2, np, pytesseract = dependencies
        image_bytes = np.fromfile(str(image_path), dtype=np.uint8)
        image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
        if image is None:
            self._audit(
                "image_fallback_failed",
                audit_file,
                source_page=source_page,
                method="opencv",
                warning="Image could not be decoded.",
            )
            return ExtractionResult()
        processed = self._preprocess_image(image, cv2)
        title_regions = self._detect_image_schedule_regions(processed, pytesseract)
        self._audit(
            "image_title_detection",
            audit_file,
            source_page=source_page,
            method="cell_ocr",
            rendered_from_pdf=rendered_from_pdf,
            detected_regions=[region["title"] for region in title_regions],
        )
        if not title_regions:
            self._audit(
                "image_fallback_warning",
                audit_file,
                source_page=source_page,
                method="cell_ocr",
                warning="No supported schedule title region was detected; no values were guessed.",
            )
            return ExtractionResult()
        tables: list[list[list[str]]] = []
        for region in title_regions:
            table_crop = self._crop_table_area(processed, region, cv2)
            if table_crop is None:
                self._audit(
                    "image_fallback_warning",
                    audit_file,
                    source_page=source_page,
                    method="opencv_grid_detection",
                    warning=f"No table grid detected below {region['title']}.",
                )
                continue
            boxes = self._detect_cell_boxes(table_crop, cv2)
            rows = self._ocr_cells(table_crop, boxes, cv2, pytesseract)
            if rows:
                tables.append(rows)
            self._audit(
                "cell_ocr_table",
                audit_file,
                source_page=source_page,
                method="opencv_grid_detection_and_cell_ocr",
                detected_title=region["title"],
                cell_count=len(boxes),
                row_count=len(rows),
            )
        return self._extract_table_rows(
            tables, "", audit_file, source_page, method="opencv_grid_cell_ocr"
        )

    def _image_dependencies(self, source_file: Path, source_page: int | None):
        try:
            import cv2
            import numpy as np
            import pytesseract
        except ImportError as exc:
            self._audit(
                "image_fallback_unavailable",
                source_file,
                source_page=source_page,
                method="opencv_grid_cell_ocr",
                warning=f"Optional image extraction dependency is missing: {exc.name}.",
            )
            return None
        if not shutil.which("tesseract"):
            self._audit(
                "image_fallback_unavailable",
                source_file,
                source_page=source_page,
                method="opencv_grid_cell_ocr",
                warning="Tesseract OCR executable is not installed; no OCR values were guessed.",
            )
            return None
        return cv2, np, pytesseract

    @staticmethod
    def _preprocess_image(image, cv2):
        upscaled = cv2.resize(image, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
        grayscale = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
        contrast = cv2.convertScaleAbs(grayscale, alpha=1.55, beta=0)
        denoised = cv2.fastNlMeansDenoising(contrast, None, 10, 7, 21)
        _, threshold = cv2.threshold(
            denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        coordinates = cv2.findNonZero(255 - threshold)
        if coordinates is None:
            return threshold
        angle = cv2.minAreaRect(coordinates)[-1]
        angle = -(90 + angle) if angle < -45 else -angle
        if abs(angle) < 0.1 or abs(angle) > 5:
            return threshold
        height, width = threshold.shape[:2]
        matrix = cv2.getRotationMatrix2D((width // 2, height // 2), angle, 1.0)
        return cv2.warpAffine(
            threshold,
            matrix,
            (width, height),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_REPLICATE,
        )

    @staticmethod
    def _detect_image_schedule_regions(image, pytesseract) -> list[dict[str, Any]]:
        data = pytesseract.image_to_data(
            image, config="--psm 11", output_type=pytesseract.Output.DICT
        )
        lines: dict[tuple[int, int, int], list[int]] = {}
        for index, token in enumerate(data["text"]):
            if not _clean(token):
                continue
            key = (data["block_num"][index], data["par_num"][index], data["line_num"][index])
            lines.setdefault(key, []).append(index)
        regions: list[dict[str, Any]] = []
        for indexes in lines.values():
            text = " ".join(_clean(data["text"][index]) for index in indexes)
            for title, pattern in TITLE_PATTERNS.items():
                if not pattern.search(text):
                    continue
                left = min(data["left"][index] for index in indexes)
                top = min(data["top"][index] for index in indexes)
                right = max(data["left"][index] + data["width"][index] for index in indexes)
                bottom = max(data["top"][index] + data["height"][index] for index in indexes)
                regions.append(
                    {"title": title, "left": left, "top": top, "right": right, "bottom": bottom}
                )
        return regions

    @staticmethod
    def _grid_mask(image, cv2):
        inverted = 255 - image
        height, width = image.shape[:2]
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, width // 30), 1))
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, height // 30)))
        horizontal = cv2.morphologyEx(inverted, cv2.MORPH_OPEN, horizontal_kernel)
        vertical = cv2.morphologyEx(inverted, cv2.MORPH_OPEN, vertical_kernel)
        return cv2.add(horizontal, vertical)

    def _crop_table_area(self, image, region: dict[str, Any], cv2):
        height, width = image.shape[:2]
        start_y = max(0, int(region["bottom"]) - 8)
        end_y = min(height, int(height * 0.90))
        candidate = image[start_y:end_y, 0:width]
        mask = self._grid_mask(candidate, cv2)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        rectangles = []
        for contour in contours:
            x, y, box_width, box_height = cv2.boundingRect(contour)
            if box_width >= width * 0.18 and box_height >= candidate.shape[0] * 0.05:
                rectangles.append((box_width * box_height, x, y, box_width, box_height))
        if not rectangles:
            return None
        _, x, y, box_width, box_height = max(rectangles)
        return candidate[y : y + box_height, x : x + box_width]

    def _detect_cell_boxes(self, table_image, cv2) -> list[tuple[int, int, int, int]]:
        mask = self._grid_mask(table_image, cv2)
        contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        height, width = table_image.shape[:2]
        boxes: list[tuple[int, int, int, int]] = []
        for contour in contours:
            x, y, box_width, box_height = cv2.boundingRect(contour)
            if box_width < 18 or box_height < 12:
                continue
            if box_width > width * 0.98 and box_height > height * 0.98:
                continue
            boxes.append((x, y, box_width, box_height))
        boxes.sort(key=lambda box: (box[1], box[0], box[2], box[3]))
        deduped: list[tuple[int, int, int, int]] = []
        for box in boxes:
            if any(abs(box[0] - old[0]) < 3 and abs(box[1] - old[1]) < 3 for old in deduped):
                continue
            deduped.append(box)
        return deduped

    @staticmethod
    def _ocr_cells(table_image, boxes, cv2, pytesseract) -> list[list[str]]:
        if not boxes:
            return []
        median_height = sorted(box[3] for box in boxes)[len(boxes) // 2]
        grouped: list[list[tuple[int, int, int, int]]] = []
        for box in boxes:
            for group in grouped:
                if abs(group[0][1] - box[1]) <= max(5, int(median_height * 0.45)):
                    group.append(box)
                    break
            else:
                grouped.append([box])
        rows: list[list[str]] = []
        for group in grouped:
            values: list[str] = []
            for x, y, width, height in sorted(group, key=lambda box: box[0]):
                cell = table_image[max(0, y + 2) : y + height - 2, max(0, x + 2) : x + width - 2]
                if cell.size == 0:
                    values.append("")
                    continue
                cell = cv2.copyMakeBorder(cell, 8, 8, 8, 8, cv2.BORDER_CONSTANT, value=255)
                values.append(_clean(pytesseract.image_to_string(cell, config="--psm 7")))
            if any(values):
                rows.append(values)
        return rows

    @staticmethod
    def _detect_schedule_titles(text: str) -> list[str]:
        return [title for title, pattern in TITLE_PATTERNS.items() if pattern.search(text)]

    @staticmethod
    def _dedupe_items(items: list[Any], identifier_field: str) -> list[Any]:
        output: list[Any] = []
        seen: set[tuple[Any, ...]] = set()
        for item in items:
            values = item.to_dict()
            key = (
                values.get(identifier_field),
                values.get("width_raw"),
                values.get("height_raw"),
                values.get("level"),
                values.get("description") or values.get("location"),
            )
            if key in seen:
                continue
            seen.add(key)
            output.append(item)
        return output

    @staticmethod
    def _dedupe_schedules(schedules):
        output = []
        seen: set[tuple[Any, ...]] = set()
        for schedule in schedules:
            item_keys = tuple(
                (
                    getattr(item, "tag", None) or getattr(item, "door_number", None),
                    item.width_raw,
                    item.height_raw,
                    getattr(item, "level", None),
                )
                for item in schedule.items
            )
            key = (schedule.schedule_type, schedule.source_page, item_keys)
            if not item_keys or key in seen:
                continue
            seen.add(key)
            output.append(schedule)
        return output

    def _audit(self, event: str, source_file: Path, **details: Any) -> None:
        self.audit.append({"event": event, "source_file": str(source_file), **details})
