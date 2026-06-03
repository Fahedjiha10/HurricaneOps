from __future__ import annotations

import csv
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from shutil import copy2

import fitz

from config import (
    EXCEL_EXTENSIONS,
    INBOX_DIR,
    JOB_SUBFOLDERS,
    JOBS_DIR,
    ORGANIZED_PLAN_COMPARTMENTS,
)


class IntakeError(RuntimeError):
    """Raised when the inbox cannot be converted into a job folder."""


SHEET_NUMBER_RE = re.compile(r"^[A-Z]{1,5}[-.]?\d+(?:\.\d+)*$", flags=re.IGNORECASE)
SHEET_NUMBER_PREFIX_RE = re.compile(
    r"^([A-Z]{1,5}[-.]?\d+(?:\.\d+)*)\s+(.+)$", flags=re.IGNORECASE
)
INDEX_SECTION_NAMES = {
    "ARCHITECTURE",
    "CIVIL",
    "COVER",
    "ELECTRICAL",
    "FIRE PROTECTION",
    "LANDSCAPE",
    "LIFE SAFETY",
    "MECHANICAL",
    "PLUMBING",
    "STRUCTURAL",
    "STRUCTURE",
}


def _find_one_file(inbox_dir: Path, extensions: tuple[str, ...], label: str) -> Path:
    matches = sorted(
        path
        for path in inbox_dir.iterdir()
        if path.is_file() and path.suffix.lower() in extensions
    )

    if not matches:
        expected = ", ".join(extensions)
        raise IntakeError(
            f"No {label} found in {inbox_dir}. Add one file with an expected "
            f"extension: {expected}."
        )

    if len(matches) > 1:
        names = ", ".join(path.name for path in matches)
        raise IntakeError(
            f"Expected one {label} in {inbox_dir}, but found {len(matches)}: {names}"
        )

    return matches[0]


def _find_one_plan_set(inbox_dir: Path) -> Path:
    matches = sorted(
        path
        for path in inbox_dir.iterdir()
        if path.is_dir() and any(path.rglob("*.pdf"))
    )

    if not matches:
        raise IntakeError(
            f"No PDF plan set folder found in {inbox_dir}. Add one folder containing "
            "the plan-set PDFs."
        )

    if len(matches) > 1:
        names = ", ".join(path.name for path in matches)
        raise IntakeError(
            f"Expected one PDF plan set folder in {inbox_dir}, but found "
            f"{len(matches)}: {names}"
        )

    return matches[0]


def inspect_inbox(inbox_dir: Path = INBOX_DIR) -> dict[str, object]:
    inbox_dir.mkdir(parents=True, exist_ok=True)
    plan_sets = sorted(
        path
        for path in inbox_dir.iterdir()
        if path.is_dir() and any(path.rglob("*.pdf"))
    )
    excel_files = sorted(
        path
        for path in inbox_dir.iterdir()
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS
    )

    errors: list[str] = []
    if len(plan_sets) != 1:
        errors.append(f"Expected one PDF plan-set folder; found {len(plan_sets)}.")
    if len(excel_files) != 1:
        errors.append(f"Expected one Excel quote template; found {len(excel_files)}.")

    plan_set = plan_sets[0] if len(plan_sets) == 1 else None
    excel_file = excel_files[0] if len(excel_files) == 1 else None
    return {
        "ready": not errors,
        "errors": errors,
        "plan_set": plan_set,
        "pdf_count": len(list(plan_set.rglob("*.pdf"))) if plan_set else 0,
        "excel_template": excel_file,
    }


def sanitize_folder_component(value: str, fallback: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_value = normalized.encode("ascii", "ignore").decode("ascii")
    safe_value = re.sub(r"[^A-Za-z0-9]+", "-", ascii_value).strip("-")
    return safe_value[:60] or fallback


def _normalized_search_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def _clean_line(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("−", "-")).strip()


def _normalize_sheet_number(value: str) -> str:
    return _clean_line(value).upper().replace(" ", "")


def _sheet_number_match_key(value: str) -> str:
    normalized = _normalize_sheet_number(value)
    return re.sub(r"(?<=[A-Z])[-.](?=\d)", "", normalized)


def _indexed_sheet_number(
    value: str, index_entries: dict[str, dict[str, str]]
) -> str:
    match_key = _sheet_number_match_key(value)
    return next(
        (
            indexed_number
            for indexed_number in index_entries
            if _sheet_number_match_key(indexed_number) == match_key
        ),
        "",
    )


def _sheet_number_in_filename(pdf_path: Path) -> str:
    stem = pdf_path.stem.replace("_", " ").replace("−", "-")
    match = re.search(r"\b([A-Z]{1,5}[-.]?\d+(?:\.\d+)*)\b", stem, flags=re.IGNORECASE)
    return _normalize_sheet_number(match.group(1)) if match else ""


def _extract_index_entries(page_texts: list[str]) -> dict[str, dict[str, str]]:
    entries: dict[str, dict[str, str]] = {}
    for text in page_texts:
        normalized_text = _normalized_search_text(text)
        if not any(marker in normalized_text for marker in ("sheet list", "sheet index", "index of drawings")):
            continue
        lines = [_clean_line(line) for line in text.splitlines() if _clean_line(line)]
        discipline = ""
        index = 0
        while index < len(lines):
            line = lines[index]
            upper_line = line.upper()
            if upper_line in INDEX_SECTION_NAMES:
                discipline = upper_line
                index += 1
                continue
            prefix_match = SHEET_NUMBER_PREFIX_RE.fullmatch(line)
            if prefix_match:
                sheet_number = _normalize_sheet_number(prefix_match.group(1))
                title = _clean_line(prefix_match.group(2))
                entries[sheet_number] = {"title": title, "discipline": discipline}
                index += 1
                continue
            if SHEET_NUMBER_RE.fullmatch(line):
                sheet_number = _normalize_sheet_number(line)
                if index + 1 < len(lines):
                    title = lines[index + 1]
                    if (
                        title.upper() not in INDEX_SECTION_NAMES
                        and not SHEET_NUMBER_RE.fullmatch(title)
                    ):
                        entries[sheet_number] = {"title": title, "discipline": discipline}
                        index += 2
                        continue
            index += 1
    return entries


def _infer_sheet_number(
    pdf_path: Path,
    page_index: int,
    page_label: str,
    text: str,
    index_entries: dict[str, dict[str, str]],
) -> str:
    if page_index == 0:
        filename_sheet_number = _sheet_number_in_filename(pdf_path)
        if filename_sheet_number:
            return filename_sheet_number
    normalized_page_label = _normalize_sheet_number(page_label)
    indexed_page_label = _indexed_sheet_number(normalized_page_label, index_entries)
    if indexed_page_label:
        return indexed_page_label
    if SHEET_NUMBER_RE.fullmatch(normalized_page_label):
        return normalized_page_label
    exact_line_numbers = [
        _normalize_sheet_number(line)
        for line in text.splitlines()
        if SHEET_NUMBER_RE.fullmatch(_clean_line(line))
    ]
    index_matches = [
        indexed_number
        for number in exact_line_numbers
        if (indexed_number := _indexed_sheet_number(number, index_entries))
    ]
    if index_matches:
        normalized_text = _normalized_search_text(text)
        is_index_page = any(
            marker in normalized_text
            for marker in ("sheet list", "sheet index", "index of drawings")
        )
        return index_matches[0] if is_index_page else index_matches[-1]
    return exact_line_numbers[0] if exact_line_numbers else ""


def _fallback_sheet_title(pdf_path: Path, page_text: str) -> str:
    normalized_text = _normalized_search_text(page_text)
    for title in (
        "DOOR SCHEDULES",
        "WINDOW SCHEDULES",
        "STOREFRONT SCHEDULES",
        "WIND PRESSURES",
        "WINDOWS & DOORS PRESSURES ELEVATIONS",
        "FLOOR PLAN",
        "ELEVATIONS",
    ):
        if _normalized_search_text(title) in normalized_text:
            return title
    return _clean_line(pdf_path.stem)


def _sheet_compartments(
    pdf_path: Path,
    sheet_number: str,
    sheet_title: str,
    discipline: str,
    page_text: str,
) -> tuple[str, ...]:
    filename_text = _normalized_search_text(pdf_path.stem)
    title_text = _normalized_search_text(sheet_title)
    discipline_text = _normalized_search_text(discipline)
    page_text_normalized = _normalized_search_text(page_text)
    context = " ".join((filename_text, title_text, discipline_text))
    compartments: set[str] = set()
    is_index_page = any(
        marker in page_text_normalized
        for marker in ("sheet list", "sheet index", "index of drawings")
    )
    allow_page_text_schedule_fallback = (
        not is_index_page
        and (not title_text or title_text == filename_text or "schedule" in title_text)
    )

    if "door schedule" in title_text or (
        "door schedule" in page_text_normalized and allow_page_text_schedule_fallback
    ):
        compartments.add("door_schedules")
    if (
        "window schedule" in title_text
        or "glazing schedule" in title_text
        or (
            allow_page_text_schedule_fallback
            and (
                "window schedule" in page_text_normalized
                or "glazing schedule" in page_text_normalized
            )
        )
        or (
            "storefront schedule" in title_text
            and "window" in page_text_normalized
        )
    ):
        compartments.add("window_schedules")
    if (
        "storefront schedule" in title_text
        or (
            allow_page_text_schedule_fallback
            and (
                "storefront schedule" in page_text_normalized
                or (
                    "storefront" in page_text_normalized
                    and "window schedule" in page_text_normalized
                )
            )
        )
    ):
        compartments.add("storefront_schedules")

    is_architectural_level_plan = (
        discipline_text == "architecture"
        and re.search(r"\b(level|roof)\b", title_text)
        and "elevation" not in title_text
    )
    if "floor plan" in context or is_architectural_level_plan:
        compartments.add("floor_plans")
    pressure_page_headings = (
        "wind pressures",
        "wind pressure roof plan",
        "windows doors wind pressures elevations",
        "door pressures",
        "window pressures",
    )
    page_lines = {
        _normalized_search_text(line) for line in page_text.splitlines() if line.strip()
    }
    if "wind pressure" in context or (
        not is_index_page
        and any(heading in page_lines for heading in pressure_page_headings)
    ):
        compartments.add("wind_pressure_elevations")
    elif "elevation" in title_text or "elevation" in filename_text:
        if "pressure" in context or "wind" in context:
            compartments.add("wind_pressure_elevations")
        else:
            compartments.add("architectural_elevations")

    if not compartments:
        compartments.add("other_plan_documents")

    return tuple(
        key
        for key, _folder_name, _label in ORGANIZED_PLAN_COMPARTMENTS
        if key in compartments
    )


def _write_indexed_page(source_pdf: Path, page_index: int, destination: Path) -> None:
    with fitz.open(source_pdf) as source_document, fitz.open() as indexed_document:
        indexed_document.insert_pdf(
            source_document, from_page=page_index, to_page=page_index
        )
        indexed_document.save(destination)


def _indexed_sheet_filename(
    source_pdf: Path, page_number: int, sheet_number: str, sheet_title: str
) -> str:
    label = sanitize_folder_component(
        " ".join(value for value in (sheet_number, sheet_title) if value),
        sanitize_folder_component(source_pdf.stem, "Plan-Sheet"),
    )
    return f"{label}__page-{page_number:03d}.pdf"


def _available_destination(folder: Path, filename: str) -> Path:
    destination = folder / filename
    counter = 2
    while destination.exists():
        destination = folder / f"{Path(filename).stem}_{counter}{Path(filename).suffix}"
        counter += 1
    return destination


def _organize_source_pdfs(job_dir: Path, plan_set_path: Path, pdf_paths: list[Path]) -> None:
    organized_dir = job_dir / "07_Organized_Plan_Set"
    indexed_sheets_dir = organized_dir / "00_Indexed_Sheets"
    indexed_sheets_dir.mkdir()
    folders = {
        key: organized_dir / folder_name
        for key, folder_name, _label in ORGANIZED_PLAN_COMPARTMENTS
    }
    for folder in folders.values():
        folder.mkdir()

    source_pages: list[dict[str, object]] = []
    page_texts: list[str] = []
    for pdf_path in pdf_paths:
        with fitz.open(pdf_path) as document:
            for page_index, page in enumerate(document):
                text = page.get_text()
                page_texts.append(text)
                source_pages.append(
                    {
                        "pdf_path": pdf_path,
                        "page_index": page_index,
                        "page_number": page_index + 1,
                        "page_label": page.get_label(),
                        "text": text,
                    }
                )

    index_entries = _extract_index_entries(page_texts)
    manifest_rows: list[dict[str, str]] = []
    counts = {key: 0 for key in folders}
    for source_page in source_pages:
        pdf_path = Path(source_page["pdf_path"])
        page_index = int(source_page["page_index"])
        page_number = int(source_page["page_number"])
        page_text = str(source_page["text"])
        sheet_number = _infer_sheet_number(
            pdf_path,
            page_index,
            str(source_page["page_label"]),
            page_text,
            index_entries,
        )
        index_entry = index_entries.get(sheet_number, {})
        sheet_title = index_entry.get("title") or _fallback_sheet_title(pdf_path, page_text)
        discipline = index_entry.get("discipline", "")
        compartments = _sheet_compartments(
            pdf_path, sheet_number, sheet_title, discipline, page_text
        )
        indexed_sheet_path = _available_destination(
            indexed_sheets_dir,
            _indexed_sheet_filename(pdf_path, page_number, sheet_number, sheet_title),
        )
        _write_indexed_page(pdf_path, page_index, indexed_sheet_path)
        copied_to: list[str] = []
        for compartment in compartments:
            destination = _available_destination(folders[compartment], indexed_sheet_path.name)
            copy2(indexed_sheet_path, destination)
            counts[compartment] += 1
            copied_to.append(str(destination.relative_to(organized_dir)))
        manifest_rows.append(
            {
                "original_relative_path": str(pdf_path.relative_to(plan_set_path)),
                "source_page_number": str(page_number),
                "sheet_number": sheet_number,
                "sheet_title": sheet_title,
                "discipline": discipline,
                "compartments": ", ".join(compartments),
                "indexed_sheet": str(indexed_sheet_path.relative_to(organized_dir)),
                "organized_copies": ", ".join(copied_to),
            }
        )

    manifest_path = organized_dir / "source_file_manifest.csv"
    with manifest_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=(
                "original_relative_path",
                "source_page_number",
                "sheet_number",
                "sheet_title",
                "discipline",
                "compartments",
                "indexed_sheet",
                "organized_copies",
            ),
        )
        writer.writeheader()
        writer.writerows(manifest_rows)
    copy2(manifest_path, organized_dir / "sheet_index.csv")

    for key, _folder_name, label in ORGANIZED_PLAN_COMPARTMENTS:
        print(f"[HurricaneOps] Organized {label.lower()}: {counts[key]}")
    print(f"[HurricaneOps] Indexed plan sheets: {len(manifest_rows)}")
    print(f"[HurricaneOps] Index titles discovered: {len(index_entries)}")
    print(f"[HurricaneOps] Created source file manifest: {manifest_path}")


def _new_job_folder(
    jobs_dir: Path,
    timestamp: datetime,
    project_name: str = "",
    address: str = "",
) -> Path:
    if project_name:
        project_part = sanitize_folder_component(project_name, "Untitled-Project")
        base_name = f"{timestamp:%Y-%m-%d}_{project_part}"
        if address:
            address_part = sanitize_folder_component(address, "Address-Not-Provided")
            base_name = f"{base_name}_{address_part}"
    else:
        base_name = timestamp.strftime("%Y-%m-%d_%H%M%S")
    candidate = jobs_dir / base_name
    counter = 2

    while candidate.exists():
        candidate = jobs_dir / f"{base_name}_{counter}"
        counter += 1

    return candidate


def _write_checklist(checklist_path: Path, plan_set_name: str, excel_name: str) -> None:
    checklist_path.write_text(
        "\n".join(
            (
                "# HurricaneOps Job Checklist",
                "",
                f"- [x] PDF plan set copied: `{plan_set_name}`",
                f"- [x] Quote workbook template copied: `{excel_name}`",
                "- [x] Plan PDFs compartmentalized for estimator review",
                "- [ ] Extract relevant PDF pages",
                "- [ ] Build schedules",
                "- [ ] Complete takeoff QA",
                "- [ ] Update quote workbook",
                "- [ ] Complete NOA / PSF check",
                "- [ ] Draft proposal email",
                "",
            )
        ),
        encoding="utf-8",
    )


def create_job_from_inbox(
    inbox_dir: Path = INBOX_DIR,
    jobs_dir: Path = JOBS_DIR,
    timestamp: datetime | None = None,
    project_name: str = "",
    address: str = "",
) -> tuple[Path, Path, Path]:
    print(f"[HurricaneOps] Checking inbox: {inbox_dir}")
    inbox_dir.mkdir(parents=True, exist_ok=True)
    jobs_dir.mkdir(parents=True, exist_ok=True)

    plan_set_path = _find_one_plan_set(inbox_dir)
    excel_path = _find_one_file(inbox_dir, EXCEL_EXTENSIONS, "Excel file")
    pdf_paths = sorted(plan_set_path.rglob("*.pdf"))
    print(f"[HurricaneOps] Found PDF plan set: {plan_set_path.name}")
    print(f"[HurricaneOps] Found {len(pdf_paths)} source PDF files in plan set.")
    print(f"[HurricaneOps] Found Excel template: {excel_path.name}")

    job_dir = _new_job_folder(
        jobs_dir, timestamp or datetime.now(), project_name=project_name, address=address
    )
    job_dir.mkdir()
    print(f"[HurricaneOps] Created job folder: {job_dir}")

    for subfolder in JOB_SUBFOLDERS:
        (job_dir / subfolder).mkdir()
        print(f"[HurricaneOps] Created subfolder: {subfolder}")

    pdf_destination = job_dir / "00_Source_PDFs" / plan_set_path.name
    excel_destination = job_dir / "04_Quote_Workbook" / excel_path.name
    for pdf_path in pdf_paths:
        relative_path = pdf_path.relative_to(plan_set_path)
        destination = pdf_destination / relative_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        copy2(pdf_path, destination)
    copy2(excel_path, excel_destination)
    _organize_source_pdfs(job_dir, plan_set_path, pdf_paths)
    print(f"[HurricaneOps] Copied PDF plan set to: {pdf_destination}")
    print(f"[HurricaneOps] Copied Excel template to: {excel_destination}")

    checklist_path = job_dir / "checklist.md"
    _write_checklist(checklist_path, plan_set_path.name, excel_path.name)
    print(f"[HurricaneOps] Created checklist: {checklist_path}")
    print("[HurricaneOps] Job intake complete.")
    return job_dir, pdf_destination, excel_destination
