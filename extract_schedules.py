from __future__ import annotations

import argparse
import sys
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parent / "05_SCRIPTS"
sys.path.insert(0, str(SCRIPTS_DIR))

from export_quote import write_extraction_outputs  # noqa: E402
from schedule_extractor import ScheduleExtractor  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract construction schedules into validated JSON and audit files."
    )
    parser.add_argument("input", type=Path, nargs="+", help="PDF, PNG, JPG, or JPEG input file.")
    parser.add_argument("--project-address", default="", help="Project address stored in JSON output.")
    parser.add_argument("--output", type=Path, required=True, help="Output folder.")
    arguments = parser.parse_args()

    extractor = ScheduleExtractor(project_address=arguments.project_address)
    result = extractor.extract_files(arguments.input)
    outputs = write_extraction_outputs(result, arguments.output)
    item_count = sum(len(schedule.items) for schedule in result.schedules)
    print(f"Extracted {item_count} structured schedule rows.")
    for label, path in outputs.items():
        print(f"{label}: {path.resolve()}")
    if not item_count:
        print("No structured rows were extracted. Review extraction_audit.json.", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
