#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from bid_pipeline import run_bid_pipeline
from config import (
    EXCEL_EXTENSIONS,
    ICLOUD_AUTO_JOBS_DIR,
    ICLOUD_AUTOMATION_DIR,
    ICLOUD_INCOMING_DIR,
    ICLOUD_PROCESSED_DIR,
    ICLOUD_SYNC_LOCK_PATH,
    ICLOUD_SYNC_STATE_PATH,
    ICLOUD_WORK_DIR,
)
from job_intake import IntakeError


def _load_state() -> dict[str, object]:
    if not ICLOUD_SYNC_STATE_PATH.exists():
        return {"packages": {}}
    return json.loads(ICLOUD_SYNC_STATE_PATH.read_text(encoding="utf-8"))


def _save_state(state: dict[str, object]) -> None:
    ICLOUD_SYNC_STATE_PATH.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")


def _package_fingerprint(package_dir: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(item for item in package_dir.rglob("*") if item.is_file()):
        stat = path.stat()
        digest.update(str(path.relative_to(package_dir)).encode("utf-8"))
        digest.update(str(stat.st_size).encode("ascii"))
        digest.update(str(stat.st_mtime_ns).encode("ascii"))
    return digest.hexdigest()


def _package_is_stable(package_dir: Path, stable_seconds: int) -> bool:
    files = [path for path in package_dir.rglob("*") if path.is_file()]
    if not files:
        return False
    newest_mtime = max(path.stat().st_mtime for path in files)
    return time.time() - newest_mtime >= stable_seconds


def _default_template() -> Path:
    candidates = sorted(
        path
        for path in ICLOUD_WORK_DIR.rglob("*.xlsx")
        if path.parent.name.strip().lower() == "quote template"
        and not path.name.startswith("~$")
    )
    if len(candidates) != 1:
        raise IntakeError(
            "Expected one default quote template in the iCloud Quote Template folder; "
            f"found {len(candidates)}."
        )
    return candidates[0]


def _package_template(package_dir: Path) -> Path:
    candidates = sorted(
        path
        for path in package_dir.rglob("*")
        if path.is_file()
        and path.suffix.lower() in EXCEL_EXTENSIONS
        and not path.name.startswith("~$")
    )
    if len(candidates) > 1:
        names = ", ".join(path.name for path in candidates)
        raise IntakeError(
            f"Expected zero or one Excel template in {package_dir.name}; found "
            f"{len(candidates)}: {names}"
        )
    return candidates[0] if candidates else _default_template()


def _package_metadata(package_dir: Path) -> dict[str, str]:
    metadata_path = package_dir / "project.json"
    metadata = (
        json.loads(metadata_path.read_text(encoding="utf-8"))
        if metadata_path.exists()
        else {}
    )
    return {
        "project_name": str(metadata.get("project_name") or package_dir.name),
        "address": str(metadata.get("address") or ""),
        "notes": str(metadata.get("notes") or "Automatically ingested from iCloud Drive."),
    }


def _stage_package(package_dir: Path) -> tempfile.TemporaryDirectory[str]:
    temp_dir = tempfile.TemporaryDirectory(prefix="hurricaneops_icloud_")
    inbox_dir = Path(temp_dir.name) / "00_INBOX"
    plan_set_dir = inbox_dir / "Uploaded_Plan_Set"
    plan_set_dir.mkdir(parents=True)

    pdf_paths = sorted(package_dir.rglob("*.pdf"))
    if not pdf_paths:
        temp_dir.cleanup()
        raise IntakeError(f"No PDF files found in iCloud package {package_dir.name}.")

    for pdf_path in pdf_paths:
        destination = plan_set_dir / pdf_path.relative_to(package_dir)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf_path, destination)

    template_path = _package_template(package_dir)
    shutil.copy2(template_path, inbox_dir / template_path.name)
    return temp_dir


def _publish_results(job_dir: Path) -> Path:
    destination = ICLOUD_PROCESSED_DIR / job_dir.name
    destination.mkdir(parents=True, exist_ok=True)
    for name in ("pipeline_summary.json", "full_job_package.zip"):
        shutil.copy2(job_dir / name, destination / name)
    organized_plan_set = job_dir / "07_Organized_Plan_Set"
    if organized_plan_set.exists():
        shutil.copytree(
            organized_plan_set,
            destination / organized_plan_set.name,
            dirs_exist_ok=True,
        )
    return destination


@contextmanager
def _worker_lock():
    ICLOUD_AUTOMATION_DIR.mkdir(parents=True, exist_ok=True)
    if ICLOUD_SYNC_LOCK_PATH.exists():
        lock_age = time.time() - ICLOUD_SYNC_LOCK_PATH.stat().st_mtime
        if lock_age > 3600:
            ICLOUD_SYNC_LOCK_PATH.unlink()
    try:
        descriptor = os.open(ICLOUD_SYNC_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        print("[HurricaneOps iCloud] Another ingest worker is already running.")
        yield False
        return
    try:
        os.write(descriptor, str(os.getpid()).encode("ascii"))
        os.close(descriptor)
        yield True
    finally:
        ICLOUD_SYNC_LOCK_PATH.unlink(missing_ok=True)


def run_once(
    stable_seconds: int = 60,
    package_name: str | None = None,
    force: bool = False,
    dry_run: bool = False,
) -> int:
    ICLOUD_AUTOMATION_DIR.mkdir(parents=True, exist_ok=True)
    ICLOUD_INCOMING_DIR.mkdir(parents=True, exist_ok=True)
    ICLOUD_PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    state = _load_state()
    package_state = state.setdefault("packages", {})

    packages = sorted(path for path in ICLOUD_INCOMING_DIR.iterdir() if path.is_dir())
    if package_name:
        packages = [path for path in packages if path.name == package_name]

    if not packages:
        print(f"[HurricaneOps iCloud] No incoming project folders in {ICLOUD_INCOMING_DIR}.")
        return 0

    failures = 0
    for package_dir in packages:
        source_key = str(package_dir.resolve())
        fingerprint = _package_fingerprint(package_dir)
        previous = package_state.get(source_key, {})
        if not force and previous.get("fingerprint") == fingerprint:
            print(f"[HurricaneOps iCloud] Unchanged package skipped: {package_dir.name}")
            continue
        if not force and not _package_is_stable(package_dir, stable_seconds):
            print(f"[HurricaneOps iCloud] Waiting for files to settle: {package_dir.name}")
            continue
        if dry_run:
            print(f"[HurricaneOps iCloud] Ready to process: {package_dir.name}")
            continue

        print(f"[HurricaneOps iCloud] Processing: {package_dir.name}")
        temp_dir = None
        try:
            metadata = _package_metadata(package_dir)
            temp_dir = _stage_package(package_dir)
            job_dir, issues = run_bid_pipeline(
                inbox_dir=Path(temp_dir.name) / "00_INBOX",
                jobs_dir=ICLOUD_AUTO_JOBS_DIR,
                project_name=metadata["project_name"],
                address=metadata["address"],
                notes=metadata["notes"],
            )
            published_dir = _publish_results(job_dir)
        except (IntakeError, OSError, ValueError) as exc:
            failures += 1
            package_state[source_key] = {
                "source": source_key,
                "fingerprint": fingerprint,
                "status": "ERROR",
                "updated_at": datetime.now().isoformat(timespec="seconds"),
                "error": str(exc),
            }
            print(f"[HurricaneOps iCloud] ERROR {package_dir.name}: {exc}", file=sys.stderr)
        else:
            package_state[source_key] = {
                "source": source_key,
                "fingerprint": fingerprint,
                "status": "PROCESSED",
                "updated_at": datetime.now().isoformat(timespec="seconds"),
                "job_dir": str(job_dir.resolve()),
                "published_dir": str(published_dir.resolve()),
                "critical_count": len(issues),
            }
            print(f"[HurricaneOps iCloud] Completed: {job_dir.name}")
        finally:
            if temp_dir is not None:
                temp_dir.cleanup()
        _save_state(state)
    return 1 if failures else 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Process new HurricaneOps iCloud queue packages.")
    parser.add_argument("--once", action="store_true", help="Process the queue once and exit.")
    parser.add_argument("--package", help="Process only one incoming folder name.")
    parser.add_argument("--stable-seconds", type=int, default=60)
    parser.add_argument("--force", action="store_true", help="Reprocess unchanged packages.")
    parser.add_argument("--dry-run", action="store_true", help="List ready packages without running.")
    args = parser.parse_args()
    with _worker_lock() as acquired:
        if not acquired:
            return 0
        return run_once(
            stable_seconds=args.stable_seconds,
            package_name=args.package,
            force=args.force,
            dry_run=args.dry_run,
        )


if __name__ == "__main__":
    raise SystemExit(main())
