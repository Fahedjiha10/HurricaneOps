#!/usr/bin/env python3
from __future__ import annotations

import sys

from bid_pipeline import run_bid_pipeline
from job_intake import IntakeError


def main() -> int:
    print("[HurricaneOps] Starting bid pipeline.")
    try:
        run_bid_pipeline()
    except IntakeError as exc:
        print(f"[HurricaneOps] ERROR: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"[HurricaneOps] ERROR: File operation failed: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
