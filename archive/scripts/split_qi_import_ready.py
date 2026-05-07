#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

DEFAULT_INPUT = Path("data/poetry_please_qi_all_drive.json")
DEFAULT_READY_OUTPUT = Path("data/poetry_please_qi_all_drive_ready.json")
DEFAULT_REVIEW_OUTPUT = Path("data/poetry_please_qi_all_drive_followup_review.json")


def has_value(value: object) -> bool:
    return isinstance(value, str) and value.strip() != ""


def main() -> int:
    parser = argparse.ArgumentParser(description="Split Poetry Please QI JSON into ready and review sets.")
    parser.add_argument("--input", default=str(DEFAULT_INPUT))
    parser.add_argument("--ready-output", default=str(DEFAULT_READY_OUTPUT))
    parser.add_argument("--review-output", default=str(DEFAULT_REVIEW_OUTPUT))
    args = parser.parse_args()

    input_path = Path(args.input)
    rows = json.loads(input_path.read_text(encoding="utf-8"))

    ready: list[dict[str, object]] = []
    review: list[dict[str, object]] = []

    for row in rows:
        if has_value(row.get("bookLink")) and has_value(row.get("releaseCatalog")):
            ready.append(row)
        else:
            review.append(row)

    ready_output = Path(args.ready_output)
    review_output = Path(args.review_output)
    ready_output.parent.mkdir(parents=True, exist_ok=True)
    ready_output.write_text(json.dumps(ready, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    review_output.parent.mkdir(parents=True, exist_ok=True)
    review_output.write_text(json.dumps(review, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(
        json.dumps(
            {
                "input": str(input_path),
                "ready": len(ready),
                "review": len(review),
                "ready_output": str(ready_output),
                "review_output": str(review_output),
            }
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
