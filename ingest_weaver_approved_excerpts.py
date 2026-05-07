#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excerpt_library import DEFAULT_DB_PATH, ingest_weaver_approved_records


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ingest approved Weaver excerpts into the local excerpt library."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        type=Path,
        help="Optional JSON file containing Weaver-approved records. If omitted, reads JSON from stdin.",
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--source-name",
        default="weaver_approved",
        help="Logical source name stored in excerpt_sources",
    )
    return parser.parse_args()


def load_records(input_path: Path | None) -> list[dict]:
    if input_path:
        payload = json.loads(input_path.read_text(encoding="utf-8"))
    else:
        payload = json.load(sys.stdin)

    if isinstance(payload, dict):
        records = payload.get("records", [])
    else:
        records = payload

    if not isinstance(records, list):
        raise ValueError("Expected a JSON list or an object with a 'records' array.")
    return records


def main() -> int:
    args = parse_args()
    records = load_records(args.input_path)
    result = ingest_weaver_approved_records(
        records,
        db_path=args.db_path,
        source_name=args.source_name,
    )
    result["record_count"] = len(records)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
